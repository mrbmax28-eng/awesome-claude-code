class_name RunWorld
extends Node3D
## Une run : biome + véhicule + vagues + ramassage. Émet run_finished à la fin
## (victoire = toutes les vagues vaincues, défaite = PV à zéro, ou abandon).

signal run_finished(run: RunState)

const BIOME_SCENES := {
	"desert": "res://scenes/biomes/desert.tscn",
}
const HUD_SCENE := "res://scenes/ui/hud.tscn"

var run: RunState
var biome: BiomeBase
var player: BaseVehicle
var camera_rig: CameraRig
var spawner: WaveSpawner
var hud: Hud
var entities: Node3D

var _cam_cfg: Dictionary
var _run_cfg: Dictionary
var _ending := false
var _coin_rng := RandomNumberGenerator.new()


func setup(p_run: RunState) -> void:
	run = p_run


func _ready() -> void:
	_cam_cfg = DataDB.config()["camera"]
	_run_cfg = DataDB.config()["run"]
	_coin_rng.seed = run.run_seed ^ 0xc011

	biome = (load(BIOME_SCENES[run.biome_id]) as PackedScene).instantiate()
	biome.setup(run.biome_id, run.run_seed)
	add_child(biome)

	entities = Node3D.new()
	entities.name = "Entities"
	add_child(entities)

	var vdef := DataDB.vehicle(run.vehicle_id)
	player = (load(vdef["scene"]) as PackedScene).instantiate()
	player.setup(run.vehicle_id)
	# Adaptation au biome : vitesse du véhicule (ex. char ralenti dans le désert).
	player.stats = player.stats.duplicate()
	player.stats["max_speed"] = float(player.stats["max_speed"]) * biome.adaptation(run.vehicle_id, "speed_mult")
	entities.add_child(player)
	player.died.connect(_on_player_died)
	player.impact.connect(func(strength: float) -> void: camera_rig.add_trauma(float(_cam_cfg["shake_on_hit"]) * strength))
	player.hp_changed.connect(_on_player_hp_changed)

	camera_rig = CameraRig.new()
	add_child(camera_rig)
	camera_rig.target = player
	camera_rig.snap_to_target()

	spawner = WaveSpawner.new()
	add_child(spawner)
	spawner.enemy_killed.connect(_on_enemy_killed)
	spawner.all_waves_cleared.connect(_on_waves_cleared)
	spawner.setup(biome, player, entities, run.run_seed, {
		"tier": int(vdef["tier"]),
		"biome_index": int(biome.def["order"]),
		"level": run.level,
		"adaptive_mult": run.difficulty_mult,
	})
	run.enemies_total = spawner.total_enemies()

	hud = (load(HUD_SCENE) as PackedScene).instantiate()
	add_child(hud)
	hud.bind(self)
	hud.abandon_requested.connect(_on_abandon)


func _physics_process(delta: float) -> void:
	if run.result == RunState.Result.NONE:
		run.elapsed_s += delta


func _on_enemy_killed(enemy: EnemyBase) -> void:
	run.kills += 1
	camera_rig.add_trauma(float(_cam_cfg["shake_on_kill"]))
	var drop: Array = enemy.def["coin_drop"]
	var count := _coin_rng.randi_range(int(drop[0]), int(drop[1]))
	count = int(round(count * biome.adaptation(run.vehicle_id, "resource_mult")))
	var scatter: float = DataDB.config()["pickups"]["coin_scatter_speed"]
	for i in count:
		var coin := Coin.new()
		var dir := Vector2.from_angle(_coin_rng.randf() * TAU)
		coin.setup(player, dir * scatter * _coin_rng.randf_range(0.4, 1.0))
		coin.position = enemy.global_position
		coin.collected.connect(_on_coin_collected)
		entities.add_child(coin)
	hud.refresh()


func _on_coin_collected(value: int) -> void:
	run.coins += value
	hud.refresh()


var _last_hp := -1.0


func _on_player_hp_changed(hp: float, _max_hp: float) -> void:
	if _last_hp >= 0.0 and hp < _last_hp:
		camera_rig.add_trauma(float(_cam_cfg["shake_on_damage"]))
	_last_hp = hp


func _on_player_died() -> void:
	_finish(RunState.Result.DEFEAT, float(_run_cfg["defeat_end_delay_s"]))


func _on_waves_cleared() -> void:
	_finish(RunState.Result.VICTORY, float(_run_cfg["victory_end_delay_s"]))


func _on_abandon() -> void:
	_finish(RunState.Result.ABANDONED, 0.0)


func _finish(result: RunState.Result, delay: float) -> void:
	if _ending:
		return
	_ending = true
	run.result = result
	player.controls_enabled = false
	if result == RunState.Result.VICTORY:
		# Les pièces encore au sol sont aspirées vers le joueur pendant le délai.
		for node in entities.get_children():
			if node is Coin:
				(node as Coin).attract()
	if delay > 0.0:
		await get_tree().create_timer(delay, false).timeout
	run_finished.emit(run)
