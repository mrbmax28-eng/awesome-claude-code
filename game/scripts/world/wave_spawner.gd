class_name WaveSpawner
extends Node
## Fait apparaître les vagues d'un biome (data/biomes.json > waves) avec une graine reproductible.

signal enemy_spawned(enemy: EnemyBase)
signal enemy_killed(enemy: EnemyBase)
signal all_waves_cleared

const ENEMY_SCENES := {
	"scarab": "res://scenes/enemies/scarab.tscn",
}

var waves: Array = []
var biome: BiomeBase
var player: BaseVehicle
var container: Node3D
var stat_args: Dictionary
var rng := RandomNumberGenerator.new()

var _wave_index := -1
var _to_spawn := 0
var _alive := 0
var _timer := 0.0
var _finished := false


## stat_args : { tier, biome_index, level, adaptive_mult } transmis à Difficulty.
func setup(p_biome: BiomeBase, p_player: BaseVehicle, p_container: Node3D, run_seed: int, p_stat_args: Dictionary) -> void:
	biome = p_biome
	player = p_player
	container = p_container
	waves = biome.def["waves"]
	stat_args = p_stat_args
	rng.seed = run_seed ^ 0x5eed
	_next_wave()


func total_enemies() -> int:
	var total := 0
	for w: Dictionary in waves:
		total += int(w["count"])
	return total


func _next_wave() -> void:
	_wave_index += 1
	if _wave_index >= waves.size():
		_finished = true
		all_waves_cleared.emit()
		return
	_to_spawn = int(waves[_wave_index]["count"])
	_timer = float(waves[_wave_index].get("first_spawn_delay_s", 0.0))


func _physics_process(delta: float) -> void:
	if _finished or player == null:
		return
	var wave: Dictionary = waves[_wave_index]
	_timer -= delta
	if _to_spawn > 0 and _timer <= 0.0 and _alive < int(wave["max_alive"]):
		_spawn(wave)
		_timer = float(wave["spawn_interval_s"])
	if _to_spawn == 0 and _alive == 0:
		_next_wave()


func _spawn(wave: Dictionary) -> void:
	var id: String = wave["enemy"]
	var enemy: EnemyBase = (load(ENEMY_SCENES[id]) as PackedScene).instantiate()
	var stats := Difficulty.enemy_stats(DataDB.enemy(id), DataDB.difficulty(),
			stat_args["tier"], stat_args["biome_index"], stat_args["level"], stat_args["adaptive_mult"])
	enemy.setup(id, stats, player)
	var p := LevelGenerator.spawn_point(rng, biome.half_size,
			Vector2(player.global_position.x, player.global_position.z), float(wave["spawn_min_dist"]), biome.decor)
	enemy.position = Vector3(p.x, 0.0, p.y)
	enemy.killed.connect(_on_enemy_killed)
	container.add_child(enemy)
	_to_spawn -= 1
	_alive += 1
	enemy_spawned.emit(enemy)


func _on_enemy_killed(enemy: EnemyBase) -> void:
	_alive -= 1
	enemy_killed.emit(enemy)
