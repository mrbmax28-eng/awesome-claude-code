extends Node
## Singleton (autoload « GameState ») — SEUL autoload du projet.
## Rôle : détenir le profil persistant et les réglages, gérer la sauvegarde
## (auto toutes les N s + à la fermeture), appliquer le résultat des runs.
## Il ne contient aucune logique de gameplay en run.

signal profile_changed

var data: Dictionary = {}
var save_system: SaveSystem
var _autosave_timer: Timer


func _ready() -> void:
	process_mode = Node.PROCESS_MODE_ALWAYS
	InputSetup.apply(DataDB.get_file("input"))
	var save_cfg: Dictionary = DataDB.config()["save"]
	save_system = SaveSystem.new(_save_path(save_cfg), int(save_cfg["version"]))
	data = save_system.load_data()
	apply_settings()
	get_tree().set_auto_accept_quit(false)
	_autosave_timer = Timer.new()
	_autosave_timer.wait_time = float(save_cfg["autosave_interval_s"])
	_autosave_timer.autostart = true
	_autosave_timer.timeout.connect(save)
	add_child(_autosave_timer)


func _notification(what: int) -> void:
	if what == NOTIFICATION_WM_CLOSE_REQUEST:
		quit_game()


## Chemin surchargeable par la ligne de commande (-- --save-path=...), utile pour les tests.
func _save_path(save_cfg: Dictionary) -> String:
	var override := CmdArgs.get_string("save-path", "")
	return override if not override.is_empty() else String(save_cfg["path"])


var profile: Dictionary:
	get:
		return data["profile"]

var settings: Dictionary:
	get:
		return data["settings"]


func save() -> bool:
	return save_system.save(data)


func quit_game() -> void:
	save()
	get_tree().quit()


func set_setting(key: String, value: Variant) -> void:
	settings[key] = value
	apply_settings()
	save()


func apply_settings() -> void:
	if DisplayServer.get_name() == "headless":
		return
	var mode := DisplayServer.WINDOW_MODE_FULLSCREEN if settings["fullscreen"] else DisplayServer.WINDOW_MODE_WINDOWED
	if DisplayServer.window_get_mode() != mode:
		DisplayServer.window_set_mode(mode)


func death_streak(biome_id: String) -> int:
	return int((profile["death_streak"] as Dictionary).get(biome_id, 0))


## Prépare une nouvelle run (graine, multiplicateur adaptatif).
func create_run(biome_id: String, vehicle_id: String, run_seed: int) -> RunState:
	var run := RunState.new(run_seed, biome_id, vehicle_id)
	var cfg: Dictionary = DataDB.difficulty()["adaptive"]
	run.difficulty_mult = Difficulty.adaptive_multiplier(
			death_streak(biome_id), cfg, bool(settings["adaptive_difficulty"]))
	return run


## Applique le résultat d'une run terminée au profil et sauvegarde.
func apply_run_result(run: RunState) -> void:
	var conv: Dictionary = DataDB.economy()["run_conversion"]
	run.scrap_gained = Economy.run_conversion(run.coins, run.is_victory(), conv)
	profile["scrap"] = int(profile["scrap"]) + run.scrap_gained
	profile["runs_played"] = int(profile["runs_played"]) + 1
	profile["total_coins"] = int(profile["total_coins"]) + run.coins
	profile["best_coins"] = maxi(int(profile["best_coins"]), run.coins)
	profile["total_kills"] = int(profile["total_kills"]) + run.kills
	var streaks: Dictionary = profile["death_streak"]
	if run.is_victory():
		profile["runs_won"] = int(profile["runs_won"]) + 1
		streaks[run.biome_id] = 0
	elif run.result == RunState.Result.DEFEAT:
		streaks[run.biome_id] = death_streak(run.biome_id) + 1
	save()
	profile_changed.emit()
