extends Node
## Point d'entrée : enchaîne menu -> run -> écran de fin.
## Arguments (après « -- ») : --seed=N  --autoplay  --quit-after-run  --time-scale=X
##   --save-path=chemin  --screenshot-dir=chemin (captures périodiques)  --quit-after-s=N

const RUN_WORLD_SCENE := "res://scenes/world/run_world.tscn"
const MENU_SCENE := "res://scenes/ui/main_menu.tscn"
const END_SCENE := "res://scenes/ui/end_screen.tscn"

var _current: Node
var _last_seed := -1


func _ready() -> void:
	Engine.time_scale = CmdArgs.get_float("time-scale", 1.0)
	if CmdArgs.has("screenshot-dir"):
		var shooter := ScreenshotTaker.new()
		shooter.out_dir = CmdArgs.get_string("screenshot-dir", "")
		shooter.interval_s = CmdArgs.get_float("screenshot-interval", 4.0)
		add_child(shooter)
	var quit_after := CmdArgs.get_float("quit-after-s", 0.0)
	if quit_after > 0.0:
		get_tree().create_timer(quit_after, true, false, true).timeout.connect(GameState.quit_game)
	if CmdArgs.has("autoplay"):
		start_run.call_deferred(CmdArgs.get_int("seed", -1))
	else:
		show_menu()


func _swap(node: Node) -> void:
	if _current:
		_current.queue_free()
	_current = node
	add_child(node)


func show_menu() -> void:
	var menu: MainMenu = (load(MENU_SCENE) as PackedScene).instantiate()
	menu.start_run_requested.connect(start_run.bind(-1))
	menu.quit_requested.connect(GameState.quit_game)
	_swap(menu)


func start_run(run_seed: int = -1) -> void:
	if run_seed < 0:
		run_seed = CmdArgs.get_int("seed", -1)
	if run_seed < 0:
		run_seed = randi() & 0x7fffffff
	_last_seed = run_seed
	var run_cfg: Dictionary = DataDB.config()["run"]
	var run := GameState.create_run(run_cfg["default_biome"], run_cfg["default_vehicle"], run_seed)
	var world: RunWorld = (load(RUN_WORLD_SCENE) as PackedScene).instantiate()
	world.setup(run)
	world.run_finished.connect(_on_run_finished)
	_swap(world)
	if CmdArgs.has("autoplay"):
		var pilot := Autopilot.new()
		pilot.world = world
		world.add_child(pilot)
	print("[run] start seed=%d adaptive_mult=%.2f" % [run_seed, run.difficulty_mult])


func _on_run_finished(run: RunState) -> void:
	GameState.apply_run_result(run)
	print("[run] end ", JSON.stringify(run.summary()))
	if CmdArgs.has("quit-after-run"):
		GameState.quit_game()
		return
	var end: EndScreen = (load(END_SCENE) as PackedScene).instantiate()
	end.retry_requested.connect(start_run.bind(-1))
	end.menu_requested.connect(show_menu)
	_swap(end)
	end.show_run(run)
