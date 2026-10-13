class_name Hud
extends CanvasLayer
## HUD minimal : PV, jauge d'énergie du véhicule (turbo pour la voiture), pièces,
## progression de la vague. Menu pause (Échap / Start).

signal abandon_requested

var world: RunWorld
var _hp_bar: ProgressBar
var _energy_bar: ProgressBar
var _energy_label: Label
var _coins_label: Label
var _wave_label: Label
var _seed_label: Label
var _pause_panel: Control
var _pause_buttons: Array[Button] = []


func _ready() -> void:
	process_mode = Node.PROCESS_MODE_ALWAYS
	layer = 10
	var root := MarginContainer.new()
	root.set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	for side: String in ["left", "right", "top", "bottom"]:
		root.add_theme_constant_override("margin_" + side, 24)
	root.mouse_filter = Control.MOUSE_FILTER_IGNORE
	add_child(root)

	var top_left := VBoxContainer.new()
	top_left.size_flags_horizontal = Control.SIZE_SHRINK_BEGIN
	top_left.size_flags_vertical = Control.SIZE_SHRINK_BEGIN
	top_left.add_theme_constant_override("separation", 6)
	root.add_child(top_left)
	top_left.add_child(UiStyle.label("PV", 16, UiStyle.MUTED))
	_hp_bar = UiStyle.bar(UiStyle.HP, Vector2(320, 22))
	top_left.add_child(_hp_bar)
	_energy_label = UiStyle.label("TURBO", 16, UiStyle.MUTED)
	top_left.add_child(_energy_label)
	_energy_bar = UiStyle.bar(UiStyle.TURBO, Vector2(320, 16))
	top_left.add_child(_energy_bar)

	var top_right := VBoxContainer.new()
	top_right.size_flags_horizontal = Control.SIZE_SHRINK_END
	top_right.size_flags_vertical = Control.SIZE_SHRINK_BEGIN
	top_right.alignment = BoxContainer.ALIGNMENT_END
	root.add_child(top_right)
	_coins_label = UiStyle.label("", 30, UiStyle.ACCENT)
	_coins_label.horizontal_alignment = HORIZONTAL_ALIGNMENT_RIGHT
	top_right.add_child(_coins_label)
	_wave_label = UiStyle.label("", 18)
	_wave_label.horizontal_alignment = HORIZONTAL_ALIGNMENT_RIGHT
	top_right.add_child(_wave_label)

	_seed_label = UiStyle.label("", 12, UiStyle.MUTED)
	_seed_label.size_flags_vertical = Control.SIZE_SHRINK_END
	_seed_label.size_flags_horizontal = Control.SIZE_SHRINK_BEGIN
	root.add_child(_seed_label)

	_build_pause_menu()


func bind(p_world: RunWorld) -> void:
	world = p_world
	world.player.hp_changed.connect(func(hp: float, max_hp: float) -> void:
		_hp_bar.max_value = max_hp
		_hp_bar.value = hp)
	world.player.energy_changed.connect(_on_energy_changed)
	_hp_bar.max_value = world.player.max_hp
	_hp_bar.value = world.player.hp
	_energy_bar.max_value = 1.0
	_energy_bar.step = 0.001
	_on_energy_changed(0.0, false)
	_seed_label.text = "graine %d" % world.run.run_seed
	refresh()


func refresh() -> void:
	_coins_label.text = "%d pièces" % world.run.coins
	_wave_label.text = "Vague : %d / %d" % [world.run.kills, world.run.enemies_total]


func _on_energy_changed(ratio: float, active: bool) -> void:
	_energy_bar.value = ratio
	var fill := _energy_bar.get_theme_stylebox("fill") as StyleBoxFlat
	fill.bg_color = UiStyle.TURBO_ACTIVE if active else UiStyle.TURBO
	_energy_label.text = "TURBO !" if active else "TURBO"


func _unhandled_input(event: InputEvent) -> void:
	if event.is_action_pressed("pause"):
		_set_paused(not get_tree().paused)
		get_viewport().set_input_as_handled()


func _set_paused(value: bool) -> void:
	if world and world.run.result != RunState.Result.NONE:
		return
	get_tree().paused = value
	_pause_panel.visible = value
	if value:
		UiStyle.focus_first(_pause_buttons)


func _build_pause_menu() -> void:
	_pause_panel = ColorRect.new()
	(_pause_panel as ColorRect).color = Color(0, 0, 0, 0.55)
	_pause_panel.set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	_pause_panel.visible = false
	add_child(_pause_panel)
	var center := CenterContainer.new()
	center.set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	_pause_panel.add_child(center)
	var panel := UiStyle.panel()
	center.add_child(panel)
	var box := VBoxContainer.new()
	box.add_theme_constant_override("separation", 14)
	panel.add_child(box)
	var title := UiStyle.label("Pause", 36)
	title.horizontal_alignment = HORIZONTAL_ALIGNMENT_CENTER
	box.add_child(title)
	var resume := UiStyle.button("Reprendre")
	resume.pressed.connect(_set_paused.bind(false))
	var abandon := UiStyle.button("Abandonner la run")
	abandon.pressed.connect(func() -> void:
		get_tree().paused = false
		_pause_panel.visible = false
		abandon_requested.emit())
	for b in [resume, abandon]:
		box.add_child(b)
		_pause_buttons.append(b)
