class_name MainMenu
extends Control
## Menu principal (le hub complet arrive en Phase 2) : stats du profil, lancer une run, options.

signal start_run_requested
signal quit_requested

var _main_box: VBoxContainer
var _options_box: VBoxContainer
var _stats_label: Label
var _main_buttons: Array[Button] = []
var _option_buttons: Array[Button] = []


func _ready() -> void:
	set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	var bg := ColorRect.new()
	bg.color = UiStyle.BG
	bg.set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	add_child(bg)
	var center := CenterContainer.new()
	center.set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	add_child(center)
	var col := VBoxContainer.new()
	col.add_theme_constant_override("separation", 18)
	center.add_child(col)
	var title := UiStyle.label("DRIFT & SCRAP", 64, UiStyle.ACCENT)
	title.horizontal_alignment = HORIZONTAL_ALIGNMENT_CENTER
	col.add_child(title)
	var sub := UiStyle.label("Prototype — Phase 1", 18, UiStyle.MUTED)
	sub.horizontal_alignment = HORIZONTAL_ALIGNMENT_CENTER
	col.add_child(sub)
	_stats_label = UiStyle.label("", 20)
	_stats_label.horizontal_alignment = HORIZONTAL_ALIGNMENT_CENTER
	col.add_child(_stats_label)

	_main_box = VBoxContainer.new()
	_main_box.add_theme_constant_override("separation", 12)
	col.add_child(_main_box)
	var play := UiStyle.button("Lancer une run (Désert)")
	play.pressed.connect(func() -> void: start_run_requested.emit())
	var options := UiStyle.button("Options")
	options.pressed.connect(_show_options.bind(true))
	var quit := UiStyle.button("Quitter")
	quit.pressed.connect(func() -> void: quit_requested.emit())
	for b in [play, options, quit]:
		_main_box.add_child(b)
		_main_buttons.append(b)

	_options_box = VBoxContainer.new()
	_options_box.add_theme_constant_override("separation", 12)
	_options_box.visible = false
	col.add_child(_options_box)
	_add_toggle("screen_shake", "Secousses d'écran")
	_add_toggle("adaptive_difficulty", "Difficulté adaptative")
	_add_toggle("fullscreen", "Plein écran")
	var back := UiStyle.button("Retour")
	back.pressed.connect(_show_options.bind(false))
	_options_box.add_child(back)
	_option_buttons.append(back)

	var help := UiStyle.label(
		"Clavier : ZQSD / WASD / flèches · Espace = drift · Maj = turbo · Échap = pause · molette = zoom\n" +
		"Manette : RT accélérer · LT freiner · stick G diriger · RB/X drift · A turbo · Start pause",
		14, UiStyle.MUTED)
	help.horizontal_alignment = HORIZONTAL_ALIGNMENT_CENTER
	col.add_child(help)

	GameState.profile_changed.connect(refresh)
	refresh()
	UiStyle.focus_first(_main_buttons)


func refresh() -> void:
	var p := GameState.profile
	_stats_label.text = "Ferraille : %d   ·   Runs : %d (victoires %d)   ·   Record : %d pièces" % [
		int(p["scrap"]), int(p["runs_played"]), int(p["runs_won"]), int(p["best_coins"])]


func _add_toggle(key: String, text: String) -> void:
	var b := UiStyle.button("")
	var update := func() -> void:
		b.text = "%s : %s" % [text, "oui" if GameState.settings[key] else "non"]
	update.call()
	b.pressed.connect(func() -> void:
		GameState.set_setting(key, not bool(GameState.settings[key]))
		update.call())
	_options_box.add_child(b)
	_option_buttons.append(b)


func _show_options(show: bool) -> void:
	_main_box.visible = not show
	_options_box.visible = show
	UiStyle.focus_first(_option_buttons if show else _main_buttons)
