class_name EndScreen
extends Control
## Écran de fin de run : résultat, pièces récoltées, conversion en ferraille.

signal retry_requested
signal menu_requested

var _buttons: Array[Button] = []


func show_run(run: RunState) -> void:
	set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	var bg := ColorRect.new()
	bg.color = Color(UiStyle.BG, 0.85)
	bg.set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	add_child(bg)
	var center := CenterContainer.new()
	center.set_anchors_and_offsets_preset(Control.PRESET_FULL_RECT)
	add_child(center)
	var panel := UiStyle.panel()
	center.add_child(panel)
	var col := VBoxContainer.new()
	col.add_theme_constant_override("separation", 12)
	panel.add_child(col)
	var titles := {
		RunState.Result.VICTORY: ["Biome nettoyé !", UiStyle.ACCENT],
		RunState.Result.DEFEAT: ["Épave…", UiStyle.HP],
		RunState.Result.ABANDONED: ["Run abandonnée", UiStyle.MUTED],
	}
	var t: Array = titles.get(run.result, ["Fin de run", UiStyle.TEXT])
	var title := UiStyle.label(t[0], 44, t[1])
	title.horizontal_alignment = HORIZONTAL_ALIGNMENT_CENTER
	col.add_child(title)
	var lines := [
		"Pièces récoltées : %d" % run.coins,
		"Ennemis détruits : %d / %d" % [run.kills, run.enemies_total],
		"Durée : %d:%02d" % [int(run.elapsed_s) / 60, int(run.elapsed_s) % 60],
	]
	if run.difficulty_mult < 1.0:
		lines.append("Difficulté adaptative : −%d %%" % int(round((1.0 - run.difficulty_mult) * 100.0)))
	for line: String in lines:
		col.add_child(UiStyle.label(line, 22))
	col.add_child(UiStyle.label("+ %d ferraille   (total %d)" % [run.scrap_gained, int(GameState.profile["scrap"])], 28, UiStyle.SCRAP))
	col.add_child(UiStyle.label("graine %d" % run.run_seed, 12, UiStyle.MUTED))
	var retry := UiStyle.button("Relancer")
	retry.pressed.connect(func() -> void: retry_requested.emit())
	var menu := UiStyle.button("Menu")
	menu.pressed.connect(func() -> void: menu_requested.emit())
	for b in [retry, menu]:
		col.add_child(b)
		_buttons.append(b)
	UiStyle.focus_first(_buttons)
