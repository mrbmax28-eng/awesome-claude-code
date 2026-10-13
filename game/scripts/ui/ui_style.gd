class_name UiStyle
extends RefCounted
## Thème d'interface commun (couleurs, tailles de police). Échelle de texte réglable (accessibilité, Phase 6).

const BG := Color("#1d1a26")
const PANEL := Color("#2b2638")
const TEXT := Color("#f6efe4")
const MUTED := Color("#b8aec8")
const ACCENT := Color("#ffb13b")
const HP := Color("#e8484f")
const TURBO := Color("#ffb13b")
const TURBO_ACTIVE := Color("#ffe27a")
const SCRAP := Color("#9fb4c7")

static var text_scale := 1.0


static func font_size(base: int) -> int:
	return int(round(base * text_scale))


static func label(text: String, size: int = 20, color: Color = TEXT) -> Label:
	var l := Label.new()
	l.text = text
	l.add_theme_font_size_override("font_size", font_size(size))
	l.add_theme_color_override("font_color", color)
	l.add_theme_color_override("font_outline_color", Color(0, 0, 0, 0.8))
	l.add_theme_constant_override("outline_size", 4)
	return l


static func button(text: String, size: int = 24) -> Button:
	var b := Button.new()
	b.text = text
	b.custom_minimum_size = Vector2(320, 56)
	b.add_theme_font_size_override("font_size", font_size(size))
	var normal := _box(PANEL, Color("#4a4160"))
	var hover := _box(Color("#3a3350"), ACCENT)
	b.add_theme_stylebox_override("normal", normal)
	b.add_theme_stylebox_override("hover", hover)
	b.add_theme_stylebox_override("focus", hover)
	b.add_theme_stylebox_override("pressed", _box(ACCENT.darkened(0.3), ACCENT))
	b.add_theme_color_override("font_color", TEXT)
	b.add_theme_color_override("font_focus_color", ACCENT)
	b.add_theme_color_override("font_hover_color", ACCENT)
	return b


static func bar(fill: Color, size: Vector2) -> ProgressBar:
	var p := ProgressBar.new()
	p.custom_minimum_size = size
	p.show_percentage = false
	p.add_theme_stylebox_override("background", _box(Color(0, 0, 0, 0.55), Color(0, 0, 0, 0.9), 2))
	p.add_theme_stylebox_override("fill", _box(fill, Color(0, 0, 0, 0), 0))
	return p


static func panel() -> PanelContainer:
	var p := PanelContainer.new()
	var sb := _box(Color(BG, 0.92), Color("#4a4160"), 2)
	sb.content_margin_left = 32
	sb.content_margin_right = 32
	sb.content_margin_top = 24
	sb.content_margin_bottom = 24
	p.add_theme_stylebox_override("panel", sb)
	return p


static func _box(bg: Color, border: Color, border_w: int = 2) -> StyleBoxFlat:
	var sb := StyleBoxFlat.new()
	sb.bg_color = bg
	sb.border_color = border
	sb.set_border_width_all(border_w)
	sb.set_corner_radius_all(6)
	sb.content_margin_left = 12
	sb.content_margin_right = 12
	return sb


## Tableau de boutons navigable clavier/manette : focus initial sur le premier.
static func focus_first(buttons: Array[Button]) -> void:
	if buttons.is_empty():
		return
	for i in buttons.size():
		var prev := buttons[(i - 1 + buttons.size()) % buttons.size()]
		var next := buttons[(i + 1) % buttons.size()]
		buttons[i].focus_neighbor_top = buttons[i].get_path_to(prev)
		buttons[i].focus_neighbor_bottom = buttons[i].get_path_to(next)
	buttons[0].grab_focus.call_deferred()
