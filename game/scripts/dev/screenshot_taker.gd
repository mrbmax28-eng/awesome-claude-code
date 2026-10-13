class_name ScreenshotTaker
extends Node
## DEBUG : enregistre une capture de la vue toutes les `interval_s` secondes (--screenshot-dir).

var out_dir: String
var interval_s := 4.0
var _t := 0.0
var _index := 0


func _ready() -> void:
	process_mode = Node.PROCESS_MODE_ALWAYS
	DirAccess.make_dir_recursive_absolute(out_dir)


func _process(delta: float) -> void:
	_t += delta
	if _t < interval_s:
		return
	_t = 0.0
	var img := get_viewport().get_texture().get_image()
	if img:
		img.save_png("%s/shot_%02d.png" % [out_dir, _index])
		_index += 1
