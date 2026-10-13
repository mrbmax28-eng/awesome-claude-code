class_name CameraRig
extends Node3D
## Caméra 3/4 fixe (~55° d'inclinaison), suivi amorti avec anticipation,
## zoom borné, secousses (screen shake) modérées désactivables.

var target: BaseVehicle
var camera: Camera3D
var cfg: Dictionary
var distance: float
var shake_enabled := true

var _trauma := 0.0
var _noise_t := 0.0


func _ready() -> void:
	cfg = DataDB.config()["camera"]
	distance = cfg["distance"]
	camera = Camera3D.new()
	camera.fov = cfg["fov_deg"]
	camera.current = true
	add_child(camera)
	rotation_degrees.x = -float(cfg["pitch_deg"])
	shake_enabled = bool(GameState.settings["screen_shake"])
	_place_camera(Vector2.ZERO)


func snap_to_target() -> void:
	if target:
		global_position = Vector3(target.global_position.x, 0.0, target.global_position.z)


func add_trauma(amount: float) -> void:
	if shake_enabled:
		_trauma = clampf(_trauma + amount, 0.0, 1.0)


func _unhandled_input(event: InputEvent) -> void:
	if event.is_action_pressed("zoom_in"):
		distance = maxf(distance - float(cfg["zoom_step"]), float(cfg["zoom_min"]))
	elif event.is_action_pressed("zoom_out"):
		distance = minf(distance + float(cfg["zoom_step"]), float(cfg["zoom_max"]))


func _process(delta: float) -> void:
	if target and is_instance_valid(target):
		var v := target.planar_velocity() * float(cfg["lookahead_s"])
		var goal := Vector3(target.global_position.x + v.x, 0.0, target.global_position.z + v.y)
		var k := 1.0 - exp(-float(cfg["follow_sharpness"]) * delta)
		global_position = global_position.lerp(goal, k)
	_trauma = maxf(_trauma - float(cfg["shake_decay"]) * delta, 0.0)
	_noise_t += delta * 30.0
	var s := _trauma * _trauma
	var offset := Vector2(sin(_noise_t * 1.3), cos(_noise_t * 1.7)) * s * float(cfg["shake_max_offset"])
	_place_camera(offset)
	camera.rotation.z = deg_to_rad(float(cfg["shake_max_roll_deg"])) * s * sin(_noise_t * 2.1)


## Le rig est incliné ; la caméra recule le long de son axe local Z.
func _place_camera(offset: Vector2) -> void:
	camera.position = Vector3(offset.x, offset.y, distance)
