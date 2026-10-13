class_name Car
extends BaseVehicle
## Voiture : drift (remplit la jauge de turbo) et turbo (vitesse accrue, immunité
## au contact, dégâts de percussion multipliés).

const COLOR_BODY := Color("#d8343f")
const COLOR_CABIN := Color("#f3efe6")
const COLOR_GLASS := Color("#3b5a7a")
const COLOR_WHEEL := Color("#2a2a2e")
const COLOR_OUTLINE := Color("#1c1214")
const COLOR_TURBO := Color("#ffb13b")

var physics: CarPhysics
var _body_pivot: Node3D
var _front_wheels: Array[Node3D] = []
var _flames: Array[MeshInstance3D] = []
var _flash_timer := 0.0


func setup(id: String) -> void:
	super.setup(id)
	physics = CarPhysics.new(stats)


func _vehicle_physics(delta: float, input: Dictionary) -> void:
	physics.heading = rotation.y
	var was_turbo := physics.turbo_active
	physics.step(input, delta)
	rotation.y = physics.heading
	velocity = Vector3(physics.velocity.x, 0.0, physics.velocity.y)
	move_and_slide()
	var pre := physics.velocity.length()
	physics.sync_velocity(planar_velocity())
	if get_slide_collision_count() > 0 and pre - physics.velocity.length() > 6.0:
		impact.emit(0.5)
	if was_turbo != physics.turbo_active or physics.drifting or physics.turbo_active:
		energy_changed.emit(physics.turbo_ratio(), physics.turbo_active)
	_animate(delta, input)


func _on_enemy_contact(enemy: EnemyBase) -> void:
	var speed := physics.velocity.length()
	var vmax: float = stats["max_speed"]
	var ratio := speed / vmax
	if not physics.turbo_active and ratio < float(stats["ram_min_speed_ratio"]):
		return
	if not enemy.can_be_hit():
		return
	var damage := float(stats["ram_damage"]) * minf(ratio, float(stats["turbo_speed_mult"]))
	if physics.turbo_active:
		damage *= float(stats["turbo_ram_mult"])
	var dir := Vector2(enemy.global_position.x - global_position.x,
			enemy.global_position.z - global_position.z).normalized()
	dir = (dir + physics.velocity.normalized()).normalized()
	enemy.receive_hit(damage, dir * float(stats["ram_knockback"]) * clampf(ratio, 0.5, 1.5))
	if not physics.turbo_active:
		physics.lose_speed(float(stats["ram_speed_loss"]))
	impact.emit(0.6)


func is_damage_immune() -> bool:
	return physics.turbo_active and bool(stats["turbo_contact_immunity"])


func energy_ratio() -> float:
	return physics.turbo_ratio()


func _on_damaged() -> void:
	_flash_timer = 0.15


# --- Visuel low-poly généré ---

func _build_visual(root: Node3D) -> void:
	_body_pivot = Node3D.new()
	root.add_child(_body_pivot)
	var o := COLOR_OUTLINE
	LowPoly.part(_body_pivot, LowPoly.box(Vector3(1.7, 0.55, 3.2)), COLOR_BODY, Vector3(0, 0.6, 0), Vector3.ZERO, o)
	LowPoly.part(_body_pivot, LowPoly.box(Vector3(1.4, 0.5, 1.5)), COLOR_CABIN, Vector3(0, 1.1, 0.25), Vector3.ZERO, o)
	LowPoly.part(_body_pivot, LowPoly.box(Vector3(1.3, 0.38, 0.1)), COLOR_GLASS, Vector3(0, 1.12, -0.52), Vector3(-20, 0, 0))
	LowPoly.part(_body_pivot, LowPoly.box(Vector3(1.75, 0.12, 0.5)), COLOR_CABIN, Vector3(0, 0.95, 1.45), Vector3.ZERO, o)
	for side: float in [-1.0, 1.0]:
		LowPoly.part(_body_pivot, LowPoly.box(Vector3(0.3, 0.18, 0.08)), Color("#fff6c8"), Vector3(side * 0.55, 0.65, -1.62))
		var flame := LowPoly.part(_body_pivot, LowPoly.prism(Vector3(0.3, 0.7, 0.3)), COLOR_TURBO, Vector3(side * 0.45, 0.5, 1.9), Vector3(90, 0, 0))
		flame.visible = false
		_flames.append(flame)
	for z: float in [-1.05, 1.05]:
		for x: float in [-0.9, 0.9]:
			var pivot := Node3D.new()
			pivot.position = Vector3(x, 0.38, z)
			root.add_child(pivot)
			LowPoly.part(pivot, LowPoly.cylinder(0.38, 0.38, 0.32, 8), COLOR_WHEEL, Vector3.ZERO, Vector3(0, 0, 90), o, 0.03)
			if z < 0.0:
				_front_wheels.append(pivot)


func _animate(delta: float, input: Dictionary) -> void:
	var steer: float = input.get("steer", 0.0)
	for w in _front_wheels:
		w.rotation.y = lerp_angle(w.rotation.y, -steer * 0.5, 12.0 * delta)
	var target_roll := 0.0
	if physics.drifting:
		target_roll = steer * 0.12
	_body_pivot.rotation.z = lerpf(_body_pivot.rotation.z, target_roll, 8.0 * delta)
	var pitch := -0.05 if physics.turbo_active else 0.0
	_body_pivot.rotation.x = lerpf(_body_pivot.rotation.x, pitch, 6.0 * delta)
	for f in _flames:
		f.visible = physics.turbo_active
		if f.visible:
			f.scale = Vector3.ONE * randf_range(0.8, 1.25)
	if _flash_timer > 0.0:
		_flash_timer -= delta
		_visual.visible = int(_flash_timer * 40.0) % 2 == 0
	else:
		_visual.visible = true
