class_name EnemyBase
extends CharacterBody3D
## Ennemi de base : poursuit la cible, attaque au contact avec un temps de recharge,
## subit recul et étourdissement quand il est percuté.

signal killed(enemy: EnemyBase)

const LAYER_WORLD := 1
const LAYER_ENEMIES := 4

var enemy_id: String
var def: Dictionary
var max_hp: float
var hp: float
var move_speed: float
var damage: float
var target: BaseVehicle
var is_dying := false

var _attack_timer := 0.0
var _stun_timer := 0.0
var _hit_invuln := 0.0
var _spawn_timer := 0.0
var _knockback := Vector2.ZERO
var _parts: Array[MeshInstance3D] = []
var _part_materials: Array[Material] = []
var _flash_timer := 0.0
var _visual: Node3D

static var _flash_material: StandardMaterial3D


## stats : résultat de Difficulty.enemy_stats().
func setup(id: String, stats: Dictionary, p_target: BaseVehicle) -> void:
	enemy_id = id
	def = DataDB.enemy(id)
	max_hp = stats["hp"]
	hp = max_hp
	move_speed = stats["speed"]
	damage = stats["damage"]
	target = p_target
	_spawn_timer = float(def["spawn_in_s"])


func _ready() -> void:
	collision_layer = LAYER_ENEMIES
	collision_mask = LAYER_WORLD
	motion_mode = CharacterBody3D.MOTION_MODE_FLOATING
	add_to_group("enemies")
	var shape := CollisionShape3D.new()
	var cyl := CylinderShape3D.new()
	cyl.radius = float(def["radius"])
	cyl.height = 0.8
	shape.shape = cyl
	shape.position.y = 0.55
	add_child(shape)
	_visual = Node3D.new()
	add_child(_visual)
	_build_visual(_visual)
	for child in _visual.find_children("*", "MeshInstance3D", true, false):
		_parts.append(child)
		_part_materials.append((child as MeshInstance3D).material_override)
	_visual.scale = Vector3.ONE * 0.01
	rotation.y = randf() * TAU


func can_be_hit() -> bool:
	return not is_dying and _hit_invuln <= 0.0 and _spawn_timer <= 0.0


func receive_hit(amount: float, knockback: Vector2) -> void:
	if not can_be_hit():
		return
	hp -= amount
	_hit_invuln = float(def["hit_invuln_s"])
	_stun_timer = float(def["stun_s"])
	_knockback = knockback * (1.0 - float(def["knockback_resistance"]))
	_flash(0.12)
	if hp <= 0.0:
		_die()


func _physics_process(delta: float) -> void:
	if is_dying:
		return
	_attack_timer = maxf(_attack_timer - delta, 0.0)
	_hit_invuln = maxf(_hit_invuln - delta, 0.0)
	_update_flash(delta)
	if _spawn_timer > 0.0:
		_spawn_timer -= delta
		var t := 1.0 - clampf(_spawn_timer / float(def["spawn_in_s"]), 0.0, 1.0)
		_visual.scale = Vector3.ONE * maxf(ease(t, -2.0), 0.01)
		return
	_visual.scale = Vector3.ONE
	var v := Vector2(velocity.x, velocity.z)
	if _stun_timer > 0.0:
		_stun_timer -= delta
		v = _knockback
		_knockback = _knockback.move_toward(Vector2.ZERO, float(def["knockback_decay"]) * delta)
	elif is_instance_valid(target) and not target.is_dead:
		v = v.move_toward(_desired_velocity(), float(def["acceleration"]) * delta)
		_try_attack()
	else:
		v = v.move_toward(Vector2.ZERO, float(def["acceleration"]) * delta)
	velocity = Vector3(v.x, 0.0, v.y)
	move_and_slide()
	position.y = 0.0
	if v.length_squared() > 0.5:
		rotation.y = lerp_angle(rotation.y, atan2(-v.x, -v.y), 10.0 * delta)
	_animate(delta, v.length())


func _desired_velocity() -> Vector2:
	var to_target := Vector2(target.global_position.x - global_position.x,
			target.global_position.z - global_position.z)
	var desired := to_target.normalized() * move_speed
	# Séparation douce entre ennemis (évite l'empilement sans collisions physiques).
	var sep_r: float = def["separation_radius"]
	var push := Vector2.ZERO
	for other in get_tree().get_nodes_in_group("enemies"):
		if other == self:
			continue
		var d := Vector2(global_position.x - other.global_position.x, global_position.z - other.global_position.z)
		var dist := d.length()
		if dist > 0.001 and dist < sep_r:
			push += d / dist * (1.0 - dist / sep_r)
	return desired + push * float(def["separation_strength"])


func _try_attack() -> void:
	if _attack_timer > 0.0:
		return
	var reach := float(def["radius"]) + BaseVehicle.BODY_RADIUS + float(def["attack_range"])
	var d := Vector2(target.global_position.x - global_position.x, target.global_position.z - global_position.z)
	if d.length() <= reach:
		target.take_damage(damage)
		_attack_timer = float(def["attack_cooldown_s"])


func _die() -> void:
	is_dying = true
	collision_layer = 0
	killed.emit(self)
	var tw := create_tween()
	tw.tween_property(_visual, "scale", Vector3(1.4, 0.1, 1.4), 0.15)
	tw.tween_callback(queue_free)


func _flash(duration: float) -> void:
	if _flash_material == null:
		_flash_material = StandardMaterial3D.new()
		_flash_material.albedo_color = Color.WHITE
		_flash_material.shading_mode = BaseMaterial3D.SHADING_MODE_UNSHADED
	_flash_timer = duration
	for p in _parts:
		p.material_override = _flash_material


func _update_flash(delta: float) -> void:
	if _flash_timer <= 0.0:
		return
	_flash_timer -= delta
	if _flash_timer <= 0.0:
		for i in _parts.size():
			_parts[i].material_override = _part_materials[i]


# --- À surcharger ---

func _build_visual(_root: Node3D) -> void:
	pass


func _animate(_delta: float, _speed: float) -> void:
	pass
