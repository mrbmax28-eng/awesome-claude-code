class_name BaseVehicle
extends CharacterBody3D
## Base commune à tous les véhicules : PV, invulnérabilité, détection de contact
## avec les ennemis, jauge d'énergie générique pour le HUD.
## Chaque véhicule dérivé implémente _vehicle_physics() et sa mécanique propre.

signal hp_changed(hp: float, max_hp: float)
signal energy_changed(ratio: float, active: bool)
signal died
## Choc encaissé par le véhicule (percussion, mur) : 0..1, utilisé pour les secousses.
signal impact(strength: float)

const LAYER_WORLD := 1
const LAYER_PLAYER := 2
const LAYER_ENEMIES := 4  # bit 3
## Rayon du corps physique (géométrie, pas équilibrage).
const BODY_RADIUS := 1.0

var vehicle_id: String
var stats: Dictionary
var max_hp: float
var hp: float
var is_dead := false
var controls_enabled := true
## Entrées injectées par l'autopilote de debug (null = entrées joueur).
var input_override: Callable

var _invuln_timer := 0.0
var _hitbox: Area3D
var _visual: Node3D


func setup(id: String) -> void:
	vehicle_id = id
	stats = DataDB.vehicle(id)["stats"]
	max_hp = stats["max_hp"]
	hp = max_hp


func _ready() -> void:
	collision_layer = LAYER_PLAYER
	collision_mask = LAYER_WORLD
	add_to_group("player")
	motion_mode = CharacterBody3D.MOTION_MODE_FLOATING
	var shape := CollisionShape3D.new()
	var capsule := CylinderShape3D.new()
	capsule.radius = BODY_RADIUS
	capsule.height = 1.0
	shape.shape = capsule
	# Décollé du sol (y verrouillé à 0) : pas de contact permanent avec le sol.
	shape.position.y = 0.65
	add_child(shape)
	_hitbox = Area3D.new()
	_hitbox.collision_layer = 0
	_hitbox.collision_mask = LAYER_ENEMIES
	_hitbox.monitorable = false
	var hit_shape := CollisionShape3D.new()
	var sphere := SphereShape3D.new()
	sphere.radius = float(stats["hitbox_radius"])
	hit_shape.shape = sphere
	hit_shape.position.y = 0.6
	_hitbox.add_child(hit_shape)
	add_child(_hitbox)
	_visual = Node3D.new()
	_visual.name = "Visual"
	add_child(_visual)
	_build_visual(_visual)
	hp_changed.emit(hp, max_hp)


func _physics_process(delta: float) -> void:
	if is_dead:
		return
	_invuln_timer = maxf(_invuln_timer - delta, 0.0)
	_vehicle_physics(delta, read_input())
	position.y = 0.0
	for body in _hitbox.get_overlapping_bodies():
		if body is EnemyBase and not (body as EnemyBase).is_dying:
			_on_enemy_contact(body)


func read_input() -> Dictionary:
	if not controls_enabled:
		return {}
	if input_override.is_valid():
		return input_override.call()
	return {
		"throttle": Input.get_action_strength("accelerate"),
		"brake": Input.get_action_strength("brake"),
		"steer": Input.get_axis("steer_left", "steer_right"),
		"drift": Input.is_action_pressed("drift"),
		"turbo": Input.is_action_pressed("turbo"),
	}


func take_damage(amount: float) -> void:
	if is_dead or _invuln_timer > 0.0 or is_damage_immune():
		return
	hp = maxf(hp - amount, 0.0)
	_invuln_timer = float(stats["contact_invuln_s"])
	hp_changed.emit(hp, max_hp)
	_on_damaged()
	if hp <= 0.0:
		is_dead = true
		died.emit()


func is_damage_immune() -> bool:
	return false


func planar_velocity() -> Vector2:
	return Vector2(velocity.x, velocity.z)


# --- À surcharger ---

func _vehicle_physics(_delta: float, _input: Dictionary) -> void:
	pass


func _on_enemy_contact(_enemy: EnemyBase) -> void:
	pass


func _build_visual(_root: Node3D) -> void:
	pass


func _on_damaged() -> void:
	pass
