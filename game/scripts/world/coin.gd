class_name Coin
extends Node3D
## Pièce ramassable : éjectée à la mort d'un ennemi, attirée par l'aimant du véhicule.
## Pas de physique : distance testée chaque frame (peu d'objets, coût négligeable).

signal collected(value: int)

const COLOR_COIN := Color("#ffd23f")
const COLOR_OUTLINE := Color("#6b4a00")

var value := 1
var target: BaseVehicle
var _velocity := Vector2.ZERO
var _magnet_speed := 20.0
var _magnet_radius := 5.0
var _pickup_radius := 1.4
var _friction := 10.0
var _lifetime := 30.0
var _magnetized := false
var _t := 0.0


func setup(p_target: BaseVehicle, scatter: Vector2) -> void:
	target = p_target
	_velocity = scatter
	var cfg: Dictionary = DataDB.config()["pickups"]
	_magnet_speed = cfg["coin_magnet_speed"]
	_lifetime = cfg["coin_lifetime_s"]
	_pickup_radius = cfg["coin_pickup_radius"]
	_friction = cfg["coin_friction"]
	_magnet_radius = target.stats["magnet_radius"]


func _ready() -> void:
	LowPoly.part(self, LowPoly.cylinder(0.32, 0.32, 0.1, 8), COLOR_COIN, Vector3.ZERO, Vector3(90, 0, 0), COLOR_OUTLINE, 0.03)
	position.y = 0.6


## Force l'attraction vers le véhicule quelle que soit la distance.
func attract() -> void:
	_magnetized = true


func _physics_process(delta: float) -> void:
	_t += delta
	_lifetime -= delta
	rotation.y += delta * 4.0
	if _lifetime <= 0.0:
		queue_free()
		return
	if not is_instance_valid(target) or target.is_dead:
		return
	var to := Vector2(target.global_position.x - global_position.x, target.global_position.z - global_position.z)
	var dist := to.length()
	if dist < _pickup_radius:
		collected.emit(value)
		queue_free()
		return
	if _magnetized or dist < _magnet_radius:
		_magnetized = true
		_velocity = to.normalized() * maxf(_magnet_speed, target.planar_velocity().length() + 5.0)
	else:
		_velocity = _velocity.move_toward(Vector2.ZERO, _friction * delta)
	position += Vector3(_velocity.x, 0.0, _velocity.y) * delta
	position.y = 0.6 + sin(_t * 5.0) * 0.12
