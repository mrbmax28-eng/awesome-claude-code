class_name Autopilot
extends Node
## Pilote automatique de DEBUG (--autoplay), utilisé pour le test de fumée de bout en bout.
## Stratégie : charge l'ennemi le plus proche ; s'il est englué (trop près et lent),
## se dégage à l'opposé du groupe puis revient. Drift dans les virages serrés,
## turbo quand la jauge est pleine et la cible alignée.
## Injecte ses entrées via BaseVehicle.input_override (même chemin que le joueur).

const ESCAPE_TIME := 1.1

var world: RunWorld
var _t := 0.0
var _escape_timer := 0.0
var _escape_dir := Vector2.ZERO
var _stuck_check_pos := Vector2.ZERO
var _stuck_check_t := 0.0
var _reverse_timer := 0.0


func _ready() -> void:
	world.player.input_override = _compute_input


func _compute_input() -> Dictionary:
	var dt := get_physics_process_delta_time()
	_t += dt
	var car := world.player
	var pos := Vector2(car.global_position.x, car.global_position.z)
	var speed := car.planar_velocity().length()
	var nearest := Vector2.ZERO
	var best := INF
	var centroid := Vector2.ZERO
	var close_count := 0
	for e: EnemyBase in get_tree().get_nodes_in_group("enemies"):
		if e.is_dying:
			continue
		var ep := Vector2(e.global_position.x, e.global_position.z)
		var d := pos.distance_to(ep)
		if d < best:
			best = d
			nearest = ep
		if d < 6.0:
			centroid += ep
			close_count += 1

	var fwd := Vector2(-sin(car.rotation.y), -cos(car.rotation.y))

	# Coincé contre un obstacle (moins de 1 m parcouru en 1 s) : marche arrière en braquant.
	_stuck_check_t += dt
	if _stuck_check_t >= 1.0:
		if pos.distance_to(_stuck_check_pos) < 1.0 and _reverse_timer <= 0.0:
			_reverse_timer = 0.9
		_stuck_check_pos = pos
		_stuck_check_t = 0.0
	if _reverse_timer > 0.0:
		_reverse_timer -= dt
		return { "throttle": 0.0, "brake": 1.0, "steer": 1.0, "drift": false, "turbo": false }

	var target: Vector2
	if _escape_timer > 0.0:
		_escape_timer -= dt
		target = pos + _escape_dir * 10.0
	elif close_count >= 2 and speed < 8.0:
		_escape_timer = ESCAPE_TIME
		_escape_dir = (pos - centroid / close_count).normalized()
		# Ne pas foncer dans le mur : se dégager vers le centre si on est au bord.
		if pos.length() > world.biome.half_size - 10.0:
			_escape_dir = (-pos).normalized()
		target = pos + _escape_dir * 10.0
	elif best < 9.0 and absf(fwd.angle_to((nearest - pos).normalized())) > 1.3:
		# Cible dans le cercle de braquage : on s'éloigne tout droit pour reprendre de l'angle.
		_escape_timer = ESCAPE_TIME * 0.7
		_escape_dir = fwd
		if pos.length() > world.biome.half_size - 10.0:
			_escape_dir = (-pos).normalized()
		target = pos + _escape_dir * 10.0
	elif best == INF:
		target = pos + Vector2.from_angle(_t) * 5.0  # tourne en rond pour ramasser les pièces
	else:
		target = nearest

	var angle := fwd.angle_to((target - pos).normalized())  # > 0 : cible à droite
	var ratio: float = (car as Car).energy_ratio() if car is Car else 0.0
	return {
		"throttle": 1.0,
		"brake": 0.0,
		"steer": clampf(angle * 2.5, -1.0, 1.0),
		"drift": absf(angle) > 0.7 and speed > 10.0,
		"turbo": ratio > 0.5 and absf(angle) < 0.25 and best < 20.0,
	}
