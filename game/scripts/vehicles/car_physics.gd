class_name CarPhysics
extends RefCounted
## Modèle de conduite arcade de la voiture, indépendant des nœuds (testable).
## Plan horizontal (x, z) ; heading = rotation.y du nœud ; avant = -Z local.
## Drift : adhérence latérale réduite -> la voiture glisse ; le glissement remplit le turbo.

var stats: Dictionary
var velocity := Vector2.ZERO
var heading := 0.0
var turbo := 0.0
var drifting := false
var turbo_active := false
var slip := 0.0


func _init(p_stats: Dictionary) -> void:
	stats = p_stats


func forward() -> Vector2:
	return Vector2(-sin(heading), -cos(heading))


func forward_speed() -> float:
	return velocity.dot(forward())


func max_speed() -> float:
	var m: float = stats["max_speed"]
	return m * float(stats["turbo_speed_mult"]) if turbo_active else m


func turbo_ratio() -> float:
	return turbo / float(stats["turbo_max"])


## input : { throttle: 0..1, brake: 0..1, steer: -1..1 (droite > 0), drift: bool, turbo: bool }
func step(input: Dictionary, dt: float) -> void:
	var fwd := forward()
	var speed_f := velocity.dot(fwd)
	var lateral := velocity - fwd * speed_f
	var throttle: float = input.get("throttle", 0.0)
	var brake: float = input.get("brake", 0.0)
	var steer: float = clampf(input.get("steer", 0.0), -1.0, 1.0)

	drifting = bool(input.get("drift", false)) and absf(speed_f) >= float(stats["drift_min_speed"])
	_update_turbo(bool(input.get("turbo", false)), dt)

	var accel: float = stats["acceleration"]
	if turbo_active:
		accel *= float(stats["turbo_accel_mult"])
		throttle = 1.0
	var limit := max_speed()
	# L'accélération ne pousse jamais au-delà de la limite (sinon accel − decay > 0 : emballement).
	if throttle > 0.0 and speed_f < limit:
		speed_f = minf(speed_f + accel * throttle * dt, limit)
	if brake > 0.0:
		if speed_f > 0.5:
			speed_f -= float(stats["brake_deceleration"]) * brake * dt
		else:
			speed_f = maxf(speed_f - accel * brake * dt, -float(stats["reverse_max_speed"]))
	if throttle <= 0.0 and brake <= 0.0:
		speed_f = move_toward(speed_f, 0.0, float(stats["rolling_drag"]) * dt)
	if speed_f > limit:
		speed_f = move_toward(speed_f, limit, float(stats["overspeed_decay"]) * dt)

	# Braquage proportionnel à la vitesse (pas de rotation à l'arrêt), inversé en marche arrière.
	var steer_factor := clampf(absf(speed_f) / float(stats["turn_speed_ref"]), 0.0, 1.0) * signf(speed_f)
	var turn_rate := deg_to_rad(float(stats["turn_rate_deg"]))
	if drifting:
		turn_rate *= float(stats["drift_turn_mult"])

	var grip: float = stats["drift_grip"] if drifting else stats["grip"]
	lateral *= exp(-grip * dt)
	velocity = fwd * speed_f + lateral
	heading -= steer * turn_rate * steer_factor * dt

	# Glissement mesuré par rapport au nouveau cap.
	var new_fwd := forward()
	slip = (velocity - new_fwd * velocity.dot(new_fwd)).length()
	if drifting and not turbo_active:
		var fill := clampf(slip / float(stats["drift_slip_ref"]), 0.0, 1.0)
		turbo = minf(turbo + float(stats["turbo_fill_rate"]) * fill * dt, float(stats["turbo_max"]))


func _update_turbo(wants_turbo: bool, dt: float) -> void:
	if not wants_turbo:
		turbo_active = false
		return
	if not turbo_active and turbo >= float(stats["turbo_min_to_start"]):
		turbo_active = true
	if turbo_active:
		turbo = maxf(turbo - float(stats["turbo_drain_rate"]) * dt, 0.0)
		if turbo <= 0.0:
			turbo_active = false


## Après move_and_slide, le corps physique peut avoir été arrêté par un mur.
func sync_velocity(v: Vector2) -> void:
	velocity = v


## Perte de vitesse lors d'un impact (ratio 0..1).
func lose_speed(ratio: float) -> void:
	velocity *= clampf(1.0 - ratio, 0.0, 1.0)
