extends GutTest
## Modèle de conduite : vitesse max, drift -> turbo, turbo, pas d'emballement.

const DT := 1.0 / 60.0

var stats: Dictionary


func before_each() -> void:
	stats = DataDB.vehicle("car")["stats"]


func _run(car: CarPhysics, input: Dictionary, seconds: float) -> void:
	for i in int(seconds / DT):
		car.step(input, DT)


func test_throttle_reaches_but_never_exceeds_max_speed() -> void:
	var car := CarPhysics.new(stats)
	_run(car, { "throttle": 1.0 }, 6.0)
	assert_almost_eq(car.forward_speed(), float(stats["max_speed"]), 0.01)


func test_no_runaway_with_sustained_turbo_and_throttle() -> void:
	# Régression : accélération − overspeed_decay > 0 faisait croître la vitesse sans borne.
	var car := CarPhysics.new(stats)
	car.turbo = float(stats["turbo_max"])
	var cap := float(stats["max_speed"]) * float(stats["turbo_speed_mult"]) + 0.01
	for i in 1200:
		car.step({ "throttle": 1.0, "turbo": i < 300 }, DT)
		assert_lte(car.velocity.length(), cap)


func test_drift_fills_turbo() -> void:
	var car := CarPhysics.new(stats)
	_run(car, { "throttle": 1.0 }, 3.0)
	_run(car, { "throttle": 1.0, "steer": 1.0, "drift": true }, 1.0)
	assert_true(car.drifting)
	assert_gt(car.turbo, 0.0)


func test_no_turbo_gain_without_drift() -> void:
	var car := CarPhysics.new(stats)
	_run(car, { "throttle": 1.0, "steer": 1.0 }, 4.0)
	assert_eq(car.turbo, 0.0)


func test_drift_slides_more_than_grip() -> void:
	var grip_car := CarPhysics.new(stats)
	var drift_car := CarPhysics.new(stats)
	_run(grip_car, { "throttle": 1.0 }, 3.0)
	_run(drift_car, { "throttle": 1.0 }, 3.0)
	_run(grip_car, { "throttle": 1.0, "steer": 1.0 }, 0.5)
	_run(drift_car, { "throttle": 1.0, "steer": 1.0, "drift": true }, 0.5)
	assert_gt(drift_car.slip, grip_car.slip * 2.0)


func test_turbo_requires_minimum_gauge() -> void:
	var car := CarPhysics.new(stats)
	car.turbo = float(stats["turbo_min_to_start"]) - 1.0
	car.step({ "turbo": true }, DT)
	assert_false(car.turbo_active)
	car.turbo = float(stats["turbo_min_to_start"])
	car.step({ "turbo": true }, DT)
	assert_true(car.turbo_active)


func test_turbo_raises_speed_and_drains() -> void:
	var car := CarPhysics.new(stats)
	_run(car, { "throttle": 1.0 }, 4.0)
	car.turbo = float(stats["turbo_max"])
	_run(car, { "throttle": 1.0, "turbo": true }, 1.0)
	assert_gt(car.forward_speed(), float(stats["max_speed"]) + 1.0)
	assert_lt(car.turbo, float(stats["turbo_max"]))


func test_no_rotation_when_stopped() -> void:
	var car := CarPhysics.new(stats)
	_run(car, { "steer": 1.0 }, 1.0)
	assert_eq(car.heading, 0.0)


func test_reverse_is_capped() -> void:
	var car := CarPhysics.new(stats)
	_run(car, { "brake": 1.0 }, 5.0)
	assert_almost_eq(car.forward_speed(), -float(stats["reverse_max_speed"]), 0.01)
