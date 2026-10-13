extends GutTest
## Formules de coût (exponentielle + soft cap) et conversion de fin de run.

const CURVE := { "base": 10.0, "growth": 1.15, "softcap_start": 50, "softcap_exponent": null }


func test_first_purchase_costs_base() -> void:
	assert_almost_eq(Economy.cost(CURVE, 0), 10.0, 1e-9)


func test_exponential_below_softcap() -> void:
	for n in [1, 10, 25, 50]:
		assert_almost_eq(Economy.cost(CURVE, n), 10.0 * pow(1.15, n), 1e-6 * pow(1.15, n), "n=%d" % n)


func test_auto_softcap_exponent_is_s_ln_k() -> void:
	assert_almost_eq(Economy.softcap_exponent(CURVE), 50.0 * log(1.15), 1e-9)


func test_softcap_is_continuous_at_start() -> void:
	# Raccord en valeur : la formule polynomiale évaluée en n = s redonne base × k^s.
	var s := 50
	var poly_at_s := 10.0 * pow(1.15, s) * pow(float(s) / s, Economy.softcap_exponent(CURVE))
	assert_almost_eq(poly_at_s, Economy.cost(CURVE, s), 1e-6)


func test_softcap_slope_matches_exponential() -> void:
	# Raccord en pente (C1) : rapport cost(s+1)/cost(s) ≈ k, à l'ordre 1 près.
	var ratio := Economy.cost(CURVE, 51) / Economy.cost(CURVE, 50)
	assert_almost_eq(ratio, 1.15, 0.01)


func test_softcap_grows_slower_than_exponential() -> void:
	for n in [60, 100, 200]:
		assert_lt(Economy.cost(CURVE, n), 10.0 * pow(1.15, n), "n=%d" % n)


func test_cost_strictly_increasing() -> void:
	var prev := -1.0
	for n in 300:
		var c := Economy.cost(CURVE, n)
		assert_gt(c, prev, "n=%d" % n)
		prev = c


func test_explicit_softcap_exponent_is_used() -> void:
	var curve := CURVE.duplicate()
	curve["softcap_exponent"] = 3.0
	assert_almost_eq(Economy.cost(curve, 100), 10.0 * pow(1.15, 50) * pow(2.0, 3.0), 1e-3)


func test_cost_int_rounds_up() -> void:
	assert_eq(Economy.cost_int(CURVE, 1), 12)  # 11.5 -> 12
	assert_eq(Economy.cost_int(CURVE, 0), 10)


func test_total_cost_sums_purchases() -> void:
	assert_eq(Economy.total_cost(CURVE, 0, 3), 10 + 12 + 14)  # 10, 11.5, 13.225


func test_data_curves_are_valid() -> void:
	for id: String in DataDB.economy()["curves"]:
		var c: Dictionary = Economy.curve(id)
		assert_gt(float(c["base"]), 0.0, id)
		assert_gt(float(c["growth"]), 1.0, id)
		assert_gte(int(c["softcap_start"]), 1, id)


func test_conversion_formula() -> void:
	var conv := { "coef": 1.0, "exponent": 0.85, "victory_mult": 1.25, "defeat_mult": 1.0, "min_coins": 1 }
	assert_eq(Economy.run_conversion(0, true, conv), 0)
	assert_eq(Economy.run_conversion(1, false, conv), 1)
	assert_eq(Economy.run_conversion(60, false, conv), int(floor(pow(60.0, 0.85))))  # 32
	assert_eq(Economy.run_conversion(60, true, conv), int(floor(pow(60.0, 0.85) * 1.25)))  # 40


func test_conversion_is_monotonic_and_victory_pays_more() -> void:
	var conv: Dictionary = DataDB.economy()["run_conversion"]
	var prev := -1
	for coins in range(0, 2000, 7):
		var s := Economy.run_conversion(coins, false, conv)
		assert_gte(s, prev)
		assert_gte(Economy.run_conversion(coins, true, conv), s)
		prev = s
