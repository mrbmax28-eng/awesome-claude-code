extends GutTest
## Courbe de difficulté linéaire et difficulté adaptative.

const BLOCK := { "per_tier": 0.6, "per_biome": 0.35, "per_level": 0.08 }
const ADAPTIVE := { "deaths_in_a_row": 3, "reduction": 0.05, "applies_to": ["hp", "speed", "damage"] }


func test_base_multiplier_is_one() -> void:
	assert_almost_eq(Difficulty.multiplier(BLOCK, 0, 0, 1), 1.0, 1e-9)


func test_multiplier_is_linear_in_each_axis() -> void:
	assert_almost_eq(Difficulty.multiplier(BLOCK, 2, 0, 1), 1.0 + 1.2, 1e-9)
	assert_almost_eq(Difficulty.multiplier(BLOCK, 0, 3, 1), 1.0 + 1.05, 1e-9)
	assert_almost_eq(Difficulty.multiplier(BLOCK, 0, 0, 11), 1.0 + 0.8, 1e-9)
	assert_almost_eq(Difficulty.multiplier(BLOCK, 1, 1, 2), 1.0 + 0.6 + 0.35 + 0.08, 1e-9)
	# Linéarité : écarts constants entre paliers successifs.
	var d1 := Difficulty.multiplier(BLOCK, 0, 0, 5) - Difficulty.multiplier(BLOCK, 0, 0, 4)
	var d2 := Difficulty.multiplier(BLOCK, 0, 0, 9) - Difficulty.multiplier(BLOCK, 0, 0, 8)
	assert_almost_eq(d1, d2, 1e-9)


func test_level_below_one_does_not_reduce() -> void:
	assert_almost_eq(Difficulty.multiplier(BLOCK, 0, 0, 0), 1.0, 1e-9)


func test_max_mult_caps() -> void:
	var capped := { "per_tier": 0.5, "max_mult": 1.8 }
	assert_almost_eq(Difficulty.multiplier(capped, 10, 0, 1), 1.8, 1e-9)


func test_adaptive_triggers_after_streak() -> void:
	assert_eq(Difficulty.adaptive_multiplier(0, ADAPTIVE, true), 1.0)
	assert_eq(Difficulty.adaptive_multiplier(2, ADAPTIVE, true), 1.0)
	assert_almost_eq(Difficulty.adaptive_multiplier(3, ADAPTIVE, true), 0.95, 1e-9)
	assert_almost_eq(Difficulty.adaptive_multiplier(7, ADAPTIVE, true), 0.95, 1e-9)


func test_adaptive_can_be_disabled() -> void:
	assert_eq(Difficulty.adaptive_multiplier(5, ADAPTIVE, false), 1.0)


func test_enemy_stats_combine_curve_and_adaptive() -> void:
	var cfg := {
		"hp": BLOCK, "speed": { "per_tier": 0.1 }, "damage": { "per_tier": 0.0 },
		"adaptive": { "applies_to": ["hp"] },
	}
	var enemy := { "base_hp": 40.0, "base_speed": 7.0, "contact_damage": 10.0 }
	var s := Difficulty.enemy_stats(enemy, cfg, 1, 0, 1, 0.95)
	assert_almost_eq(float(s["hp"]), 40.0 * 1.6 * 0.95, 1e-6)
	assert_almost_eq(float(s["speed"]), 7.0 * 1.1, 1e-6)  # non concerné par l'adaptatif
	assert_almost_eq(float(s["damage"]), 10.0, 1e-6)


func test_difficulty_file_is_complete() -> void:
	var cfg := DataDB.difficulty()
	for key in ["hp", "speed", "damage", "adaptive"]:
		assert_true(cfg.has(key), key)
	for stat: String in cfg["adaptive"]["applies_to"]:
		assert_true(cfg.has(stat), "applies_to: " + stat)
