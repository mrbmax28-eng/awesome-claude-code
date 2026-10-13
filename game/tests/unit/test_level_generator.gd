extends GutTest
## Génération procédurale : reproductibilité par graine et respect des règles.


func _biome() -> Dictionary:
	return DataDB.biome("desert")


func test_same_seed_same_layout() -> void:
	var a := LevelGenerator.generate_decor(_biome(), 1234)
	var b := LevelGenerator.generate_decor(_biome(), 1234)
	assert_eq(a.size(), b.size())
	for i in a.size():
		assert_eq(a[i]["type"], b[i]["type"])
		assert_eq(a[i]["position"], b[i]["position"])
		assert_eq(a[i]["scale"], b[i]["scale"])


func test_different_seed_different_layout() -> void:
	var a := LevelGenerator.generate_decor(_biome(), 1)
	var b := LevelGenerator.generate_decor(_biome(), 2)
	var same := a.size() == b.size()
	if same:
		for i in a.size():
			if a[i]["position"] != b[i]["position"]:
				same = false
				break
	assert_false(same)


func test_spawn_area_is_clear_and_items_inside_arena() -> void:
	var def := _biome()
	var clear_r: float = def["spawn_clear_radius"]
	var half: float = def["arena"]["half_size"]
	for seed_value in [3, 77, 4242]:
		for item in LevelGenerator.generate_decor(def, seed_value):
			var p: Vector2 = item["position"]
			assert_gte(p.length(), clear_r)
			assert_lte(absf(p.x), half)
			assert_lte(absf(p.y), half)


func test_min_spacing_between_same_type() -> void:
	var def := _biome()
	var spacing := {}
	for rule: Dictionary in def["decor"]:
		spacing[rule["type"]] = float(rule["min_spacing"])
	var items := LevelGenerator.generate_decor(def, 99)
	for i in items.size():
		for j in range(i + 1, items.size()):
			if items[i]["type"] == items[j]["type"]:
				var d: float = (items[i]["position"] as Vector2).distance_to(items[j]["position"])
				assert_gte(d, spacing[items[i]["type"]])


func test_counts_within_rule_bounds() -> void:
	var def := _biome()
	var items := LevelGenerator.generate_decor(def, 5)
	for rule: Dictionary in def["decor"]:
		var n := items.filter(func(it: Dictionary) -> bool: return it["type"] == rule["type"]).size()
		assert_lte(n, int(rule["count"][1]), rule["type"])
		# Le rejet peut placer moins que le minimum si la place manque ; on tolère 20 %.
		assert_gte(n, int(floor(int(rule["count"][0]) * 0.8)), rule["type"])


func test_spawn_point_respects_distance() -> void:
	var rng := RandomNumberGenerator.new()
	rng.seed = 10
	var decor := LevelGenerator.generate_decor(_biome(), 10)
	for i in 50:
		var p := LevelGenerator.spawn_point(rng, 36.0, Vector2.ZERO, 14.0, decor)
		assert_gte(p.length(), 14.0)
