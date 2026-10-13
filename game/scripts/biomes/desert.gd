class_name DesertBiome
extends BiomeBase
## Désert (dominante orange) : rochers, cactus, dunes, cailloux.
## Adaptation (data) : ressources accrues pour le buggy, ralentissement pour le char.


func _build_wall_visuals(height: float) -> void:
	var rng := RandomNumberGenerator.new()
	rng.seed = hash(biome_id)
	var spacing: float = def["arena"]["wall_rocks_spacing"]
	var edge := half_size + 1.5
	var steps := int(ceil(half_size * 2.0 / spacing)) + 1
	for side in 4:
		for i in steps:
			var t := -half_size + i * spacing
			var pos := Vector3(t, 0, -edge)
			match side:
				1: pos = Vector3(t, 0, edge)
				2: pos = Vector3(-edge, 0, t)
				3: pos = Vector3(edge, 0, t)
			var s := rng.randf_range(1.6, 2.6)
			var rock := _rock(s, height / 3.0 * rng.randf_range(0.9, 1.4))
			rock.position = pos
			rock.rotation.y = rng.randf() * TAU
			add_child(rock)


func _build_decor_item(item: Dictionary) -> Node3D:
	var s: float = item["scale"]
	match String(item["type"]):
		"rock":
			return _rock(s, 1.0)
		"cactus":
			return _cactus(s)
		"dune":
			return _dune(s)
		"pebble":
			var n := Node3D.new()
			LowPoly.part(n, LowPoly.sphere(s, 5, 3), palette["rock"], Vector3(0, s * 0.3, 0))
			return n
	push_warning("DesertBiome: type de décor inconnu '%s'" % item["type"])
	return null


func _rock(s: float, height_mult: float) -> Node3D:
	var n := Node3D.new()
	var o := outline_color()
	LowPoly.part(n, LowPoly.cylinder(0.55 * s, 0.85 * s, 1.1 * s * height_mult, 5), palette["rock"],
			Vector3(0, 0.5 * s * height_mult, 0), Vector3.ZERO, o, 0.06)
	LowPoly.part(n, LowPoly.cylinder(0.2 * s, 0.45 * s, 0.6 * s * height_mult, 5), palette["rock_dark"],
			Vector3(0.45 * s, 0.3 * s * height_mult, 0.3 * s), Vector3(0, 30, 0), o, 0.05)
	return n


func _cactus(s: float) -> Node3D:
	var n := Node3D.new()
	var o := outline_color()
	var c: Color = palette["cactus"]
	LowPoly.part(n, LowPoly.cylinder(0.32 * s, 0.38 * s, 2.6 * s, 6), c, Vector3(0, 1.3 * s, 0), Vector3.ZERO, o, 0.05)
	LowPoly.part(n, LowPoly.cylinder(0.2 * s, 0.22 * s, 0.9 * s, 6), c, Vector3(0.55 * s, 1.6 * s, 0), Vector3.ZERO, o, 0.04)
	LowPoly.part(n, LowPoly.box(Vector3(0.45 * s, 0.25 * s, 0.25 * s)), palette["cactus_dark"], Vector3(0.3 * s, 1.15 * s, 0), Vector3.ZERO, o, 0.04)
	LowPoly.part(n, LowPoly.cylinder(0.18 * s, 0.2 * s, 0.7 * s, 6), c, Vector3(-0.5 * s, 1.9 * s, 0), Vector3.ZERO, o, 0.04)
	LowPoly.part(n, LowPoly.box(Vector3(0.4 * s, 0.22 * s, 0.22 * s)), palette["cactus_dark"], Vector3(-0.28 * s, 1.6 * s, 0), Vector3.ZERO, o, 0.04)
	return n


func _dune(s: float) -> Node3D:
	var n := Node3D.new()
	var mi := LowPoly.part(n, LowPoly.sphere(1.0, 7, 3), palette["dune"], Vector3(0, -0.15 * s, 0))
	mi.scale = Vector3(s, 0.35 * s, s * 0.7)
	return n
