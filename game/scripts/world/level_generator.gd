class_name LevelGenerator
extends RefCounted
## Génération procédurale reproductible : même graine + même biome => même décor.
## Placement par échantillonnage avec rejet (espacement minimal, zone de départ dégagée).

const MAX_ATTEMPTS_PER_ITEM := 40


## Retourne une liste de { type, position: Vector2, scale, rotation, collide: bool, radius }
## (radius = rayon du collider, 0 si non collisionnable).
static func generate_decor(biome_def: Dictionary, run_seed: int) -> Array[Dictionary]:
	var rng := RandomNumberGenerator.new()
	rng.seed = run_seed
	var half: float = float(biome_def["arena"]["half_size"]) - 3.0
	var clear_r: float = biome_def["spawn_clear_radius"]
	var placed: Array[Dictionary] = []
	for rule: Dictionary in biome_def["decor"]:
		var count := rng.randi_range(int(rule["count"][0]), int(rule["count"][1]))
		var spacing: float = rule["min_spacing"]
		for i in count:
			for attempt in MAX_ATTEMPTS_PER_ITEM:
				var p := Vector2(rng.randf_range(-half, half), rng.randf_range(-half, half))
				if p.length() < clear_r:
					continue
				if not _far_enough(p, spacing, rule["type"], placed):
					continue
				var s := rng.randf_range(float(rule["scale"][0]), float(rule["scale"][1]))
				var collide := bool(rule["collide"])
				placed.append({
					"type": rule["type"],
					"position": p,
					"scale": s,
					"rotation": rng.randf() * TAU,
					"collide": collide,
					"radius": s * float(rule.get("collider_radius", 0.0)) if collide else 0.0,
				})
				break
	return placed


## L'espacement s'applique entre objets du même type et vis-à-vis des obstacles collisionnables.
static func _far_enough(p: Vector2, spacing: float, type: String, placed: Array[Dictionary]) -> bool:
	for other in placed:
		if other["type"] != type and not other["collide"]:
			continue
		if p.distance_to(other["position"]) < spacing:
			return false
	return true


## Point d'apparition d'ennemi valide : dans l'arène, loin du joueur, hors obstacles.
static func spawn_point(rng: RandomNumberGenerator, half_size: float, player_pos: Vector2,
		min_dist: float, decor: Array[Dictionary]) -> Vector2:
	var half := half_size - 4.0
	var best := Vector2(half, half)
	for attempt in 30:
		var p := Vector2(rng.randf_range(-half, half), rng.randf_range(-half, half))
		if p.distance_to(player_pos) < min_dist:
			continue
		var blocked := false
		for d in decor:
			if d["collide"] and p.distance_to(d["position"]) < float(d["radius"]) + 2.0:
				blocked = true
				break
		if not blocked:
			return p
		best = p
	return best
