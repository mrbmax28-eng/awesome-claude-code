class_name Difficulty
extends RefCounted
## Courbe de difficulté documentée dans data/difficulty.json. Fonctions pures.


## 1 + per_tier×tier + per_biome×biome_index + per_level×(level − 1), borné par max_mult.
static func multiplier(block: Dictionary, tier: int, biome_index: int, level: int) -> float:
	var m := 1.0
	m += float(block.get("per_tier", 0.0)) * tier
	m += float(block.get("per_biome", 0.0)) * biome_index
	m += float(block.get("per_level", 0.0)) * maxi(level - 1, 0)
	if block.has("max_mult"):
		m = minf(m, float(block["max_mult"]))
	return maxf(m, 0.0)


## Multiplicateur adaptatif : 1 − reduction si le joueur a enchaîné assez de morts dans ce biome.
static func adaptive_multiplier(death_streak: int, cfg: Dictionary, enabled: bool) -> float:
	if not enabled or death_streak < int(cfg["deaths_in_a_row"]):
		return 1.0
	return 1.0 - float(cfg["reduction"])


## Stats effectives d'un ennemi : {hp, speed, damage}.
static func enemy_stats(enemy_def: Dictionary, cfg: Dictionary, tier: int, biome_index: int,
		level: int, adaptive_mult: float) -> Dictionary:
	var applies: Array = cfg["adaptive"]["applies_to"]
	var out := {}
	var pairs := { "hp": "base_hp", "speed": "base_speed", "damage": "contact_damage" }
	for stat: String in pairs:
		var value := float(enemy_def[pairs[stat]]) * multiplier(cfg[stat], tier, biome_index, level)
		if applies.has(stat):
			value *= adaptive_mult
		out[stat] = value
	return out
