class_name RunState
extends RefCounted
## État d'une run en cours (non persistant : une run interrompue est perdue).

enum Result { NONE, VICTORY, DEFEAT, ABANDONED }

var run_seed: int
var biome_id: String
var vehicle_id: String
var level: int = 1
var difficulty_mult: float = 1.0
var coins: int = 0
var kills: int = 0
var enemies_total: int = 0
var elapsed_s: float = 0.0
var result: Result = Result.NONE
var scrap_gained: int = 0


func _init(p_seed: int, p_biome: String, p_vehicle: String) -> void:
	run_seed = p_seed
	biome_id = p_biome
	vehicle_id = p_vehicle


func is_victory() -> bool:
	return result == Result.VICTORY


func summary() -> Dictionary:
	return {
		"seed": run_seed, "biome": biome_id, "vehicle": vehicle_id, "level": level,
		"result": Result.keys()[result], "coins": coins, "kills": kills,
		"enemies_total": enemies_total, "elapsed_s": snappedf(elapsed_s, 0.01),
		"scrap_gained": scrap_gained, "difficulty_mult": difficulty_mult,
	}
