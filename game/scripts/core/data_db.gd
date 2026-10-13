class_name DataDB
extends RefCounted
## Accès en lecture seule aux fichiers d'équilibrage de res://data/.
## Pas un autoload : cache statique, chargé à la demande.

const DATA_DIR := "res://data/"

static var _cache: Dictionary = {}


static func get_file(file_name: String) -> Dictionary:
	if _cache.has(file_name):
		return _cache[file_name]
	var path := DATA_DIR + file_name + ".json"
	var text := FileAccess.get_file_as_string(path)
	if text.is_empty():
		push_error("DataDB: fichier introuvable ou vide : %s" % path)
		return {}
	var parsed: Variant = JSON.parse_string(text)
	if typeof(parsed) != TYPE_DICTIONARY:
		push_error("DataDB: JSON invalide : %s" % path)
		return {}
	_cache[file_name] = parsed
	return parsed


static func get_entry(file_name: String, id: String) -> Dictionary:
	var file := get_file(file_name)
	if not file.has(id):
		push_error("DataDB: '%s' absent de %s.json" % [id, file_name])
		return {}
	return file[id]


## Liste des identifiants d'un fichier (les clés commençant par '_' sont de la doc).
static func ids(file_name: String) -> Array[String]:
	var out: Array[String] = []
	for key: String in get_file(file_name).keys():
		if not key.begins_with("_"):
			out.append(key)
	return out


static func config() -> Dictionary:
	return get_file("config")


static func vehicle(id: String) -> Dictionary:
	return get_entry("vehicles", id)


static func enemy(id: String) -> Dictionary:
	return get_entry("enemies", id)


static func biome(id: String) -> Dictionary:
	return get_entry("biomes", id)


static func economy() -> Dictionary:
	return get_file("economy")


static func difficulty() -> Dictionary:
	return get_file("difficulty")


static func clear_cache() -> void:
	_cache.clear()
