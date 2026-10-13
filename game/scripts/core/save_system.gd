class_name SaveSystem
extends RefCounted
## Sauvegarde JSON locale, écriture atomique (fichier .tmp puis renommage) avec copie .bak.
## Le contenu chargé est fusionné avec les valeurs par défaut : un champ manquant
## (ancienne version de sauvegarde) reprend sa valeur par défaut.

var path: String
var version: int


func _init(save_path: String, save_version: int) -> void:
	path = save_path
	version = save_version


static func default_data() -> Dictionary:
	return {
		"version": 1,
		"profile": {
			"scrap": 0,
			"runs_played": 0,
			"runs_won": 0,
			"total_coins": 0,
			"best_coins": 0,
			"total_kills": 0,
			"death_streak": {},
		},
		"settings": {
			"screen_shake": true,
			"adaptive_difficulty": true,
			"fullscreen": false,
		},
	}


func save(data: Dictionary) -> bool:
	data["version"] = version
	var tmp_path := path + ".tmp"
	var file := FileAccess.open(tmp_path, FileAccess.WRITE)
	if file == null:
		push_error("SaveSystem: écriture impossible (%s)" % error_string(FileAccess.get_open_error()))
		return false
	file.store_string(JSON.stringify(data, "\t"))
	file.close()
	if FileAccess.file_exists(path):
		DirAccess.copy_absolute(path, path + ".bak")
	var err := DirAccess.rename_absolute(tmp_path, path)
	if err != OK:
		push_error("SaveSystem: renommage impossible (%s)" % error_string(err))
		return false
	return true


## Charge la sauvegarde ; en cas de fichier corrompu, tente le .bak puis les valeurs par défaut.
func load_data() -> Dictionary:
	for candidate: String in [path, path + ".bak"]:
		if not FileAccess.file_exists(candidate):
			continue
		var json := JSON.new()
		if json.parse(FileAccess.get_file_as_string(candidate)) == OK and typeof(json.data) == TYPE_DICTIONARY:
			return merge_defaults(json.data, default_data())
		push_warning("SaveSystem: fichier illisible ignoré : %s" % candidate)
	return default_data()


## Fusion récursive : garde les valeurs de `data`, complète avec `defaults`.
## Les types incompatibles (ex. chaîne à la place d'un nombre) reprennent la valeur par défaut.
static func merge_defaults(data: Dictionary, defaults: Dictionary) -> Dictionary:
	var out := defaults.duplicate(true)
	for key: Variant in data:
		if not out.has(key):
			out[key] = data[key]
		elif typeof(out[key]) == TYPE_DICTIONARY and typeof(data[key]) == TYPE_DICTIONARY:
			if (out[key] as Dictionary).is_empty():
				out[key] = data[key]
			else:
				out[key] = merge_defaults(data[key], out[key])
		elif _compatible(out[key], data[key]):
			out[key] = data[key]
	return out


static func _compatible(default_value: Variant, value: Variant) -> bool:
	var a := typeof(default_value)
	var b := typeof(value)
	if a == b:
		return true
	# JSON ne distingue pas int et float.
	return (a == TYPE_INT or a == TYPE_FLOAT) and (b == TYPE_INT or b == TYPE_FLOAT)
