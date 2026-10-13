class_name CmdArgs
extends RefCounted
## Arguments utilisateur passés après « -- » : godot --path game -- --seed=42 --autoplay


static func _parse() -> Dictionary:
	var out := {}
	for arg: String in OS.get_cmdline_user_args():
		if not arg.begins_with("--"):
			continue
		var body := arg.substr(2)
		var eq := body.find("=")
		if eq == -1:
			out[body] = "true"
		else:
			out[body.substr(0, eq)] = body.substr(eq + 1)
	return out


static func has(key: String) -> bool:
	return _parse().has(key)


static func get_string(key: String, fallback: String) -> String:
	return String(_parse().get(key, fallback))


static func get_int(key: String, fallback: int) -> int:
	var v: String = get_string(key, "")
	return int(v) if v.is_valid_int() else fallback


static func get_float(key: String, fallback: float) -> float:
	var v: String = get_string(key, "")
	return float(v) if v.is_valid_float() else fallback
