class_name InputSetup
extends RefCounted
## Crée les actions de l'InputMap à partir de data/input.json (clavier physique, souris, manette).


static func apply(actions: Dictionary) -> void:
	for action: String in actions:
		if action.begins_with("_"):
			continue
		var def: Dictionary = actions[action]
		if InputMap.has_action(action):
			InputMap.erase_action(action)
		InputMap.add_action(action, float(def.get("deadzone", 0.5)))
		for key_name: String in def.get("keys", []):
			var code := OS.find_keycode_from_string(key_name)
			if code == KEY_NONE:
				push_error("InputSetup: touche inconnue '%s'" % key_name)
				continue
			var ev := InputEventKey.new()
			ev.physical_keycode = code
			InputMap.action_add_event(action, ev)
		for button: Variant in def.get("joy_buttons", []):
			var jb := InputEventJoypadButton.new()
			jb.button_index = int(button) as JoyButton
			jb.device = -1
			InputMap.action_add_event(action, jb)
		for axis_def: Variant in def.get("joy_axes", []):
			var ja := InputEventJoypadMotion.new()
			ja.axis = int(axis_def[0]) as JoyAxis
			ja.axis_value = signf(float(axis_def[1]))
			ja.device = -1
			InputMap.action_add_event(action, ja)
		for mb: Variant in def.get("mouse_buttons", []):
			var me := InputEventMouseButton.new()
			me.button_index = int(mb) as MouseButton
			InputMap.action_add_event(action, me)
