extends GutTest
## Intégrité des fichiers data/ : clés requises, références croisées valides.

const VEHICLE_STATS := [
	"max_hp", "max_speed", "reverse_max_speed", "acceleration", "brake_deceleration",
	"rolling_drag", "overspeed_decay", "turn_rate_deg", "turn_speed_ref", "grip", "drift_grip",
	"drift_turn_mult", "drift_min_speed", "drift_slip_ref", "turbo_max", "turbo_fill_rate",
	"turbo_drain_rate", "turbo_min_to_start", "turbo_speed_mult", "turbo_accel_mult",
	"ram_min_speed_ratio", "ram_damage", "turbo_ram_mult", "ram_knockback", "ram_speed_loss",
	"turbo_contact_immunity", "contact_invuln_s", "magnet_radius", "hitbox_radius",
]
const ENEMY_KEYS := [
	"base_hp", "base_speed", "contact_damage", "attack_cooldown_s", "attack_range", "coin_drop",
	"radius", "acceleration", "knockback_resistance", "stun_s", "hit_invuln_s",
	"separation_radius", "separation_strength", "spawn_in_s", "knockback_decay", "colors",
]


func test_all_data_files_parse() -> void:
	for f in ["config", "vehicles", "enemies", "biomes", "difficulty", "economy", "input"]:
		assert_false(DataDB.get_file(f).is_empty(), f)


func test_vehicles_have_all_stats_and_valid_upgrades() -> void:
	var curves: Dictionary = DataDB.economy()["curves"]
	for id in DataDB.ids("vehicles"):
		var v := DataDB.vehicle(id)
		for key: String in VEHICLE_STATS:
			assert_true((v["stats"] as Dictionary).has(key), "%s.%s" % [id, key])
		assert_true(ResourceLoader.exists(v["scene"]), "%s.scene" % id)
		var ups: Array = v["upgrades"]
		assert_between(ups.size(), 3, 5, "%s : 3 à 5 améliorations" % id)
		for up: Dictionary in ups:
			assert_true((v["stats"] as Dictionary).has(up["stat"]), up["id"])
			assert_true(curves.has(up["curve"]), up["id"])


func test_enemies_have_required_keys() -> void:
	for id in DataDB.ids("enemies"):
		var e := DataDB.enemy(id)
		for key: String in ENEMY_KEYS:
			assert_true(e.has(key), "%s.%s" % [id, key])
		assert_lte(int(e["coin_drop"][0]), int(e["coin_drop"][1]), id)


func test_biome_references_are_valid() -> void:
	var enemies := DataDB.ids("enemies")
	for id in DataDB.ids("biomes"):
		var b := DataDB.biome(id)
		for e: String in b["enemies"]:
			assert_has(enemies, e, "%s.enemies" % id)
		for w: Dictionary in b["waves"]:
			assert_has(enemies, w["enemy"], "%s.waves" % id)
			assert_gt(int(w["max_alive"]), 0)
		for rule: Dictionary in b["decor"]:
			assert_lte(int(rule["count"][0]), int(rule["count"][1]))
		assert_true(ResourceLoader.exists(b["script"]), "%s.script" % id)
		for key in ["dominant", "ground", "sky", "outline"]:
			assert_true((b["palette"] as Dictionary).has(key), "%s.palette.%s" % [id, key])


func test_input_actions_registered() -> void:
	for action in DataDB.ids("input"):
		assert_true(InputMap.has_action(action), action)
		var events := InputMap.action_get_events(action)
		assert_gt(events.size(), 0, action)
	# Chaque action de jeu doit être jouable au clavier/souris ET à la manette.
	for action in DataDB.ids("input"):
		var has_pad := false
		var has_kbm := false
		for ev in InputMap.action_get_events(action):
			if ev is InputEventJoypadButton or ev is InputEventJoypadMotion:
				has_pad = true
			if ev is InputEventKey or ev is InputEventMouseButton:
				has_kbm = true
		assert_true(has_pad and has_kbm, "%s : clavier/souris + manette" % action)


func test_physical_keys_map_to_actions() -> void:
	# Liaison en touche physique : la position W (QWERTY) = Z (AZERTY) déclenche « accelerate ».
	var ev := InputEventKey.new()
	ev.physical_keycode = KEY_W
	ev.pressed = true
	assert_true(InputMap.event_is_action(ev, "accelerate"))
	var pad := InputEventJoypadMotion.new()
	pad.axis = JOY_AXIS_TRIGGER_RIGHT
	pad.axis_value = 1.0
	assert_true(InputMap.event_is_action(pad, "accelerate"))
	var turbo := InputEventJoypadButton.new()
	turbo.button_index = JOY_BUTTON_A
	turbo.pressed = true
	assert_true(InputMap.event_is_action(turbo, "turbo"))
