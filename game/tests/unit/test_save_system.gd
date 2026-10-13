extends GutTest
## Sauvegarde JSON : aller-retour, fusion avec les défauts, récupération après corruption.

const PATH := "user://test_save.json"

var sys: SaveSystem


func before_each() -> void:
	for p in [PATH, PATH + ".bak", PATH + ".tmp"]:
		if FileAccess.file_exists(p):
			DirAccess.remove_absolute(p)
	sys = SaveSystem.new(PATH, 1)


func after_all() -> void:
	for p in [PATH, PATH + ".bak", PATH + ".tmp"]:
		if FileAccess.file_exists(p):
			DirAccess.remove_absolute(p)


func test_missing_file_gives_defaults() -> void:
	assert_eq(sys.load_data(), SaveSystem.default_data())


func test_roundtrip() -> void:
	var data := SaveSystem.default_data()
	data["profile"]["scrap"] = 1234
	data["profile"]["death_streak"] = { "desert": 2 }
	data["settings"]["screen_shake"] = false
	assert_true(sys.save(data))
	var loaded := sys.load_data()
	assert_eq(int(loaded["profile"]["scrap"]), 1234)
	assert_eq(int(loaded["profile"]["death_streak"]["desert"]), 2)
	assert_false(loaded["settings"]["screen_shake"])
	assert_false(FileAccess.file_exists(PATH + ".tmp"))


func test_missing_fields_are_filled_from_defaults() -> void:
	var f := FileAccess.open(PATH, FileAccess.WRITE)
	f.store_string('{"version": 1, "profile": {"scrap": 5}}')
	f.close()
	var loaded := sys.load_data()
	assert_eq(int(loaded["profile"]["scrap"]), 5)
	assert_eq(int(loaded["profile"]["runs_played"]), 0)
	assert_true(loaded["settings"].has("adaptive_difficulty"))


func test_wrong_types_fall_back_to_default() -> void:
	var f := FileAccess.open(PATH, FileAccess.WRITE)
	f.store_string('{"profile": {"scrap": "beaucoup"}, "settings": {"screen_shake": 0}}')
	f.close()
	var loaded := sys.load_data()
	assert_eq(int(loaded["profile"]["scrap"]), 0)
	assert_true(loaded["settings"]["screen_shake"])


func test_corrupted_file_uses_backup() -> void:
	var data := SaveSystem.default_data()
	data["profile"]["scrap"] = 10
	sys.save(data)
	data["profile"]["scrap"] = 20
	sys.save(data)  # le .bak contient maintenant scrap = 10
	var f := FileAccess.open(PATH, FileAccess.WRITE)
	f.store_string("{ tronqué")
	f.close()
	assert_eq(int(sys.load_data()["profile"]["scrap"]), 10)
