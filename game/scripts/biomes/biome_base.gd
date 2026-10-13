class_name BiomeBase
extends Node3D
## Base d'un biome : sol, enceinte, éclairage, ambiance (palette data/biomes.json),
## décor procédural et règles d'adaptation au véhicule actif.
## Les biomes dérivés fournissent la construction des objets de décor spécifiques.

var biome_id: String
var def: Dictionary
var palette: Dictionary = {}
var decor: Array[Dictionary] = []
var half_size: float


func setup(id: String, run_seed: int) -> void:
	biome_id = id
	def = DataDB.biome(id)
	for key: String in def["palette"]:
		palette[key] = Color(def["palette"][key])
	half_size = def["arena"]["half_size"]
	decor = LevelGenerator.generate_decor(def, run_seed)


func _ready() -> void:
	_build_environment()
	_build_ground()
	_build_walls()
	for item in decor:
		var node := _build_decor_item(item)
		if node == null:
			continue
		node.position = Vector3(item["position"].x, 0.0, item["position"].y)
		node.rotation.y = item["rotation"]
		add_child(node)
		if item["collide"]:
			_add_static_collider(node, float(item["radius"]))


## Multiplicateur d'adaptation pour le véhicule actif (1.0 si non défini).
func adaptation(vehicle_id: String, key: String) -> float:
	var table: Dictionary = def.get("adaptation", {})
	return float((table.get(vehicle_id, {}) as Dictionary).get(key, 1.0))


func outline_color() -> Color:
	return palette.get("outline", Color.BLACK)


func _build_environment() -> void:
	var env := Environment.new()
	env.background_mode = Environment.BG_COLOR
	env.background_color = palette["sky"]
	env.ambient_light_source = Environment.AMBIENT_SOURCE_COLOR
	env.ambient_light_color = palette["ambient"]
	var light: Dictionary = def["lighting"]
	env.ambient_light_energy = light["ambient_energy"]
	# Tonemap linéaire : conserve les aplats saturés de la palette.
	env.tonemap_mode = Environment.TONE_MAPPER_LINEAR
	var we := WorldEnvironment.new()
	we.environment = env
	add_child(we)
	var sun := DirectionalLight3D.new()
	sun.light_color = palette["sun"]
	sun.light_energy = light["sun_energy"]
	sun.rotation_degrees = Vector3(light["sun_pitch_deg"], light["sun_yaw_deg"], 0)
	sun.shadow_enabled = true
	sun.shadow_opacity = light["shadow_opacity"]
	sun.shadow_blur = 2.0
	sun.directional_shadow_mode = DirectionalLight3D.SHADOW_ORTHOGONAL
	sun.directional_shadow_max_distance = 80.0
	add_child(sun)


func _build_ground() -> void:
	var size := half_size * 2.0 + float(def["arena"]["ground_margin"])
	var ground := LowPoly.part(self, LowPoly.box(Vector3(size, 1.0, size)), palette["ground"], Vector3(0, -0.5, 0))
	ground.name = "Ground"
	var body := StaticBody3D.new()
	body.collision_layer = BaseVehicle.LAYER_WORLD
	var shape := CollisionShape3D.new()
	var box := BoxShape3D.new()
	box.size = Vector3(size, 1.0, size)
	shape.shape = box
	shape.position.y = -0.5
	body.add_child(shape)
	add_child(body)


## Enceinte : une collision invisible par côté + des rochers décoratifs le long du bord.
func _build_walls() -> void:
	var h: float = def["arena"]["wall_height"]
	for i in 4:
		var body := StaticBody3D.new()
		body.collision_layer = BaseVehicle.LAYER_WORLD
		var shape := CollisionShape3D.new()
		var box := BoxShape3D.new()
		var horizontal := i < 2
		box.size = Vector3(half_size * 2.0 + 4.0, h, 2.0) if horizontal else Vector3(2.0, h, half_size * 2.0 + 4.0)
		shape.shape = box
		var sgn := -1.0 if i % 2 == 0 else 1.0
		shape.position = Vector3(0, h * 0.5, sgn * (half_size + 1.0)) if horizontal else Vector3(sgn * (half_size + 1.0), h * 0.5, 0)
		body.add_child(shape)
		add_child(body)
	_build_wall_visuals(h)


# --- À surcharger ---

func _build_wall_visuals(_height: float) -> void:
	pass


func _build_decor_item(_item: Dictionary) -> Node3D:
	return null


func _add_static_collider(node: Node3D, radius: float) -> void:
	var body := StaticBody3D.new()
	body.collision_layer = BaseVehicle.LAYER_WORLD
	var shape := CollisionShape3D.new()
	var cyl := CylinderShape3D.new()
	cyl.radius = radius
	cyl.height = 3.0
	shape.shape = cyl
	shape.position.y = 1.5
	body.add_child(shape)
	node.add_child(body)
