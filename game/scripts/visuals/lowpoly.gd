class_name LowPoly
extends RefCounted
## Fabrique de formes low-poly : primitives à normales plates, aplats de couleur, contour sombre.

const OUTLINE_SHADER := preload("res://scripts/visuals/outline.gdshader")

static var _materials: Dictionary = {}
static var _outline_materials: Dictionary = {}


static func material(color: Color, outline: Color = Color(0, 0, 0, 0), outline_width: float = 0.05) -> StandardMaterial3D:
	var key := "%s|%s|%s" % [color.to_html(), outline.to_html(), outline_width]
	if _materials.has(key):
		return _materials[key]
	var mat := StandardMaterial3D.new()
	mat.albedo_color = color
	mat.roughness = 1.0
	mat.specular_mode = BaseMaterial3D.SPECULAR_DISABLED
	mat.diffuse_mode = BaseMaterial3D.DIFFUSE_LAMBERT
	if outline.a > 0.0:
		mat.next_pass = outline_material(outline, outline_width)
	_materials[key] = mat
	return mat


static func outline_material(color: Color, width: float) -> ShaderMaterial:
	var key := "%s|%s" % [color.to_html(), width]
	if _outline_materials.has(key):
		return _outline_materials[key]
	var mat := ShaderMaterial.new()
	mat.shader = OUTLINE_SHADER
	mat.set_shader_parameter("outline_color", color)
	mat.set_shader_parameter("outline_width", width)
	_outline_materials[key] = mat
	return mat


## Convertit une primitive en ArrayMesh à normales plates (facettes visibles).
static func flat(mesh: PrimitiveMesh) -> ArrayMesh:
	var st := SurfaceTool.new()
	st.create_from(mesh, 0)
	# Sommets dédoublés par triangle -> une normale par face (rendu facetté).
	st.deindex()
	st.generate_normals()
	return st.commit()


static func box(size: Vector3) -> ArrayMesh:
	var m := BoxMesh.new()
	m.size = size
	return flat(m)


static func cylinder(radius_top: float, radius_bottom: float, height: float, sides: int = 6) -> ArrayMesh:
	var m := CylinderMesh.new()
	m.top_radius = radius_top
	m.bottom_radius = radius_bottom
	m.height = height
	m.radial_segments = sides
	m.rings = 1
	return flat(m)


static func sphere(radius: float, segments: int = 6, rings: int = 4) -> ArrayMesh:
	var m := SphereMesh.new()
	m.radius = radius
	m.height = radius * 2.0
	m.radial_segments = segments
	m.rings = rings
	return flat(m)


static func prism(size: Vector3) -> ArrayMesh:
	var m := PrismMesh.new()
	m.size = size
	return flat(m)


## Crée un MeshInstance3D coloré, ajouté à `parent`.
static func part(parent: Node3D, mesh: Mesh, color: Color, pos: Vector3 = Vector3.ZERO,
		rot_deg: Vector3 = Vector3.ZERO, outline: Color = Color(0, 0, 0, 0), outline_width: float = 0.05) -> MeshInstance3D:
	var mi := MeshInstance3D.new()
	mi.mesh = mesh
	mi.material_override = material(color, outline, outline_width)
	mi.position = pos
	mi.rotation_degrees = rot_deg
	parent.add_child(mi)
	return mi
