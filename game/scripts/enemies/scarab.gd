class_name Scarab
extends EnemyBase
## Scarabée des sables : chasseur au contact, modèle low-poly généré.

var _legs: Array[Node3D] = []
var _anim_t := 0.0


func _build_visual(root: Node3D) -> void:
	var c: Dictionary = def["colors"]
	var outline := Color("#141a20")
	var body := Color(c["body"])
	var shell := Color(c["shell"])
	var legs := Color(c["legs"])
	LowPoly.part(root, LowPoly.sphere(0.75, 6, 3), shell, Vector3(0, 0.55, 0.1), Vector3.ZERO, outline).scale = Vector3(1.0, 0.6, 1.25)
	LowPoly.part(root, LowPoly.sphere(0.38, 6, 3), body, Vector3(0, 0.5, -0.85), Vector3.ZERO, outline)
	LowPoly.part(root, LowPoly.prism(Vector3(0.12, 0.45, 0.12)), legs, Vector3(-0.18, 0.62, -1.2), Vector3(-70, 0, -15))
	LowPoly.part(root, LowPoly.prism(Vector3(0.12, 0.45, 0.12)), legs, Vector3(0.18, 0.62, -1.2), Vector3(-70, 0, 15))
	for z: float in [-0.35, 0.1, 0.55]:
		for side: float in [-1.0, 1.0]:
			var pivot := Node3D.new()
			pivot.position = Vector3(side * 0.6, 0.35, z)
			root.add_child(pivot)
			LowPoly.part(pivot, LowPoly.box(Vector3(0.6, 0.1, 0.1)), legs, Vector3(side * 0.3, -0.12, 0), Vector3(0, 0, side * -25))
			_legs.append(pivot)


func _animate(delta: float, speed: float) -> void:
	_anim_t += delta * (4.0 + speed * 2.0)
	for i in _legs.size():
		_legs[i].rotation.y = sin(_anim_t + i * 1.7) * 0.4 * clampf(speed / 4.0, 0.0, 1.0)
