// Contour sombre léger par coque inversée : on pousse les sommets vers l'extérieur
// depuis l'origine du mesh (adapté aux primitives convexes centrées) et on ne dessine
// que les faces arrière.
shader_type spatial;
render_mode unshaded, cull_front, depth_draw_opaque;

uniform vec4 outline_color : source_color = vec4(0.1, 0.07, 0.05, 1.0);
uniform float outline_width = 0.05;

void vertex() {
	VERTEX += normalize(VERTEX + vec3(1e-5)) * outline_width;
}

void fragment() {
	ALBEDO = outline_color.rgb;
}
