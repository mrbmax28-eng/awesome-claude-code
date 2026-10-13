class_name Economy
extends RefCounted
## Formules de coût et de conversion. Fonctions pures, paramètres issus de data/economy.json.


## Exposant du soft cap. Si non fourni : p = s × ln(k), ce qui raccorde la courbe
## polynomiale à l'exponentielle en valeur ET en pente au point n = s.
static func softcap_exponent(curve: Dictionary) -> float:
	var p: Variant = curve.get("softcap_exponent")
	if p == null or float(p) <= 0.0:
		return float(curve["softcap_start"]) * log(float(curve["growth"]))
	return float(p)


## Coût brut (flottant) du n-ième achat, n ≥ 0.
static func cost(curve: Dictionary, n: int) -> float:
	assert(n >= 0, "Economy.cost: n doit être >= 0")
	var base: float = curve["base"]
	var growth: float = curve["growth"]
	var s: float = maxf(1.0, float(curve["softcap_start"]))
	if n <= s:
		return base * pow(growth, n)
	return base * pow(growth, s) * pow(float(n) / s, softcap_exponent(curve))


## Coût affiché/payé : arrondi à l'entier supérieur.
static func cost_int(curve: Dictionary, n: int) -> int:
	return int(ceil(cost(curve, n) - 1e-9))


## Somme des coûts des achats n = from_n .. from_n + count - 1.
static func total_cost(curve: Dictionary, from_n: int, count: int) -> int:
	var total := 0
	for i in count:
		total += cost_int(curve, from_n + i)
	return total


static func curve(id: String) -> Dictionary:
	var curves: Dictionary = DataDB.economy()["curves"]
	if not curves.has(id):
		push_error("Economy: courbe '%s' inconnue" % id)
		return curves["default"]
	return curves[id]


## Conversion pièces de run -> ferraille.
## ferraille = floor(coef × pièces^exponent × mult), mult = victory_mult ou defeat_mult.
static func run_conversion(coins: int, victory: bool, conv: Dictionary) -> int:
	if coins < int(conv["min_coins"]):
		return 0
	var mult: float = conv["victory_mult"] if victory else conv["defeat_mult"]
	return int(floor(float(conv["coef"]) * pow(float(coins), float(conv["exponent"])) * mult + 1e-9))
