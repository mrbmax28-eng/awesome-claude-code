#!/usr/bin/env bash
# Test de fumée de bout en bout : lance N runs pilotées par l'autopilote (headless,
# 60 fps fixes), vérifie que chaque run se termine et que la conversion est cohérente.
# Usage : tests/smoke/run_smoke.sh [chemin/vers/godot] [graines...]
set -euo pipefail
GODOT="${1:-godot}"
shift || true
SEEDS=("${@:-42 7 1234 99 2026}")
read -r -a SEEDS <<< "${SEEDS[*]}"
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
fail=0
for seed in "${SEEDS[@]}"; do
  line="$(timeout 120 "$GODOT" --headless --path "$ROOT" --fixed-fps 60 -- \
    --autoplay --seed="$seed" --quit-after-run --save-path="$TMP/save_$seed.json" 2>&1 \
    | grep '^\[run\] end' || true)"
  if [[ -z "$line" ]]; then
    echo "ÉCHEC graine $seed : la run ne s'est pas terminée"; fail=1; continue
  fi
  python3 - "$seed" "${line#\[run\] end }" "$ROOT/data/economy.json" <<'PY' || fail=1
import json, math, sys
seed, run, econ = sys.argv[1], json.loads(sys.argv[2]), json.load(open(sys.argv[3]))
c = econ["run_conversion"]
mult = c["victory_mult"] if run["result"] == "VICTORY" else c["defeat_mult"]
expected = 0 if run["coins"] < c["min_coins"] else math.floor(c["coef"] * run["coins"] ** c["exponent"] * mult + 1e-9)
ok = run["result"] in ("VICTORY", "DEFEAT") and run["kills"] <= run["enemies_total"] \
     and run["scrap_gained"] == expected and (run["result"] != "VICTORY" or run["kills"] == run["enemies_total"])
print(("OK   " if ok else "ÉCHEC") + f" graine {seed} : {run['result']} kills={run['kills']}/{run['enemies_total']} "
      f"pièces={run['coins']} ferraille={run['scrap_gained']} (attendu {expected}) t={run['elapsed_s']}s")
sys.exit(0 if ok else 1)
PY
done
exit $fail
