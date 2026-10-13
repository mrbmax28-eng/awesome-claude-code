# Drift & Scrap — jeu incrémental 3/4 low-poly

> Titre provisoire. Prototype **Phase 1** : voiture contrôlable, biome Désert, une vague
> d'ennemis, récolte de pièces, fin de run avec conversion en ferraille, sauvegarde.

## Prérequis

- **Godot 4.6.x** (testé avec 4.6.3-stable, binaire standard, pas la version .NET).
- Rendu : **Compatibility** (OpenGL 3.3) — choisi pour tourner sur des GPU d'entrée de gamme.
- Physique : **GodotPhysics3D** (voir « Choix techniques »).

## Lancer le jeu

```bash
# depuis le dossier game/
godot --path .                    # lance le jeu (menu principal)
godot --path . -e                 # ouvre l'éditeur
godot --path . -- --seed=1234     # force la graine de la run (debug, reproductible)
```

Sous Windows : `Godot_v4.6.3-stable_win64.exe --path .` (ou ouvrir `project.godot` dans l'éditeur, F5).

## Contrôles

| Action | Clavier / souris | Manette Xbox |
|---|---|---|
| Accélérer | W / ↑ (Z en AZERTY) | RT |
| Freiner / marche arrière | S / ↓ | LT |
| Diriger | A / D / ← / → (Q / D en AZERTY) | Stick gauche |
| Drift (remplit le turbo) | Espace / clic droit | RB ou X |
| Turbo | Maj / clic gauche | A |
| Zoom | Molette | Croix haut / bas |
| Pause | Échap / P | Start |
| Menus | Flèches + Entrée, souris | Croix / stick + A |

Les touches de déplacement sont liées en **position physique** : W/A/S/D sur QWERTY
correspondent automatiquement à Z/Q/S/D sur AZERTY. Liaisons modifiables dans `data/input.json`.

**Boucle** : drifte pour remplir la jauge, déclenche le turbo pour percuter les scarabées
(dégâts ×2,5, immunité au contact). Percuter sans turbo blesse aussi, à partir de 35 % de la
vitesse max. Les pièces sont aspirées à proximité. Run gagnée quand la vague est détruite ;
les pièces sont converties en ferraille : `floor(pièces^0,85 × 1,25)` en victoire, `× 1,0` en défaite.

## Tests

```bash
# Tests unitaires (GUT 9.6.1) : formules de coût, conversion, difficulté, conduite,
# génération procédurale, sauvegarde, intégrité des données.
godot --headless --path . -s addons/gut/gut_cmdln.gd -gconfig=res://.gutconfig.json

# Test de fumée de bout en bout : runs complètes pilotées par l'autopilote.
tests/smoke/run_smoke.sh godot 42 7 1234
```

Options de debug (après `--`) : `--autoplay` (autopilote), `--quit-after-run`,
`--save-path=…` (sauvegarde isolée), `--screenshot-dir=…` + `--screenshot-interval=s`,
`--quit-after-s=s`, `--time-scale=x`.

## Organisation

```
data/                 Équilibrage et paramètres (JSON) — aucune valeur d'équilibrage dans le code
  vehicles.json       Stats et améliorations par véhicule
  enemies.json        Stats de base des ennemis
  biomes.json         Palette, éclairage, décor procédural, vagues, table d'adaptation
  difficulty.json     Courbe de difficulté (formule documentée dans le fichier)
  economy.json        Courbes de coût + soft cap, conversion de fin de run
  config.json         Sauvegarde, caméra, run, ramassage
  input.json          Liaisons clavier / souris / manette
scenes/               main, world, vehicles, biomes, enemies, ui (scènes minces, visuels générés)
scripts/core/         game_state (SEUL autoload), data_db, economy, difficulty, save_system,
                      input_setup, run_state, cmd_args
scripts/vehicles/     base_vehicle, car, car_physics (modèle de conduite pur, testable)
scripts/biomes/       biome_base, desert
scripts/world/        run_world, wave_spawner, level_generator, camera_rig, coin
scripts/enemies/      enemy_base, scarab
scripts/ui/           hud, main_menu, end_screen, ui_style
scripts/visuals/      lowpoly (fabrique de meshes facettés), outline.gdshader
scripts/dev/          autopilot, screenshot_taker (outils de debug uniquement)
scripts/progression/  (Phase 2 : skill_tree, tech_nodes, milestones)
tests/unit/           Tests GUT
tests/smoke/          Test de bout en bout
addons/gut/           Framework de test GUT (MIT)
```

### Singleton

`GameState` (`scripts/core/game_state.gd`) est le seul autoload. Il détient le profil et les
réglages, gère la sauvegarde (auto toutes les 60 s via `config.json`, et à la fermeture de la
fenêtre) et applique le résultat des runs. Aucune logique de gameplay en run.

### Sauvegarde

`user://save.json` (Windows : `%APPDATA%\Godot\app_userdata\Drift & Scrap\save.json`).
Écriture atomique (`.tmp` puis renommage) avec copie `.bak` ; fichier corrompu → reprise du
`.bak`, sinon valeurs par défaut. Les champs manquants d'une ancienne sauvegarde sont complétés.

## Formules

- **Coût** (`economy.json`) : `coût(n) = base × k^n` pour `n ≤ s` (s = 50), puis
  `base × k^s × (n/s)^p`. Par défaut `p = s × ln k`, ce qui raccorde les deux morceaux en
  valeur **et** en pente (croissance polynomiale de degré ≈ 7 pour k = 1,15).
- **Difficulté** (`difficulty.json`) : `stat = base × (1 + a·palier + b·biome + c·(niveau − 1))`,
  bornée par `max_mult` si présent, puis × 0,95 si le joueur a perdu 3 runs d'affilée dans ce biome.
- **Conversion** : `ferraille = floor(coef × pièces^exposant × mult)`.

## Choix techniques

- **Compatibility plutôt que Forward+** : cible entrée de gamme, et la DA (aplats, pas d'effets
  volumétriques) n'exploite pas Forward+. Ombres directionnelles douces disponibles.
- **GodotPhysics3D plutôt que Jolt** : avec Jolt (défaut de Godot 4.6), 4 runs sur 12
  identiques se sont bloquées en headless (« Jolt Physics job system exceeded the maximum number
  of jobs »). Avec GodotPhysics3D : 0 blocage. Le jeu n'utilise que des corps cinématiques simples.
- **Contours** : coque inversée (passe de shader `cull_front`) sur chaque primitive convexe.
- **Visuels générés dans le code** (pas d'assets externes à ce stade, cf. `CREDITS.md`).
