# Pack HEATING — paliers + pieds moteurs

4 modules à importer ensemble. Ils remplacent les versions précédentes de `M_FRAME_HEATING`, `M_FRAME_HEATING_PALIERS` et `Module_Ajout_Pieces_Techniques`. Les autres modules FRAME (`ZONES`, `ESQUISSES`, `ASSEMBLAGES`, `POUTRES_SOLIDES`) ne changent pas.

| Module | Version | Changement |
|---|---|---|
| `M_PILOTE_HEATING` | 2026.10.05.01 | **Nouveau** : enchaîne tout, contrôle chaque étape, tient un journal |
| `M_FRAME_HEATING_PALIERS` | 2026.10.05.09 | Mêmes règles de pièces que le module moteur, arrêt si un composant n'est pas chargé, entrées automatiques |
| `Module_Ajout_Pieces_Techniques` | 2026.10.05.18 | Entrée sans fenêtre `TECH_ExecuterAuto`, règles rouleau/moteur rendues publiques. Calcul inchangé |
| `M_FRAME_HEATING` | (sans numéro) | Mode silencieux, dernière erreur lisible, export IGES sans fenêtre. Calcul inchangé |

## Installation

1. VBE (Alt+F11) : **supprimer** les anciens `M_FRAME_HEATING`, `M_FRAME_HEATING_PALIERS` et `Module_Ajout_Pieces_Techniques` (et tout module renommé en `…1`).
2. Importer les 4 `.bas`, puis Débogage > Compiler VBAProject.
3. Lancer `PILOTE_Configurer` (une seule fois) : côté des moteurs et fichier `ASSEMBLAGE_MOTEUR_CYLINDER.asm`.

## Utilisation (assemblage HEATING actif dans Creo)

| # | Action | Ce qui est vérifié |
|---|---|---|
| 1 | `PILOTE_1_PaliersEtIGES` | Versions des modules, classeur, configuration, composants tous chargés, squelette unique, rouleaux tous portés par `heating.asm`, squelette à jour → recalcul → paliers → vérification dans `Niveaux_HEATING` → export IGES |
| 2 | **Creo** : recharger `AXES_HEATING_3D.igs` dans la fonction importée du squelette, régénérer, enregistrer | — |
| 3 | `PILOTE_2_ControleEtPieds` | Mêmes contrôles de départ → contrôle Creo (aucun KO bloquant) → pieds moteurs → contrôle final : chaque rouleau a son moteur du bon côté, chaque pied descend jusqu'au palier de son groupe, nombre de rouleaux équipés = nombre de rouleaux détectés |

`PILOTE_Verifier` lance seulement les contrôles de départ, sans rien modifier.

Au premier problème, le pilote s'arrête, dit quoi corriger et note tout dans la feuille `JOURNAL_PILOTE`. L'état courant est affiché dans `PILOTE_HEATING!C9`.

## Pourquoi l'étape 2 reste manuelle

Aucune macro ne peut redéfinir la fonction IGES importée d'un squelette existant. La seule voie automatique serait de supprimer le squelette et d'en recréer un, ce qui casserait toutes les pièces qui s'y appuient. La phase 2 vérifie donc que le rechargement a bien été fait avant de toucher aux pieds.

## Règles communes garanties par le pack

- Les **mêmes pièces porteuses** sont vues par les paliers et par le module moteur (une seule source de règles) : un rouleau équipé d'un pied a forcément son palier.
- Composant non chargé, repère `ROLL` dans un squelette, rouleau porté par un sous-assemblage (son pied se calerait sur un autre squelette) : **arrêt avant toute modification**.
- Palier à 1500 mm sous le rouleau le plus bas de chaque groupe, **dans Creo**. Le décalage Excel → Creo est mesuré dans le squelette et compensé.
- Le contrôle distingue **KO** (bloquant) et **ALERTE** (à connaître). Exemples d'alerte : squelette décalé mais compensé, cadre X/Z non centré sur l'assemblage (décision encore en attente).
