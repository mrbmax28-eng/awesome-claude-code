# M_FRAME_HEATING_PALIERS — version 2026.10.05.07

Remplace `M_FRAME_HEATING_PALIERS6.bas`. Mêmes points d'entrée (`HEAT_PlacerPaliersRouleaux`, `HEAT_ControlerSqueletteHeating`), plus `HEAT_VersionPaliers` pour vérifier la version importée.

## Garantie

Chaque palier est à 1500 mm sous le centre du rouleau le plus bas de son groupe. C'est aussi la **première ligne du squelette** sous chaque rouleau du groupe, c'est-à-dire la ligne sur laquelle le module pieds moteurs se cale. Si cette garantie ne peut pas être tenue, la macro s'arrête sans rien modifier.

## Corrections

| Problème de la version 6 | Effet | Correction |
|---|---|---|
| Décalage Y du squelette seulement signalé | Distance réelle = 1500 − décalage (3194 mm, −194 mm…) | Arrêt si le squelette n'est pas à Y = 0 |
| Rouleau seul situé sous son groupe | Palier **au-dessus** de ce rouleau | Référence = rouleau le plus bas du groupe final |
| Niveaux entre le palier et les rouleaux laissés actifs | Le pied se cale sur ce niveau, pas sur le palier | Désactivés (`[OFF PALIER]`), listés avant confirmation |
| Palier d'un groupe sous les rouleaux du groupe inférieur | Rouleaux calés sur le mauvais palier | Arrêt avant écriture |
| Contrôle de sécurité toujours faux (`decalageY = 0`) | Aucune protection | Contrôle sur chaque rouleau du groupe |
| Pas de vérification après `HEAT_Recalculer` | Palier ignoré sans message (ex. HEATING en mode automatique) | Palier recherché dans `Niveaux_HEATING`, sinon annulation |
| Feuille `PALIERS_HEATING` effacée en cas d'arrêt ou d'annulation | Les paliers en place ne sont plus documentés | Feuille restaurée, motif noté en I3 |
| Restauration sur erreur sans recalcul | HEATING calculé avec des niveaux qui n'existent plus | `HEAT_Recalculer` relancé après restauration |
| Message de confirmation > 1024 caractères | Question OUI/NON tronquée | Texte limité, question toujours affichée |
| Contrôle : « première ligne » prise n'importe où | OK/KO sur une poutre qui ne passe pas sous le rouleau | Ligne qui croise la verticale du rouleau, vérifiée pour chaque rouleau |

## Installation

1. VBE (Alt+F11) : supprimer l'ancien `M_FRAME_HEATING_PALIERS`.
2. Importer ce `.bas`, puis Débogage > Compiler VBAProject.
3. `HEAT_VersionPaliers` doit afficher `2026.10.05.07`.

## Ordre d'utilisation

1. Squelette HEATING en 0 ; 0 ; 0 dans `heating.asm`.
2. `HEAT_PlacerPaliersRouleaux`, puis accepter l'export IGES.
3. Recharger `AXES_HEATING_3D.igs` dans le squelette.
4. `HEAT_ControlerSqueletteHeating` : toutes les lignes doivent être OK.
5. Seulement ensuite, la macro des pieds moteurs.
