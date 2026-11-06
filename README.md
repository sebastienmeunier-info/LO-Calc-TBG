# Tableaux de bord Calc — PoC extension LibreOffice (Python)

Extension de preuve de concept pour LibreOffice Calc (7.4+, testé pour la série 24.x)
répliquant les tableaux de bord Excel : plusieurs tables croisées dynamiques
(DataPilot), graphiques liés, et un « segment » (slicer) basique.

## Contenu

- `dashboard_comp.py` — composant UNO Python (protocol handler) :
  - `create` : génère une feuille « Dashboard » avec 3 tables croisées
    dynamiques + 3 graphiques à partir de la 1re feuille de données.
  - `slicer` : dialogue à deux étapes (champ, puis valeur) qui applique un
    filtre de page synchronisé sur toutes les tables croisées du classeur.
  - `refresh` : réactualise toutes les tables.
- `Addons.xcu` — menu « Tableau de bord » dans la barre de menus de Calc +
  enregistrement du protocol handler (`com.example.lo-dashboard-ext:*`).
- `description.xml`, `META-INF/manifest.xml` — métadonnées d'extension.

## Installation

1. Ouvrez `lo-dashboard-ext-0.1.0.oxt` (double-clic, ou dans LibreOffice :
   **Outils > Gestionnaire d'extensions > Ajouter**).
2. Redémarrez LibreOffice (obligatoire pour les composants Python UNO).
3. Prérequis : le runtime Python de LibreOffice (inclus par défaut ; sur Linux,
   installer `libreoffice-script-provider-python` si absent).

## Utilisation

1. Ouvrez un classeur dont la **première feuille** contient vos données :
   ligne 1 = en-têtes, colonnes de catégories (texte) + colonnes numériques.
2. Menu **Tableau de bord > Créer un tableau de bord...** : la feuille
   « Dashboard » est créée avec 3 TCD et 3 graphiques.
3. Menu **Tableau de bord > Segment (filtre)...** : choisissez un champ puis
   une valeur ; le filtre est appliqué à toutes les TCD concernées
   (« (Tous) » pour retirer le filtre).
4. Menu **Tableau de bord > Actualiser les tables** après modification
   des données sources.

## Limites connues (PoC)

- La détection des colonnes est heuristique (1re feuille uniquement).
- Le slicer ne supporte qu'une valeur à la fois (limite de l'API DataPilot :
  pas de sélection multi-valeurs sur un champ de page) ; la sélection
  multiple est l'étape suivante prévue (filtrage de la plage source ou
  champ de page + recalcul).
- Les segments ne sont pas persistés dans le fichier .xlsx (les slicers
  Excel sont ignorés à l'ouverture).
