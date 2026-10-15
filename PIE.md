# Secteurs et anneaux

Prototype original en formes libres PowerPoint éditables. Les boutons **Secteurs** et **Anneau** ouvrent le même éditeur ; **Modifier les données** réouvre ses données et réglages après sauvegarde du PPTX. Les exemples [Pie.json](examples/Pie.json) et [Doughnut.json](examples/Doughnut.json) sont synthétiques et peuvent être ouverts avec **Importer JSON**.

## Utilisation

L'onglet **Données** contient les libellés, valeurs, couleurs facultatives `#RRGGBB` et éclatements individuels en pourcentage du rayon. Un collage Excel attend deux colonnes : libellé et valeur, avec ou sans en-tête. Le collage et l'import remplacent les lignes et leurs couleurs/éclatements ; les autres réglages restent conservés. Supprimer une ligne conserve les réglages des lignes survivantes.

L'onglet **Réglages** propose quatre ordres : feuille horaire, feuille antihoraire, valeurs croissantes horaires et valeurs décroissantes horaires. Les égalités gardent l'ordre de la feuille. Couleurs et éclatements restent liés aux indices des données quand l'ordre d'affichage change.

Le départ est exprimé en degrés : `-90` en haut, `0` à droite, progression horaire positive. Le trou est exprimé en pourcentage du diamètre extérieur. Une valeur nulle donne un secteur plein, quel que soit le nom du type. Un titre central nécessite un trou suffisamment grand pour contenir son texte. Les montants affichent valeurs, pourcentages, les deux ou aucun ; les libellés de catégories sont indépendants.

L'onglet **Lien Excel** lit les données enregistrées d'une plage sans en-tête. Utiliser deux colonnes ; les champs Secondaire/Total non neutres sont refusés pour ces types. Le lien enregistré permet les commandes d'actualisation manuelle du ruban. Si le nombre de lignes change à l'actualisation, les styles individuels sont réinitialisés et un avertissement est enregistré ; les autres réglages restent conservés. À nombre de lignes identique, les styles restent liés à leur position dans les données.

## Modèle JSON

`Kind` vaut `Pie` ou `Doughnut`. Les lignes utilisent exclusivement `Value` ; `Series`, `Waterfall` et `Appearance` ne sont pas pris en charge pour ces types. Les paramètres résident dans `Pie` :

| Champ | Valeur / portée |
|---|---|
| `StartAngle` | Degrés finis ; défaut original `-90` |
| `HoleRatio` | `[0,1[` ; absent/null : `0` pour Pie, `0.5` pour Doughnut |
| `Order` | `SheetClockwise`, `SheetCounterclockwise`, `Ascending`, `Descending` |
| `LabelContent` | `Percent`, `Value`, `Both`, `None` |
| `ShowCategory` | Affichage du libellé de catégorie |
| `PercentDigits` | De 0 à 6 décimales ; pourcentages arrondis indépendamment |
| `CenterTitle` | Texte central facultatif |
| `Slices` | Liste de styles à indices de données uniques : `Index`, `Explode` de 0 à 0.5, `Color` facultative |

Les valeurs doivent être finies et non négatives, avec au moins une valeur positive. Les zéros conservent leur ligne dans le modèle et ne produisent pas de forme. La somme utilise des valeurs normalisées pour accepter des nombres très grands sans débordement. Une contribution inférieure à la précision angulaire disponible peut ne pas être dessinée.

## Géométrie et limites

Chaque segment est un contour fermé de Bézier ; le bord intérieur d'un anneau est parcouru en sens inverse. Le trou laisse voir le fond et les objets situés derrière le graphique. Aucun disque blanc ne masque le centre. Les cercles sont approchés par des courbes cubiques.

Les labels externes sont répartis verticalement de chaque côté, avec traits de rappel. Les catégories longues et titres centraux peuvent être abrégés visuellement ; le texte complet reste dans les données. Les montants et pourcentages restent complets. Un cadre trop dense est refusé avant remplacement plutôt que de superposer les labels. Les pourcentages arrondis ne sont pas ajustés pour totaliser exactement 100 %.

Rotation, trou, ordre, couleurs et éclatement se règlent dans l'éditeur. Les poignées de rotation/trou/explosion sur la diapositive, labels intérieurs, déplacement manuel persistant des labels, menus spécifiques, anneaux concentriques et connecteurs universels restent absents. La réédition régénère les formes et ne conserve pas les retouches natives internes. Les limites générales de rotation des groupes, d'Undo/Redo et d'Office restent celles de [PARITY.md](PARITY.md).

## Référence et provenance

Le [manuel public des secteurs et anneaux](https://www.thinkcell.com/en/resources/manual/pie) décrit les quatre ordres, la rotation, l'explosion, le trou réglable, les labels de valeurs/pourcentages et un titre central. Il décrit aussi des gestes de manipulation sur la diapositive qui ne sont pas encore reproduits ici.

L'inspection statique privée de la DLL de référence a trouvé des noms associés à ces états et à leur sérialisation. Des chaînes ne prouvent ni algorithme ni comportement d'exécution. Aucun défaut numérique du produit de référence n'a été établi pour angle, trou ou données négatives ; les valeurs par défaut et refus ci-dessus sont des choix originaux. Aucun code décompilé ni ressource propriétaire n'est inclus. **La parité fonctionnelle, visuelle et d'interaction reste non vérifiée.**

## Validation reproductible

`Build.ps1 -Test` couvre la géométrie pure et les allers-retours de l'éditeur. `Test-Pies.ps1` exige PowerPoint fermé, puis vérifie des présentations synthétiques, leurs formes libres, métadonnées, remplacement transactionnel et sauvegarde/réouverture ; les PNG sont produits dans le dossier indiqué par `-OutputDirectory`. Ces contrôles ne comparent pas le résultat à une exécution de think-cell.
