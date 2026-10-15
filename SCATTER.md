# Dispersion et bulles

Prototype original en formes PowerPoint éditables. **Dispersion** et **Bulles** ouvrent le même éditeur ; **Modifier les données** réouvre les points, axes et annotations du groupe sélectionné. Exemples synthétiques : [Scatter.json](examples/Scatter.json), [Bubble.json](examples/Bubble.json).

## Données et conversion

Chaque ligne représente un point : **Libellé, X, Y, Taille, Groupe, Couleur, Marqueur**. Le collage accepte trois à sept colonnes, avec ou sans en-tête. Les cellules numériques vides restent absentes, distinctes de zéro. Une taille renseignée sélectionne automatiquement `Bubble` ; retirer toutes les tailles rétablit `Scatter`.

Les points sans X ou Y restent enregistrés sans marqueur. En mode bulles, une taille absente ou nulle ne produit pas de bulle ; les tailles négatives sont refusées. Une ligne avec taille positive sans X/Y peut servir de référence commune d'échelle. Ces règles de données manquantes/négatives sont des choix originaux, sauf l'emploi documenté d'une ligne de référence de taille.

Les groupes utilisent une palette stable suivant leur première apparition dans les données. La couleur facultative `#RRGGBB` surcharge celle du groupe. Les marqueurs acceptent `Circle`, `Square`, `Triangle`, `Diamond`, `Plus`, `Cross`, `None`. Les bulles sont circulaires ; une conversion depuis une dispersion réinitialise les formes incompatibles et conserve un avertissement dans le modèle.

En mode `Area`, le diamètre vaut `MaxBubbleDiameter × sqrt(Size / taille maximale)` ; en mode `Diameter`, il vaut `MaxBubbleDiameter × Size / taille maximale`. La taille maximale inclut les lignes de référence sans coordonnées. Les grandes bulles sont dessinées avant les petites. Aucun diamètre minimal artificiel n'est appliqué ; les très petites formes restent soumises à la précision native d'Office.

## Axes et étiquettes

Les axes X/Y ont des bornes indépendantes, automatiques ou fixes, une option logarithmique et un sens inversé. La plage automatique comporte une marge originale ; ses graduations ne reproduisent pas un algorithme établi de think-cell. Les points hors bornes fixes restent dans le modèle sans tracé. Sur axe logarithmique, une coordonnée non positive est placée à la borne basse, avec avertissement visible ; au moins une coordonnée positive dessinable est requise pour construire cet axe. Les bornes explicites logarithmiques doivent être positives.

Deux contenus d'étiquettes sont proposés : `None`, `Label`, `X`, `Y`, `XY`, `Size`. L'affichage automatique est désactivé au-delà de 300 lignes de points. Le placement choisit parmi des positions proches et ajoute des traits de rappel lorsque nécessaire. Les catégories longues peuvent être abrégées ; les nombres restent complets selon leur format d'affichage. Le texte original reste dans les données. Ce placement peut conserver des collisions dans les configurations denses ; il ne reproduit pas le moteur de référence.

## Tendances et partitions

Les tendances disponibles sont linéaires, pour tous les groupes ou un groupe précis. Le calcul normalise les valeurs et utilise des rationnels exacts en interne pour limiter les pertes numériques ; les coefficients exportés sont arrondis en `double`. Dans les bulles, les poids sont les valeurs brutes de Taille. Cette interprétation du poids est un choix original : la documentation dit taille des bulles sans établir ici une équivalence entre valeur brute, aire et diamètre. Une tendance sans données suffisantes ou pente normalisée représentable n'est pas dessinée.

R² et équation sont facultatifs et indépendants de la présence de la droite. En cas de coefficient brut non représentable, l'équation utilise les coordonnées normalisées. Les tailles des bulles ne sont pas utilisées comme poids dans une dispersion.

Une partition est une droite infinie définie par deux points distincts en coordonnées de données, puis coupée au cadre de tracé. Tendances et partitions exigent deux axes linéaires. Les valeurs se modifient dans leurs onglets dédiés ; déplacer une forme native ne met pas encore ces paramètres à jour.

## Excel et persistance

Le lien enregistré accepte une plage contiguë sans en-tête, de 1 à 1 000 lignes et de trois à sept colonnes dans l'ordre ci-dessus. Le complément lit les valeurs enregistrées dans une instance Excel séparée, avec macros et liens externes désactivés. L'ajout/retrait de la colonne Taille peut convertir le type lors de l'actualisation.

Si le nombre de points ne change pas, les couleurs/marqueurs absents des colonnes Excel restent liés à leur position. Si le nombre change, ces styles non fournis sont réinitialisés avec avertissement. Les colonnes de styles explicitement présentes remplacent les valeurs précédentes, y compris les cellules vides. Le collage remplace entièrement les points et leurs styles. Axes, options et annotations restent conservés.

Les modèles sont enregistrés dans les tags `OSC_*` du groupe, avec un plafond de **1 Mio de JSON UTF-8**. Les groupes de 1 000 points dépassent l'ancien plafond de tags ; le lecteur conserve la même structure de tags avec un budget borné plus grand. Une erreur de validation ou de rendu conserve le groupe précédent. Les retouches natives internes sont perdues lors d'une régénération.

## Modèle JSON

`ChartSpec.Kind` vaut explicitement `Scatter` ou `Bubble` ; les points résident dans `Scatter.Points`, sans `Rows`, `Series`, `Appearance`, `Pie` ni `Waterfall`. L'éditeur et Excel déduisent le type ; un JSON explicitement `Scatter` contenant une taille est refusé pour éviter une interprétation silencieuse.

`Scatter` porte `XAxis`/`YAxis` (`Minimum`, `Maximum`, `Logarithmic`, `Reversed`), `XTitle`, `YTitle`, `ShowGrid`, `ShowLegend`, `MarkerSize`, `MaxBubbleDiameter`, `BubbleSizing`, `MarkerScheme`, `ShowLabels`, `LabelContent`, `SecondLabelContent`, `LabelFontSize`, `Trendlines` et `Partitions`. Les exemples fournissent une configuration minimale.

Une tendance contient `Group` (`null` = tous les groupes ; `""` = sans groupe), `Color`, `ShowRSquared`, `ShowEquation`. Une partition contient `X1`, `Y1`, `X2`, `Y2`, `Label`, `Color`. Le regroupement respecte la casse et les espaces ; `"A"` et `" A"` restent distincts.

## Limites et provenance

Restent absents : poignées directes de points et de tailles, légende des tailles de bulles, annotations interactives, dates, tendances non linéaires, contraintes persistantes de placement manuel et conservation des retouches. L'intégration Excel n'est pas automatique et ne lit pas les données non enregistrées. Les limites générales d'Office et de rotation des groupes restent décrites dans [PARITY.md](PARITY.md).

Sources primaires : [manuel dispersion/bulles](https://www.thinkcell.com/en/resources/manual/scatterbubble), [axes](https://www.thinkcell.com/en/resources/manual/axes), [conversion des types](https://www.thinkcell.com/en/resources/kb/0096), [échelle de tailles commune](https://www.thinkcell.com/en/resources/kb/0145), [moindres carrés pondérés NIST](https://itl.nist.gov/div898/handbook/pmd/section4/pmd432.htm). Les lectures statiques privées de noms natifs sont des pistes, sans preuve d'algorithme ni d'exécution. Aucun code décompilé ou asset propriétaire n'est inclus. **Parité fonctionnelle, visuelle et d'interaction non vérifiée.**

Validation reproductible : `Build.ps1 -Test`, puis `Test-Scatter.ps1` avec PowerPoint fermé. La suite native produit un PPTX synthétique, des PNG et des mesures dans `test-results/scatter` ; elle n'exécute pas think-cell.
