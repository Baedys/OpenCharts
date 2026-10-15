# Validation du jalon apparence et interopérabilité

Tests locaux réalisés le 15 octobre 2026, sous Windows avec PowerPoint/Excel de bureau x64, version Office 16.0.20430.20146.

| Vérification | Résultat | Limite |
|---|---|---|
| Géométrie et validation des entrées | 3 098 assertions réussies | Ce nombre compte les assertions, pas les scénarios |
| Préservation des données dans les éditeurs | 15 assertions réussies | Formulaires instanciés sans parcours interactif |
| Lecteur et convertisseur Python | 58 tests réussis | Fixtures synthétiques, pas un corpus exhaustif |
| Intégration Office avec un graphique converti | 181 assertions réussies, processus terminé avec code 0 | Une installation Office x64 locale |
| Chargement COM/ruban | `connected=True; ribbon=True` | Les clics sur tous les boutons restent à tester |
| Inspection des exemples fournis | 32 contrôles privés réussis, hashes inchangés | Trois présentations, cinq modèles de graphiques |
| Bibliothèque de types native | 905 déclarations de fonctions et 11 GUIDs comparés, zéro divergence | Déclarations, pas 905 fonctionnalités utilisateur ; voir INTEROP.md |

L'intégration vérifie les groupes de formes, les métadonnées, le remplacement transactionnel, la lecture et l'actualisation Excel, les matrices multi-séries et les données après réouverture. Sept variantes multi-séries sont couvertes : colonnes et barres groupées/empilées/100 %, courbe groupée. Un cas synthétique de 220 × 104 points vérifie le cadre importé et les positions/dimensions de chaque zone de texte, avant et après sauvegarde/réouverture. Les textes sans retour automatique à la ligne conservent leur géométrie.

Le test optionnel d'import utilise un JSON dérivé d'une présentation de référence locale. Le fichier d'origine, le JSON et le rendu de ce test ne sont pas publiés. La suite standard comporte 105 assertions ; le profil de référence testé avec `-InputJson` en ajoute 76. Le nombre ajouté dépend du nombre de textes du graphique. L'égalité des métadonnées compare les modèles sérialisés après désérialisation commune : l'ordre des propriétés d'un objet JSON n'est pas significatif.

## Comparaison visuelle limitée

Une présentation de référence contenant un graphique empilé a été rendue par PowerPoint, puis comparée au groupe de formes reconstruit avec ses données et son apparence importées. Sur les deux PNG de 1440 × 810 pixels, les neuf rectangles correspondant aux trois couleurs de série ont un écart maximal de **1 pixel sur leurs bords**. Les couleurs de remplissage sont identiques. Cette mesure porte sur les segments ; elle ne mesure ni toutes les différences de texte ni toutes les interactions. Les deux images ont également été inspectées visuellement.

Une expérience séparée a vérifié les décalages des étiquettes du profil : rendre explicites les modes `factor` absents du fichier a produit zéro pixel différent ; remettre les neuf décalages à zéro a changé 939 pixels. Cela justifie l'interprétation de ces champs pour ce profil et cette version d'Office seulement. Le convertisseur conserve un avertissement correspondant.

Ces observations portent sur le rendu PowerPoint des éléments enregistrés. Aucun exécutable ou complément think-cell n'a été installé ni exécuté pour ces vérifications. Elles ne démontrent pas une parité globale avec le produit.

## Revue et corrections

La revue indépendante en lecture seule a conduit à corriger les débordements hors cadre, le support des petits cadres explicites, le transfert des décalages d'étiquettes, la validation des indices de rôles et le refus des transformations ou styles typographiques non pris en charge. Les changements de type sont validés aux dimensions réelles avant fermeture de l'éditeur. Un petit graphique peut conserver son profil importé ; passer à la disposition standard exige d'abord un groupe d'au moins 320 × 240 points.

Le collage valide les données avant de changer la grille ; l'application valide ensuite le modèle final avec son apparence conservée. Les tests d'éditeur couvrent ce collage dans un petit cadre, le remappage des couleurs et décalages après suppression d'une série, la remise à zéro de l'échelle après modification numérique et les changements de type.

## Restant à vérifier

- Parcours réels des éditeurs, collage pendant l'édition d'une cellule et commandes du ruban.
- Undo/Redo, copie interprésentations, modifications manuelles et autres versions d'Office.
- Erreurs de formules Excel et classeurs complexes.
- Formats think-cell absents des références, placement automatique et fidélité de leur apparence.
- Comportement d'exécution des interfaces natives ; les résultats statiques sont détaillés dans INTEROP.md.
