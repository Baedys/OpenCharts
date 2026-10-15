# Validation du jalon connecteurs sur diapositive

Tests locaux réalisés les 15–16 octobre 2026, sous Windows avec PowerPoint/Excel de bureau x64, version Office 16.0.20430.20146.

| Vérification | Résultat | Limite |
|---|---|---|
| Géométrie et validation des entrées | 3 822 assertions réussies | Ce nombre compte les assertions, pas les scénarios |
| Solveur de cascades | 21 scénarios, 169 assertions réussies | Six scénarios numériques du manuel ; topologie explicitement définie |
| Géométrie et mutations des connecteurs | 12 scénarios, 363 assertions réussies | Cibles physiques, ambiguïtés, contraintes manuelles et rollback ; pas de gestes UI |
| Création de connecteurs et ancrages zéro | 14 scénarios, 66 assertions réussies | Politique de reconstruction explicite ; pas de priorité native revendiquée |
| Pont d’événements COM | 55 contrôles gérés réussis | Cycle de vie testé sans Office ; transport réel couvert séparément |
| Minuteur de thread | 16 contrôles réussis | STA local pompé, sans HWND ; exécution Office couverte séparément |
| Poignées d’interface | 34 contrôles WinForms locaux réussis | Capture, géométrie, owner, annulation et destruction ; pas de gestes Office |
| Projection écran | 32 contrôles réussis | Échelles synthétiques et coordonnées négatives ; pas de garantie multi-écran |
| Clavier d’édition | 22 scénarios, 146 assertions réussies | Routage injecté, cycle de vie et hook réel sur STA isolé ; pas de touches Office injectées par ces tests |
| Interactions Office | 53 assertions réussies, processus terminé avec code 0 | Formes natives et événements COM ; pas de glissement de souris injecté |
| Préservation des données dans les éditeurs | 25 assertions réussies | Formulaires instanciés sans parcours interactif |
| Lecteur et convertisseur Python | 58 tests réussis | Fixtures synthétiques, pas un corpus exhaustif |
| Intégration Office avec un graphique converti | 399 assertions réussies, processus terminé avec code 0 | Une installation Office x64 locale |
| Chargement COM/ruban | `connected=True; ribbon=True` | Les clics sur tous les boutons restent à tester |
| Inspection des exemples fournis | 32 contrôles privés réussis, hashes inchangés | Trois présentations, cinq modèles de graphiques |
| Bibliothèque de types native | 905 déclarations de fonctions et 11 GUIDs comparés, zéro divergence | Déclarations, pas 905 fonctionnalités utilisateur ; voir INTEROP.md |

Les poignées d’interface ajoutent 212 contrôles locaux (34 fenêtre, 32 projection, 146 clavier). Le parcours intégré vérifie un cycle Undo/Redo complet sur le total à gauche, l’annulation sans édition du choix Base/Cumul, son choix Base puis Undo, le menu natif depuis la poignée et le masquage/rétablissement après zoom. Le PPTX sauvegardé pendant une session contient un groupe, trois tags de connecteurs de cascade et aucun tag auxiliaire `OSC_EDIT`. Le XML compte dix objets `cxnSp`, car axes et grille emploient aussi des lignes natives. Les détails et les limites de ces gestes figurent dans [UI-VALIDATION.md](UI-VALIDATION.md). La suite COM de 53 assertions a aussi réussi pendant cette intégration ; ses actions utilisent les formes natives, pas les entrées de la souris.

L'intégration vérifie les groupes de formes, les métadonnées, le remplacement transactionnel, la lecture et l'actualisation Excel, les matrices multi-séries et les données après réouverture. Sept variantes multi-séries sont couvertes : colonnes et barres groupées/empilées/100 %, courbe groupée. Un cas synthétique de 220 × 104 points vérifie le cadre importé et les positions/dimensions de chaque zone de texte, avant et après sauvegarde/réouverture. Les textes sans retour automatique à la ligne conservent leur géométrie. Trois cascades synthétiques ajoutent des segments multiples, des `e`, un total à gauche, des signes mixtes et un connecteur libre oblique. La suite vérifie le recalcul après remplacement, le refus transactionnel d’ancrages incompatibles et la conservation des données, contraintes et zones de texte après réouverture.

Le test optionnel d'import utilise un JSON dérivé d'une présentation de référence locale. Le fichier d'origine, le JSON et le rendu de ce test ne sont pas publiés. La suite standard comporte 323 assertions ; le profil de référence testé avec `-InputJson` en ajoute 76. Le nombre ajouté dépend du nombre de textes du graphique. L'égalité des métadonnées compare les modèles sérialisés après désérialisation commune : l'ordre des propriétés d'un objet JSON n'est pas significatif.

## Comparaison visuelle limitée

Une présentation de référence contenant un graphique empilé a été rendue par PowerPoint, puis comparée au groupe de formes reconstruit avec ses données et son apparence importées. Sur les deux PNG de 1440 × 810 pixels, les neuf rectangles correspondant aux trois couleurs de série ont un écart maximal de **1 pixel sur leurs bords**. Les couleurs de remplissage sont identiques. Cette mesure porte sur les segments ; elle ne mesure ni toutes les différences de texte ni toutes les interactions. Les deux images ont également été inspectées visuellement.

Une expérience séparée a vérifié les décalages des étiquettes du profil : rendre explicites les modes `factor` absents du fichier a produit zéro pixel différent ; remettre les neuf décalages à zéro a changé 939 pixels. Cela justifie l'interprétation de ces champs pour ce profil et cette version d'Office seulement. Le convertisseur conserve un avertissement correspondant.

Ces observations portent sur le rendu PowerPoint des éléments enregistrés. Aucun exécutable ou complément think-cell n'a été installé ni exécuté pour ces vérifications. Elles ne démontrent pas une parité globale avec le produit.

## Cascades

Le calcul exact a été vérifié sur six exemples numériques du manuel et sur des cas supplémentaires : sommes indépendantes, calcul arrière, `e` avec segments numériques, ancrages sur segments, résidu de connecteur libre, valeurs décimales et extrêmes. Les images et données de référence restent privées. Les sorties numériques des exemples ne constituent pas une validation de la topologie automatique ni des interactions. Les trois rendus PowerPoint synthétiques ont été inspectés visuellement.

La revue a fait corriger les collisions de lignes effectives entre anciens indices implicites et nouveaux `SheetRow` explicites, qui pouvaient écraser un segment dans la grille. Ces entrées sont refusées avant ouverture de l’éditeur. Le collage simple pendant l’édition d’une cellule est laissé au contrôle texte ; seuls les tableaux sont interceptés. Les tests d’éditeur couvrent les données et contraintes, le recalcul, les IDs stables, les cellules absentes et le refus d’un collage numérique invalide sans modification. Aucun parcours complet au clavier ou à la souris n’est revendiqué.

## Connecteurs sur la diapositive

`Test-Interactions.ps1` charge le complément réel puis modifie les connecteurs natifs par le modèle objet PowerPoint. Il vérifie la sélection d'un enfant de groupe, le rattachement d'une extrémité, le recalcul de `e` sans figer sa valeur, les ancrages automatiques, le recalcul automatique par minuteur, la sauvegarde/réouverture et l'absence de formes auxiliaires. Il couvre aussi la suppression native d'un connecteur, la navigation immédiate vers une autre diapositive, la suppression du groupe ou de la diapositive et l'entrée dans un ancien graphique sans tags sémantiques. Les 53 assertions couvrent aussi création depuis des segments, ancrages physiques à zéro, annulation et sauvegarde d'une commande incomplète, navigation, suppression, refus d'un autre graphique avec mêmes IDs et Échap dans les menus. Elles ne représentent pas 53 gestes de souris ni une comparaison d'interaction avec think-cell.

La revue a fait corriger la provenance des ancrages : modifier dans la grille un ancrage créé par un profil le rend manuel, pour éviter son retrait lors d'une réparation automatique ultérieure. L'activation des anciens groupes et la régénération protègent les callbacks de sélection réentrants. Les appels COM externes restent sur le STA du complément avec `StandardOleMarshalObject`. Le minuteur Win32 de thread remplace les tentatives d'accès à `HWND`, refusées par IDispatch sur l'hôte testé.

Les probes de sauvegarde ont observé `BeforeSave` pendant `Save()`, mais aucun callback pendant les appels COM `SaveAs` et `SaveCopyAs` testés. Aucun auxiliaire d'édition n'est donc inséré dans le document. Une sauvegarde automatisée avant validation d'un déplacement peut toujours conserver une géométrie non synchronisée : terminer d'abord la session. Voir [INTERACTION.md](INTERACTION.md).

Les parcours Windows supplémentaires ont validé un déplacement de poignée avec recalcul de −9 à +9, l'activation et la fin de session depuis le sous-menu natif OpenCharts, l'affichage d'une ambiguïté Base/Cumul, sa restauration par Échap et un choix explicite de Base donnant `e = 3`. Ils restent distincts des assertions COM automatisées. Les commandes contextuelles sont enregistrées avec le Ribbon XML Office 2009 pour ne pas dépendre exclusivement de `WindowBeforeRightClick`. Les limites et la procédure reproductible sont consignées dans [UI-VALIDATION.md](UI-VALIDATION.md).

La revue des commandes directes a fait corriger la vérification d'identité du graphique destination et l'annulation par Échap bref dans les sous-menus. Un premier passage Office a révélé une destruction du menu trop tôt pendant `Closed` ; elle est désormais différée après le traitement WinForms. Les appels de fermeture clavier et de traitement d'Échap imbriqué font partie des contrôles Office.

## Revue et corrections

La revue indépendante en lecture seule a conduit à corriger les débordements hors cadre, le support des petits cadres explicites, le transfert des décalages d'étiquettes, la validation des indices de rôles et le refus des transformations ou styles typographiques non pris en charge. Les changements de type sont validés aux dimensions réelles avant fermeture de l'éditeur. Un petit graphique peut conserver son profil importé ; passer à la disposition standard exige d'abord un groupe d'au moins 320 × 240 points.

Le collage valide les données avant de changer la grille ; l'application valide ensuite le modèle final avec son apparence conservée. Les tests d'éditeur couvrent ce collage dans un petit cadre, le remappage des couleurs et décalages après suppression d'une série, la remise à zéro de l'échelle après modification numérique et les changements de type.

## Restant à vérifier

- Étendre la validation Undo/Redo des poignées d’interface au-delà du total à gauche et corriger le repli par modification native directe, qui conserve son défaut d’incohérence après un seul Undo. Voir [UI-VALIDATION.md](UI-VALIDATION.md).
- Parcours réels des éditeurs, collage pendant l'édition d'une cellule et commandes du ruban.
- Généralisation des gestes au-delà des trois scénarios de [UI-VALIDATION.md](UI-VALIDATION.md), autres choix et annulations, glissements prolongés et écrans multiples.
- Undo/Redo, copie interprésentations, modifications manuelles et autres versions d'Office.
- Erreurs de formules Excel et classeurs complexes.
- Formats think-cell absents des références, placement automatique et fidélité de leur apparence.
- Comportement d'exécution des interfaces natives ; les résultats statiques sont détaillés dans INTEROP.md.
