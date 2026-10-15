# Validation du premier jalon

Tests locaux réalisés le 15 octobre 2026, sous Windows avec PowerPoint/Excel de bureau x64, version Office 16.0.20430.20146.

| Vérification | Résultat | Limite |
|---|---|---|
| Géométrie et validation des entrées | 2 984 assertions réussies | Ce nombre compte les assertions, pas les scénarios |
| Préservation des données dans les éditeurs | 6 assertions réussies | Formulaires instanciés sans parcours interactif |
| Lecteur et convertisseur Python | 37 tests réussis | Fixtures synthétiques, pas un corpus exhaustif |
| Intégration Office avec un graphique converti | 65 assertions réussies, processus de test terminé avec code 0 | Une installation Office x64 locale |
| Chargement COM/ruban | `connected=True; ribbon=True` | Les clics sur tous les boutons restent à tester |
| Inspection des exemples fournis | 32 contrôles privés réussis, hashes inchangés | Trois présentations, cinq modèles de graphiques |

L'intégration vérifie les groupes de formes, les métadonnées, le remplacement transactionnel, la lecture et l'actualisation Excel, les matrices multi-séries et les données après réouverture. Sept variantes multi-séries sont couvertes : colonnes et barres groupées/empilées/100 %, courbe groupée.

Le test optionnel d'import utilise un JSON dérivé d'une présentation de référence locale. Le fichier d'origine, le JSON et le rendu de ce test ne sont pas publiés. Les tests synthétiques peuvent être exécutés sans ces références. Le scénario Office standard comporte 61 contrôles ; l'option `-InputJson` en ajoute quatre.

Les rendus synthétiques et le rendu converti ont été inspectés visuellement. Cela ne constitue pas une comparaison pixel à pixel avec think-cell. Aucun résultat de parité visuelle n'est revendiqué.

## Revue

Une revue indépendante en lecture seule a identifié puis fait corriger : un type de graphique importé en minuscules converti silencieusement en colonnes, une référence de compilation manquante pour le test Office et des combinaisons du convertisseur refusées par le moteur. Les tests de non-régression couvrent la casse du type et les rejets du convertisseur.

## Restant à vérifier

- Parcours réels des éditeurs, collage pendant l'édition d'une cellule et commandes du ruban.
- Undo/Redo, copie interprésentations, modifications manuelles et autres versions d'Office.
- Erreurs de formules Excel et classeurs complexes.
- Formats think-cell non présents dans les trois références et fidélité de leur apparence.
