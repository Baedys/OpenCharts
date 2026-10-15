# Suivi de parité

L'objectif final est la parité fonctionnelle et d'interaction avec think-cell. Un graphique d'apparence similaire ou un test interne réussi ne prouve pas cette parité. Référence initiale fournie : build 14.0.38.718 ; la documentation publique actuelle peut décrire une autre version.

Statuts : **prototype** = réalisation partielle testable ; **absent** = non réalisé ; **non vérifié** = comparaison de référence manquante. Aucun élément n'a encore le statut « parité vérifiée ».

| Domaine | État | Écart restant / critère de fermeture |
|---|---|---|
| Complément Windows PowerPoint | Prototype | Chargement et ruban confirmés sur Office x64 local ; callbacks interactifs, désinstallation et autres configurations à vérifier |
| Office x86 / ARM64 / Mac | Non vérifié / absent | Builds et corpus spécifiques |
| Ruban et insertion | Prototype | Reproduire workflows, placement interactif et raccourcis de référence |
| Menu radial / contextuel | Absent | Corpus d'interactions sélection, poignées et propriétés |
| Éditeur de données | Prototype | Matrice jusqu'à huit séries et cellules absentes ; formules et parité des interactions restantes |
| Waterfall | Prototype | Totaux automatiques `e`, empilement, connecteurs modifiables, total à gauche |
| Colonnes et barres | Prototype | Groupées, empilées, 100 % et séries multiples réalisées ; comparaison visuelle et interactions restantes |
| Courbes et aires | Prototype / absent | Courbes multi-séries avec interruption aux données absentes ; dates, aires et combinaisons restantes |
| Secteurs et anneaux | Absent | Angles, trous, labels et petits segments |
| Dispersion et bulles | Absent | Axes numériques, surface des bulles, quadrants, régressions |
| Mekko | Prototype | Nombre arbitraire de segments et axes unitaires |
| Gantt | Prototype | Calendrier, semaines ISO, jours non ouvrés, jalons, dépendances |
| Placement d'étiquettes | Absent | Optimisation anticollision, déplacements manuels conservés, leaders |
| Axes et graduations | Prototype | Graduations lisibles, ruptures, échelles communes et secondaires |
| Annotations | Absent | Écarts, taux de croissance, CAGR, accolades et lignes de valeur |
| Couleurs et styles | Prototype | Styles de document, palette, polices et formats numériques |
| Formes natives éditables | Prototype | Apparence après duplication, copie et ouverture sans complément |
| Persistance de nos graphiques | Prototype | Tests versionnés de migration et corruption |
| Lecture/écriture objets think-cell | Prototype / absent | Inspection CFB/XML et conversion limitée des valeurs ; écriture, fidélité visuelle et roundtrip absents |
| Lien Excel enregistré | Prototype | Noms de plage, erreurs Excel, tableaux structurés et sources déplacées |
| Lien Excel ouvert | Absent | Données non enregistrées, sélection, reconnexion au classeur |
| Actualisation automatique | Absent | Événements, états hors ligne, conflits et rafraîchissements groupés |
| Undo/Redo | Non vérifié | Plusieurs opérations puis annulation/rétablissement sans perte |
| Redimensionnement | Prototype | Polices et placement identiques à référence ; rotations |
| Conservation des retouches | Absent | Styles, annotations, animations et interactions après mise à jour |
| Disposition automatique | Absent | Contraintes d'alignement, boîtes de texte et ajustements de taille |
| Tableaux / textes liés | Absent | Mise à jour des valeurs et formats depuis Excel |
| Agendas et processus | Absent | Structure, navigation et synchronisation |
| Bibliothèque / recherche | Absent | Inventaire et workflows à cadrer sur version de référence |
| Automatisation par modèle | Absent | Mise à jour JSON/Excel par lot et isolation des éléments hors cible |
| Installation / mise à jour | Prototype | Signature, paquet reproductible et restauration après échec |
| Parité visuelle globale | Non vérifié | Comparaisons de captures à résolution et polices identiques |

## Corpus de validation à constituer

Chaque scénario doit enregistrer les entrées, l'action, les sorties numériques, le rendu, la persistance et les limites observées. Inclure : nombres positifs/négatifs/nuls, données manquantes, longues étiquettes, séries nombreuses, petits graphiques, changement de thème, duplication, copie interprésentations, réouverture, Undo/Redo, source Excel absente et fichier déplacé.

Les captures de référence et leurs résultats restent séparés du code distribué. Les observations statiques du binaire ne sont jamais assimilées à une exécution observée.

## Sources de comportement

- [Catalogue graphique](https://www.think-cell.com/en/product/think-cell-charts)
- [Waterfall](https://www.think-cell.com/en/resources/manual/waterfall)
- [Gantt](https://www.think-cell.com/en/resources/manual/gantt)
- [Gestion des liens Excel](https://www.think-cell.com/en/resources/manual/manage-data-links)
- [Disposition automatique](https://www.think-cell.com/en/resources/manual/automatic-layout)
- [Compatibilité des présentations](https://www.think-cell.com/en/resources/manual/compatibility)
- [API d'automatisation](https://www.think-cell.com/en/resources/manual/api)
