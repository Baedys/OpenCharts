# Suivi de parité

L'objectif final est la parité fonctionnelle et d'interaction avec think-cell. Un graphique d'apparence similaire ou un test interne réussi ne prouve pas cette parité. Référence initiale fournie : build 14.0.38.718 ; la documentation publique actuelle peut décrire une autre version.

Statuts : **prototype** = réalisation partielle testable ; **absent** = non réalisé ; **non vérifié** = comparaison de référence manquante. Aucun élément n'a encore le statut « parité vérifiée ».

| Domaine | État | Écart restant / critère de fermeture |
|---|---|---|
| Complément Windows PowerPoint | Prototype | Chargement et ruban confirmés sur Office x64 local ; callbacks interactifs, désinstallation et autres configurations à vérifier |
| Office x86 / ARM64 / Mac | Non vérifié / absent | Builds et corpus spécifiques |
| Ruban et insertion | Prototype | Reproduire workflows, placement interactif et raccourcis de référence |
| Menu radial / contextuel | Prototype limité | Menu de cascade réalisé ; menu radial complet et corpus de gestes de référence restants |
| Éditeur de données | Prototype | Matrice jusqu'à huit séries et cellules absentes ; formules et parité des interactions restantes |
| Waterfall | Prototype | Segments multiples, `e`, contraintes exactes, connecteurs/ancrages modifiables et total à gauche réalisés ; poignées natives et cibles physiques de segments réalisées ; création directe des liens et ancrage physique à zéro réalisés ; règles automatiques générales et parité des gestes restantes |
| Colonnes et barres | Prototype | Groupées, empilées, 100 % et séries multiples réalisées ; un profil empilé comparé localement, généralisation et interactions restantes |
| Courbes et aires | Prototype / absent | Courbes multi-séries avec interruption aux données absentes ; dates, aires et combinaisons restantes |
| Secteurs et anneaux | Prototype | Rotation, quatre ordres, trous transparents, explosions et labels externes réalisés ; gestes sur diapositive, labels intérieurs et comparaison de référence restants. Voir PIE.md |
| Dispersion et bulles | Prototype | Axes numériques linéaires/log/inversés, groupes, tailles en aire/diamètre, deux labels, partitions et régressions linéaires réalisés ; gestes, légende de tailles, modèles non linéaires et comparaison de référence restants. Voir SCATTER.md |
| Mekko | Prototype | Nombre arbitraire de segments et axes unitaires |
| Gantt | Prototype | Calendrier, semaines ISO, jours non ouvrés, jalons, dépendances |
| Placement d'étiquettes | Prototype limité | Géométrie liée et décalages importés pour un profil empilé ; placements et traits de rappel originaux pour secteurs/dispersion ; optimisation générale et parité restantes |
| Axes et graduations | Prototype | Bornes persistées importables dans le profil vérifié ; ruptures, échelles communes et secondaires restantes |
| Annotations | Absent | Écarts, taux de croissance, CAGR, accolades et lignes de valeur |
| Couleurs et styles | Prototype limité | Couleurs de thème résolues, polices et formats simples importés ; gras/italique et styles non uniformes refusés pour l'apparence |
| Formes natives éditables | Prototype | Apparence après duplication, copie et ouverture sans complément |
| Persistance de nos graphiques | Prototype | Tests versionnés de migration et corruption |
| Lecture/écriture objets think-cell | Prototype / absent | Inspection CFB/XML, valeurs et profil d'apparence empilé ; écriture propriétaire, fidélité générale et roundtrip absents |
| Lien Excel enregistré | Prototype | Noms de plage, erreurs Excel, tableaux structurés et sources déplacées |
| Lien Excel ouvert | Absent | Données non enregistrées, sélection, reconnexion au classeur |
| Actualisation automatique | Absent | Événements, états hors ligne, conflits et rafraîchissements groupés |
| Undo/Redo | Prototype / couverture partielle | Ancrage zéro et glissement par poignée d’interface annulés/rétablis en une étape dans les fixtures testées. Modification native directe : défaut de deux étapes conservé. Historiques variés et autres commandes restants |
| Redimensionnement | Prototype | Polices et placement identiques à référence ; rotations |
| Conservation des retouches | Absent | Styles, annotations, animations et interactions après mise à jour |
| Disposition automatique | Absent | Contraintes d'alignement, boîtes de texte et ajustements de taille |
| Tableaux / textes liés | Absent | Mise à jour des valeurs et formats depuis Excel |
| Agendas et processus | Absent | Structure, navigation et synchronisation |
| Bibliothèque / recherche | Absent | Inventaire et workflows à cadrer sur version de référence |
| Automatisation par modèle | Absent | Mise à jour JSON/Excel par lot et isolation des éléments hors cible |
| Installation / mise à jour | Prototype | Signature, paquet reproductible et restauration après échec |
| Parité visuelle globale | Non vérifié | Comparaisons de captures à résolution et polices identiques |

## Profil d'apparence observé

La conversion `Column` + `Stacked` accepte un sous-ensemble vérifié : valeurs non négatives, correspondance unique des séries avec le cache OOXML, labels de valeurs centrés, catégories, sommes et noms de séries liés à leurs formes. Palette, ordre d'empilement, tailles de police, formats simples, bornes d'axe, zone de tracé et cadre source sont transférés. Les écarts et hauteurs de labels doivent être uniformes à 0,1 point près. Des transformations de groupes non neutres, styles par point ou références ambiguës entraînent un avertissement et la conversion des données seules.

Un premier comparatif privé à 1 440 × 810 pixels a trouvé les neuf segments dans les couleurs attendues, avec un écart maximal d'un pixel sur leurs boîtes englobantes. Cette mesure ne valide pas la typographie, les interactions ni tous les pixels ; l'alignement du texte et les règles de redimensionnement restent des axes de validation.

Un probe distinct, sous Office 16.0.20430.20146, compare le rendu natif de la référence à deux copies : expliciter `xMode/yMode=factor` sur les neuf labels produit **0 pixel différent** ; mettre leurs décalages à zéro produit **939 pixels différents**. Cette expérience étaye le mode implicite sur ce profil. Elle ne compare pas notre moteur à think-cell et n'établit pas un défaut universel pour tous les graphiques ou toutes les versions Office.

Les modes explicites `factor` sont pris en charge ; les modes absents utilisent cette observation avec un avertissement. Les bornes min/max explicites du cache sont conservées, mais le choix initial « automatique » ou « fixe » n'est pas déduit des codes internes. Aucune ligne du tableau ne devient « parité vérifiée » sur cette seule base.

Les **58 tests Python synthétiques** couvrent la lecture, les refus de conversion, les propriétés du profil, les décalages, les indices uniques et les transformations. Ils ne remplacent pas les comparaisons de référence. Les documents, images et données propriétaires de ces comparatifs ne sont pas distribués.

## Corpus de validation à constituer

Chaque scénario doit enregistrer les entrées, l'action, les sorties numériques, le rendu, la persistance et les limites observées. Inclure : nombres positifs/négatifs/nuls, données manquantes, longues étiquettes, séries nombreuses, petits graphiques, changement de thème, duplication, copie interprésentations, réouverture, Undo/Redo, source Excel absente et fichier déplacé.

Les captures de référence et leurs résultats restent séparés du code distribué. Les observations statiques du binaire ne sont jamais assimilées à une exécution observée.

## Sources de comportement

- [Catalogue graphique](https://www.think-cell.com/en/product/think-cell-charts)
- [Waterfall](https://www.think-cell.com/en/resources/manual/waterfall)
- [Secteurs et anneaux](https://www.thinkcell.com/en/resources/manual/pie)
- [Gantt](https://www.think-cell.com/en/resources/manual/gantt)
- [Gestion des liens Excel](https://www.think-cell.com/en/resources/manual/manage-data-links)
- [Disposition automatique](https://www.think-cell.com/en/resources/manual/automatic-layout)
- [Compatibilité des présentations](https://www.think-cell.com/en/resources/manual/compatibility)
- [API d'automatisation](https://www.think-cell.com/en/resources/manual/api)

## Cascades à contraintes

Le nouveau modèle résout des graphes explicites avec rationnels internes et conserve `e` comme inconnue. Quatre dispositions initiales sont proposées dans l’éditeur ; leur choix reste explicite. Les scénarios numériques issus du manuel sont couverts, mais l'édition des connecteurs sur la diapositive reste partielle et la topologie automatique générale n'est pas reproduite. Voir [WATERFALL.md](WATERFALL.md).
