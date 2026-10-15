# Parcours souris et menus contextuels

Vérifications effectuées le 15 octobre 2026 sur PowerPoint Windows x64 16.0.20430.20146. Les graphiques sont synthétiques, créés par OpenCharts. Les gestes utilisent le contrôle Windows Computer Use ; les valeurs et contraintes sont ensuite relues par le modèle objet PowerPoint. Ces parcours ne sont pas une comparaison d'exécution avec think-cell. Les fichiers JSON fournis reprennent les données et contraintes des deux fixtures utilisées ; leur import par dialogue n'a pas été exercé pendant ces gestes.

## Total à gauche

Importer [WaterfallLeftTotal.json](examples/WaterfallLeftTotal.json), puis sélectionner le graphique et choisir **Éditer les connecteurs**. Les données sont `e ; 3 ; 2 ; 4`, avec un départ de `e` à zéro et la dernière sortie à zéro. Le résultat initial est `e = −9`.

Déplacer la poignée droite du premier connecteur vers le bord gauche de la dernière colonne, au niveau zéro. Après relâchement et stabilisation, le modèle conserve une cellule calculée et donne `e = +9`. Le premier connecteur relie désormais la sortie de C0 à la sortie de C3. Ce déplacement a été réalisé à la souris et son résultat relu dans les métadonnées ; il ne repose pas sur une modification des coordonnées par COM.

## Menu natif Office

Sur un groupe sélectionné, ouvrir le menu contextuel puis **OpenCharts**. Les commandes **Modifier les données** et **Modifier les connecteurs** ont été observées. La seconde active les poignées natives.

Sur un connecteur sélectionné pendant une session, le sous-menu affiche aussi **Supprimer ce connecteur** et **Terminer les connecteurs**. L'action de fin a été exécutée depuis ce menu. Après réactivation sur un nouveau graphique, les commandes liées à une session inactive restent masquées. Les tests COM séparés vérifient la suppression et l'absence des commandes OpenCharts sur une forme ordinaire ; la suppression par clic de menu n'a pas été exercée dans ce parcours souris.

Les commandes sont ajoutées aux menus Office `ContextMenuObjectsGroup`, `ContextMenuShape` et `ContextMenuShapeConnector`. Ces IDs proviennent du classeur officiel [PowerPoint Controls](https://github.com/OfficeDev/office-fluent-ui-command-identifiers/blob/main/Microsoft%20365/Current%20Channel/powerpointcontrols.xlsx), SHA-256 `8a62d1f0479b867458c33111ed88234b986c1ac3d3868ec519453c0a71c99309`. La structure suit le [format contextMenus de Microsoft](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/db5b22ab-d58e-459b-81a2-b9e51c23e9cc). Les IDs sont utilisés sans redistribuer le classeur.

## Choix et annulation d'une ambiguïté

Importer [WaterfallAmbiguousAnchor.json](examples/WaterfallAmbiguousAnchor.json). Les données sont `e ; 3 ; 0 ; 4` ; le résultat initial est `e = −7`. Activer les connecteurs depuis le sous-menu OpenCharts.

Déplacer la poignée droite du premier connecteur sur le bord gauche de C2, au niveau du segment de valeur zéro. Sa base et son cumul coïncident géométriquement, mais leurs ancres symboliques diffèrent. Le menu de choix affiche **Colonne 2 — Base**, **Colonne 2 — Cumul** et **Annuler le déplacement**.

Appuyer sur Échap a fermé ce menu, restauré la ligne initiale et conservé `e = −7`, les trois connecteurs et l'ancre droite du premier lien sur C1 avec poids `[0]`. Dans ce parcours, la sélection a également été effacée et la session s'est terminée.

Après réactivation, le même déplacement suivi du choix **Colonne 2 — Base** a produit `e = 3`. La métadonnée droite du premier lien est devenue C2 avec poids `[0]` ; les autres liens sont conservés. La session est restée active. Le choix **Cumul**, le clic sur **Annuler le déplacement**, l'annulation par clic extérieur et les modifications numériques après chaque choix restent à vérifier par des gestes distincts.

## Limites de la preuve

Les clics droits automatisés initiaux ont ouvert le menu Office sans callback `WindowBeforeRightClick`, aussi bien via `ComEventsHelper` que via un abonnement PIA Microsoft séparé et pompé. Les clics gauches ont livré `WindowSelectionChange`. Ces observations ne déterminent pas si l'origine tient à l'injection d'entrée, au contexte Office ou à un autre facteur ; elles ne prouvent pas que l'événement serait universellement absent. Le menu Ribbon XML fournit donc les commandes sans dépendre de sa livraison. L'ancien menu géré reste disponible lorsque cet événement est livré.

Un premier glissement a été refusé par l'outil parce qu'un autre processus recouvrait le point de départ. Après activation de PowerPoint et nouvelle capture, le glissement a réussi. Ce refus n'est pas compté comme un déplacement validé. Un incident du helper Computer Use a été résolu en réinitialisant sa session JavaScript avant reprise.

Ces résultats couvrent deux petites cascades et une installation Office. Ils ne valident pas le menu radial think-cell, les gestes prolongés, toutes les annulations, Undo/Redo, les présentations concurrentes, les écrans multiples ou toutes les versions d'Office.
