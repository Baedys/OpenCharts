# Parcours souris et menus contextuels

Vérifications effectuées le 15 octobre 2026 sur PowerPoint Windows x64 16.0.20430.20146. Les graphiques sont synthétiques, créés par OpenCharts. Les gestes utilisent le contrôle Windows Computer Use ; les valeurs et contraintes sont ensuite relues par le modèle objet PowerPoint. Ces parcours ne sont pas une comparaison d'exécution avec think-cell. Les fichiers JSON fournis reprennent les données et contraintes des trois fixtures utilisées ; leur import par dialogue n'a pas été exercé pendant ces gestes.

## Secteurs et anneaux — 16 octobre 2026

Sur le profil Office x64 local à 200 %, le bouton **Modifier les données** du ruban a ouvert l'éditeur d'un anneau synthétique. L'inspection a fait corriger une mise à l'échelle trop tôt dans la construction WinForms, des en-têtes de grille rognés et un contrôle de décimales invisible. Le formulaire final expose les données et tous les champs de réglages sans rognage dans ce parcours.

Le champ du trou est passé de 50 à 60 %. **Appliquer** a régénéré l'anneau ; la relecture du modèle donne `HoleRatio=0.6`. Un Ctrl+Z a rétabli le modèle initial sans options explicites et le trou de 50 % ; un Ctrl+Y a rétabli modèle et trou de 60 %. La présentation synthétique a été sauvegardée et exportée, puis fermée. Aucun document utilisateur n'a été modifié.

Ce parcours ciblé ne constitue pas une couverture de tous les réglages, du collage interactif ni des gestes de référence think-cell. Les 12 fixtures et 1 684 assertions de `tests/PieOffice.cs` couvrent séparément le rendu natif et la sauvegarde/réouverture, sans entrées souris.

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

## Création directe, ancrage à zéro et Échap imbriqué

La troisième fixture correspond à [WaterfallIndependentSums.json](examples/WaterfallIndependentSums.json) : deux colonnes numériques 3 et 5, sans connecteur, avec deux ancrages de base automatiques à zéro. Le graphique a été créé par le harness, puis les commandes suivantes ont été déclenchées uniquement par gestes Windows :

1. Sélection du segment C0 ; clic droit → **OpenCharts → Ajouter un connecteur depuis… → Sommet du segment**.
2. Sélection du segment C1 → **Pied du segment → Égaliser les niveaux**. Le menu de réparation affiche séparément les bases de C0 et C1, ainsi que **Conserver les ancrages**.
3. Choix de l'ancrage automatique de C1. Le graphique affiche les sorties 3 et 8. La relecture retrouve un lien C0 `[1]` → C1 `[0]`, `Constrain=true`, et seulement P0 à zéro. Les données restent 3 et 5 ; P1 est retiré.
4. Sur C1, **OpenCharts → Placer le sommet à zéro**. Sa sortie devient zéro et son pied −5 ; C0 va de −8 à −5. La relecture conserve le même lien et les deux données, retire l'ancrage automatique restant et retrouve un ancrage manuel C1 `[1]` à zéro.
5. Nouvel ajout depuis le sommet de C1, sélection de C0, ouverture du sous-menu **Pied du segment** contenant Égaliser/Libre, puis frappe brève d'Échap. Tous les menus se ferment. Relecture : `connector-editing=False`, `connector-adding=False`, un seul lien, mêmes données et même ancrage manuel qu'avant cette tentative.

Le parcours vérifie donc le callback réel des items, le remplacement d'un menu de destination par un menu de réparation, le choix explicite d'une réparation et l'annulation depuis un sous-menu imbriqué. La destruction des menus est différée après leur fermeture : un premier test COM avait détecté une `ObjectDisposedException` lorsqu'ils étaient détruits synchroniquement dans `Closed`.

Les variantes **Connecteur libre**, création depuis Pied/Cumul, placement du pied à zéro, choix de l'autre réparation et **Conserver les ancrages** n'ont pas été exercées par gestes dans cette fixture. Les tests de modèle couvrent leur politique de contraintes ; cela ne remplace pas leurs futurs parcours UI.

## Undo/Redo : réussite limitée et défaut confirmé

Essais complémentaires des 15–16 octobre 2026 sur la même installation Office, code du commit `9a9976a`. Fixtures et activation de l'éditeur préparées par le harness COM ; glissement, Ctrl+Z et Ctrl+Y réalisés avec les entrées Windows réelles.

Sur les deux sommes indépendantes, **Placer le sommet à zéro** sur C0 remplace P0 par une ancre manuelle C0 `[1]`. Un Ctrl+Z restaure P0/P1 automatiques, les valeurs 3 et 5 et les sorties initiales. Un Ctrl+Y rétablit l'ancre manuelle avec son même ID et le rendu attendu. La session se termine lors de l'annulation et reste inactive après rétablissement. Ce test couvre une commande, sans séquence de plusieurs éditions.

Sur le total à gauche `e ; 3 ; 2 ; 4`, déplacer la poignée droite du premier lien vers la sortie de C3 produit `e = +9`. Un premier Ctrl+Z rétablit le modèle initial, `e = −9`, mais restaure l'ancienne ligne telle qu'elle était juste avant régénération : sa poignée droite est encore sur C3. Le rendu et les contraintes sont donc incohérents. Un second Ctrl+Z rétablit le connecteur initial vers la base de C1. Deux Ctrl+Y successifs reproduisent cet état intermédiaire puis le résultat cohérent `e = +9`, ancre droite C3 `[1]`. Ce cycle complet a été vérifié sur le code publié ; les données numériques et la définition calculée de `e` sont conservées.

**Ce défaut reste ouvert pour les modifications natives directes.** Le nouveau parcours par poignées d’interface décrit ci-dessous évite cette première modification native. Dans l’ancien parcours, terminer les deux étapes restaure la cohérence ; cela ne garantit pas un nombre fixe d'étapes pour tout historique Office. Ne pas enregistrer un état intermédiaire incohérent. Aucun Undo automatique supplémentaire n'est lancé par le complément : il pourrait annuler une autre action de l'utilisateur.

Une variante privée supprimant `StartNewUndoEntry` avant la régénération a présenté le même défaut et n'est pas retenue. La [documentation Microsoft](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.startnewundoentry) décrit la création d'une entrée, sans contrat de fusion rétroactive. L'interface native et les formes seules ne suffisent donc pas à prouver une annulation atomique. Les 53 assertions COM restent valides dans leur périmètre ; elles ne testaient pas ces frappes Undo/Redo.

## Limites de la preuve

Les clics droits automatisés initiaux ont ouvert le menu Office sans callback `WindowBeforeRightClick`, aussi bien via `ComEventsHelper` que via un abonnement PIA Microsoft séparé et pompé. Les clics gauches ont livré `WindowSelectionChange`. Ces observations ne déterminent pas si l'origine tient à l'injection d'entrée, au contexte Office ou à un autre facteur ; elles ne prouvent pas que l'événement serait universellement absent. Le menu Ribbon XML fournit donc les commandes sans dépendre de sa livraison. Le complément ne branche plus l'ancien menu géré sur cet événement ; la méthode publique est conservée comme API obsolète.

Un premier glissement a été refusé par l'outil parce qu'un autre processus recouvrait le point de départ. Après activation de PowerPoint et nouvelle capture, le glissement a réussi. Ce refus n'est pas compté comme un déplacement validé. Un incident du helper Computer Use a été résolu en réinitialisant sa session JavaScript avant reprise.

Ces résultats couvrent trois petites cascades et une installation Office. Ils ne valident pas le menu radial think-cell, les gestes prolongés, toutes les annulations, Undo/Redo, les présentations concurrentes, les écrans multiples ou toutes les versions d'Office.


## Poignées d’interface : annulation cohérente dans deux fixtures

Validation du 16 octobre 2026, même Office x64, écran Windows à 200 %. Le complément charge maintenant une fenêtre transparente et non activante au-dessus du connecteur sélectionné. Les points visibles sont une interface WinForms ; le PPTX conserve seulement ses formes de graphique. Les essais suivants utilisent le complément intégré, après un prototype privé distinct.

- **Total à gauche, `e ; 3 ; 2 ; 4`** : déplacement de la poignée droite vers la sortie de C3, `e = +9`, ancre droite C3 `[1]`. Un Ctrl+Z restaure `e = −9`, ancre C1 `[0]` et les coordonnées initiales de la ligne. Un Ctrl+Y restaure `e = +9` et la nouvelle ligne. Les IDs de formes avant/après annulation correspondent aux deux groupes attendus. La session se termine à Undo et reste inactive après Redo.
- **Cible coïncidente, `e ; 3 ; 0 ; 4`** : le dépôt vers C2 ouvre Base/Cumul/Annuler. Pendant le menu, `e = −7`, les trois liens, l’ancre C1 `[0]`, les IDs et les coordonnées des lignes restent inchangés. Échap ferme le choix ; mêmes éléments conservés, compteur d’éditions inchangé, session et poignées rétablies. Le choix explicite Base donne ensuite `e = 3`, ancre droite C2 `[0]`. Un Ctrl+Z restaure en une étape `e = −7`, C1 `[0]` et la ligne initiale. Redo de ce second cas n’a pas été exercé.
- **Menu natif sur une poignée** : le clic droit ouvre le menu Office contenant OpenCharts. Échap ferme ce menu sans terminer la session ni modifier le modèle. La transmission utilise les messages de bouton attendus par Office ; `WM_CONTEXTMENU` seul n’avait pas ouvert le menu.
- **Clic immobile** : clic sur le bord d’une poignée, compteur d’éditions et IDs des lignes inchangés. Le seuil de mouvement utilise les coordonnées de MouseDown, indépendamment de la position ultérieure du curseur.
- **Reprise après Undo** : une première version échouait à lire `GroupItems.Count` sur la référence venant de la sélection restaurée (`0x800A01A8`). Le démarrage relit désormais la forme depuis la diapositive propriétaire et son ID. Sur le code corrigé, déplacement, Undo, réactivation, SaveAs puis nouvelle réactivation ont conservé `e = −9`, C1 `[0]`, les IDs et coordonnées des trois lignes, avec poignées visibles. Une initialisation échouée efface son état local sans réinterroger le groupe invalide ni masquer l’erreur d’origine. Les traces de diagnostic temporaires sont retirées du livrable.
- **Zoom** : à 111 %, poignées alignées avec les extrémités. À 120 %, cadre coupé à droite : poignées d’interface masquées, session et modèle conservés. À 15 %, graphique entièrement visible : poignées réapparues aux positions reprojetées, sans édition. Les quatre coins et les volets sont contrôlés avec une politique conservatrice ; le glissement sur un graphique partiellement visible n’est pas pris en charge par cette interface.

Ces parcours ne valident pas une annulation atomique pour toute commande ou tout historique. Le repli natif par COM, les poignées PowerPoint accessibles quand l’interface est masquée, les gestes sur plusieurs écrans et les autres versions Office restent des limites. La frappe réelle de Suppr, Échap pendant le maintien du bouton, les champs de volets et toutes les variantes de fermeture n’ont pas été exercés par ces nouveaux gestes. Les 34 contrôles locaux de fenêtre, 32 de projection et 146 assertions de routage clavier complètent ces essais sans remplacer leur validation Office. La revue a corrigé le focus clavier, le masquage après zoom, les clics sans mouvement et le menu de poignée ; aucun nouveau défaut concret lors du dernier passage statique.
