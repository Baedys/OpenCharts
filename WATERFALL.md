# Cascade : comportement de référence et limites

Cette spécification provient du manuel et des illustrations livrés avec le package de référence 14.0.38.718. Elle décrit les faits documentés et les choix de reconstruction proposés. Elle ne certifie pas la parité des interactions ni l'implémentation de chaque comportement dans l'interface actuelle. Aucun graphique, image ou jeu de données propriétaire n'est distribué ici.

## Données, inconnues et connexions

Une colonne peut contenir plusieurs segments signés et au plus un segment `e`. Ce symbole représente une inconnue calculée pour satisfaire les contraintes du graphique ; sa valeur doit être recalculée après modification des données ou des connecteurs. Ce n'est pas une valeur de total saisie et figée.

Les segments positifs et négatifs conservent leur étendue. La progression utilise leur somme algébrique : son niveau de sortie peut donc traverser l'intérieur d'un segment au lieu de correspondre à son sommet ou son pied. Il faut distinguer l'origine du calcul, la sortie signée et les bords physiques des segments.

Un connecteur sélectionné possède deux poignées. Le déplacement d'une poignée change sa connexion et provoque un réarrangement du graphique ainsi que le recalcul des `e`. Supprimer un connecteur commence une nouvelle sommation ; le menu contextuel permet d'en ajouter. Des connexions contradictoires peuvent donner des connecteurs obliques. Leur priorité de résolution n'est pas établie par les documents disponibles.

Un segment peut être forcé sur la ligne de base par son sommet ou son pied. Les étiquettes affichent par défaut l'étendue positive du segment, même si la donnée est négative ; un format numérique peut rendre le signe visible. L'ordre des segments est une propriété distincte des valeurs.

## Dispositions initiales limitées

Les règles suivantes constituent des profils de reconstruction compatibles avec les exemples documentés. Elles ne prétendent pas décrire l'algorithme général d'insertion de think-cell. On note `B_i` l'origine signée de la colonne et `O_i = B_i + somme(valeurs_i)` sa sortie.

### Chaîne commençant par une colonne numérique

L'origine de la première colonne est ancrée à zéro. Chaque colonne numérique suivante prend pour origine la sortie précédente. Une colonne contenant seulement `e` est ancrée à zéro et sa sortie rejoint le niveau précédent ; elle représente alors un sous-total. La colonne suivante reprend ce niveau.

Ce profil explique les exemples de sous-totaux intermédiaires et finaux, de colonnes numériques à signes mixtes et de nouvelle sommation après suppression d'un connecteur. Il ne définit pas automatiquement les ancres d'une colonne où `e` partage l'espace avec d'autres segments.

### `e` à gauche, suivi d'incréments positifs

Dans le cas initial documenté, la colonne `e` part de la ligne de base tandis que le sommet de la dernière colonne positive y est ancré. Les connexions successives relient `O_i` à `B_(i+1)`. Le `e` initial devient alors l'opposé de la somme des incréments : sa présence à gauche ne crée pas automatiquement un total positif.

Pour obtenir le total positif à gauche, l'interaction documentée déplace l'extrémité droite du premier connecteur vers le sommet de la dernière colonne. Une reconstruction équivalente conserve l'ancrage du `e`, relie sa sortie à la sortie finale et ancre l'origine de la première colonne numérique à zéro. Les connexions entre les colonnes numériques sont conservées.

### Cas non déterminés par les seules cellules

Le manuel indique qu'un calcul peut commencer par `e` et être résolu depuis la droite. Un autre exemple réutilise un graphique existant, efface ses données puis en saisit de nouvelles pour obtenir le même dessin. Cela prouve la possibilité du calcul arrière, mais pas une règle universelle de choix des connecteurs après effacement ou changement de signe.

Une première colonne `e` et une dernière colonne numérique ne suffisent donc pas à déterminer une topologie unique. Une dernière valeur connue n'est pas automatiquement un total absolu. Les composants sans ancrage suffisant, les colonnes mêlant `e` et segments numériques, et les connexions contradictoires exigent des contraintes explicites ou un diagnostic d'indétermination.

## Conséquences pour la reconstruction

Conserver séparément les données, les identifiants stables des colonnes et segments, les rôles des extrémités de connecteurs et les ancrages à la ligne de base. Un profil initial peut créer ces contraintes ; une modification numérique ne doit pas remplacer silencieusement une topologie explicitement éditée.

Le solveur doit distinguer un système déterminé, un système indéterminé et des contraintes incompatibles. Une valeur manquante ne devient pas zéro. La gestion des conflits doit rester explicite tant que la priorité des contraintes et le placement des connecteurs obliques ne sont pas établis.

Les points encore inconnus comprennent le suivi des extrémités après changement de signe, le choix automatique d'ancrage dans les cas complexes, les contraintes conservées lors d'un effacement de données, et les interactions de rotation ou de retournement. La validation d'un solveur sur des équations explicites ne prouve pas ces comportements d'interface.
## Réalisation actuelle dans OpenCharts

Le bouton **Cascade** crée désormais une cascade à segments et contraintes. L'ancien modèle à une seule valeur et cases Total reste lisible pour les graphiques déjà enregistrés. Le nouveau modèle conserve les valeurs numériques, les cellules `e`, les identifiants stables, les connecteurs et les ancrages dans les métadonnées du groupe PowerPoint. Les `e` sont recalculés à chaque rendu ; leur résultat ne remplace pas leur définition.

L'éditeur présente les catégories en colonnes et jusqu'à huit lignes de segments. Le collage attend une première ligne de catégories et une première colonne de noms de lignes ; une cellule vide représente un segment absent. Double-cliquer sur un en-tête renomme une catégorie. Les lignes de segments absentes restent identifiées par `SheetRow`, afin de conserver leur ordre et leur couleur entre colonnes.

Les onglets **Connecteurs** et **Ancrages** permettent de choisir une colonne et son niveau Base ou Cumul, d'ajouter et retirer des contraintes, et d'ancrer un niveau à une valeur. L'option Égaliser impose un niveau commun. Décochée, elle conserve un connecteur libre, qui apparaît oblique et rouge si les niveaux diffèrent. Les contraintes contradictoires sont signalées ; aucune priorité automatique arbitraire n'est appliquée. Après suppression d'un lien, ajouter un ancrage à zéro pour démarrer une somme indépendante.

Quatre dispositions explicites recréent les liens et ancrages en conservant les données :

- **Cumul vers droite** : départ numérique à gauche, accumulation vers droite.
- **Cumul vers gauche** : départ numérique à droite, accumulation vers gauche. Ce choix est une disposition explicite, pas une prétention de détection automatique de la référence.
- **Total à gauche** : première colonne avec `e`, cumul numérique partant de zéro à la seconde colonne, sortie finale reliée au total.
- **Fin à zéro** : première colonne avec `e`, cumul vers droite dont la dernière sortie est ancrée à zéro.

Ces dispositions ancrent la base des colonnes contenant `e` à zéro et relient leur sortie au cumul concerné. Si `e` partage sa colonne avec d'autres valeurs, il complète cette somme selon ces contraintes explicites. Le déplacement sur la diapositive permet aussi de rattacher une extrémité au sommet ou au pied physique d'un segment ; la grille expose Base et Cumul.

Le solveur utilise des rationnels internes fondés sur les représentations décimales des entrées. Il distingue les contraintes incompatibles des systèmes indéterminés, puis convertit la solution vers les coordonnées numériques du rendu. Il accepte jusqu'à 80 colonnes, huit segments par colonne, un `e` par colonne et 1 024 connecteurs/ancrages par collection. La taille persistable reste limitée par le stockage des métadonnées ; un graphique trop volumineux est refusé avant remplacement du précédent.

Le rendu utilise des formes PowerPoint éditables, des piles positives/négatives divergentes, des étiquettes d'étendue avec signes optionnels et des connecteurs en tirets. Les poignées natives des connecteurs sont utilisables dans le mode d'édition sur la diapositive. Le menu radial complet, la priorité des conflits, le placement automatique des étiquettes et les règles de rotation restent à réaliser. Les liens Excel du nouveau modèle ne sont pas encore disponibles ; utiliser le collage de matrice. Les ancrages pondérés disponibles dans le modèle JSON conservent leurs segments identifiés lors de l'édition ; la grille expose seulement Base et Cumul pour créer de nouvelles extrémités ; les cibles physiques supplémentaires sont accessibles sur la diapositive.

## Édition sur la diapositive

Sélectionner le groupe ou un connecteur, puis **Éditer les connecteurs**. Les formes existantes du graphique servent à l'édition ; aucun objet auxiliaire n'est ajouté. Déplacer une seule extrémité à la fois, vers une base, un cumul ou un bord de segment. Une cible distante de plus de 14 points est refusée. Plusieurs ancrages distincts au même emplacement ouvrent un choix explicite. **Appliquer le déplacement** force la vérification ; **Terminer les connecteurs** termine la session.

La sélection d'un autre connecteur valide d'abord le précédent. Un clic hors du graphique ou un changement de diapositive termine la session et tente de valider le déplacement. Un groupe déplacé, redimensionné ou tourné pendant l'édition termine la session. La suppression native du connecteur sélectionné recalcule le modèle.

La règle de réparation des sommes est volontairement bornée : détacher une base numérique peut créer un ancrage automatique à zéro et retirer les ancrages de sortie automatiques à zéro devenus incompatibles dans sa chaîne numérique. Les ancrages manuels ne sont jamais retirés par cette règle. Les autres topologies doivent déjà être déterminées ou sont refusées avec restauration du modèle précédent. Il s'agit d'un choix de reconstruction, pas d'une priorité interne de think-cell observée en exécution.

Les ancrages créés par un profil portent `AutoGenerated`. Modifier leur niveau ou leur ancrage dans la grille les rend manuels. Une ancienne présentation sans ce champ est traitée de manière conservatrice : ses ancrages sont manuels. Les modèles contraints d'anciens builds sans tags sémantiques sont régénérés lors de l'entrée dans le mode d'édition. Les anciennes cascades à lignes simples doivent rester dans leur éditeur de données.

Le groupe est régénéré après validation : les retouches et animations de ses enfants ne sont pas préservées. Les sauvegardes automatisées doivent terminer l'édition avant `SaveAs` ou `SaveCopyAs` pour éviter une géométrie modifiée encore non reflétée dans les métadonnées. Voir [INTERACTION.md](INTERACTION.md).
