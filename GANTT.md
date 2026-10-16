# Gantt avec calendrier

Le bouton **Gantt** ouvre un calendrier daté, rendu en formes PowerPoint éditables. [GanttCalendar.json](examples/GanttCalendar.json) fournit un exemple synthétique. Les anciens graphiques `Gantt` à `Rows` numériques restent lisibles et utilisent leur ancien éditeur en jours relatifs.

## Calendrier et objets

Dans **Calendrier**, saisir les bornes inclusives `AAAA-MM-JJ`. Le mode automatique choisit des échelles suivant la place disponible ; le mode manuel conserve années, trimestres, mois, semaines et jours choisis. Une échelle manuelle trop serrée est refusée. Les seuils de sélection et la disposition sont originaux, sans comparaison de référence.

Les week-ends peuvent rester ordinaires, être ombrés ou disparaître du calendrier. Les jours de week-end et le premier jour de semaine se règlent séparément. Les numéros de semaine utilisent ISO 8601 ; avec un début de semaine autre que lundi, le numéro correspond à sa première date affichée, un choix original faute de contrat précis. Les périodes d'ombre ajoutées ne compressent jamais le calendrier.

Une activité peut contenir plusieurs barres, pentagones, chevrons, jalons et crochets. Les colonnes Responsable et Remarque, le retrait des activités et leur couleur se règlent dans l'éditeur. La rangée permanente en bas accepte jalons et crochets. Les objets hors plage gardent leurs données et libellés pour une réapparition ultérieure.

Les pentagones et chevrons contiennent leur texte ; leur hauteur grandit pour les lignes nécessaires. Les dates éventuelles restent sous la forme. Les hauteurs ne sont pas encore harmonisées entre tous les objets du graphique ; ce placement et le contraste du texte sont des choix originaux.

Deux modes de dates sont distincts :

- `InclusiveDays` : dates entières, début à minuit et fin à la fin du jour indiqué. Une barre du 1er au 1er couvre une journée.
- `Exact` : instants ISO sans fuseau, fin exclusive. Du `2025-01-01T12:00:00` au `2025-01-02` couvre une demi-journée. Les heures ne sont pas arrondies à des journées.

Un jalon en jours entiers apparaît au milieu de sa journée si l'échelle jours est visible, sinon à son début. Les jalons exacts conservent leur heure source : la fidélité de toutes les combinaisons de jalons ancrés n'est pas établie. Quand un objet tombe uniquement sur des week-ends masqués, sa forme disparaît mais son libellé reste près de la frontière comprimée ; cette politique originale reste à comparer.

Les barres coupées par le cadre portent des pointes de débordement. **Cadrer sur les objets** ajuste la plage à tous les objets enregistrés, sans inclure les ancres inutilisées. Les connecteurs entre objets sont visuels et suivent les dates lors d'une régénération. Ils n'imposent aucun calcul de planning.

## Excel et ancres

Le lien Excel attend une plage contiguë sans en-tête : première ligne réservée aux jalons, même vide ; première colonne pour le libellé d'activité ; colonnes suivantes pour des dates. Les lignes ordinaires commencent à la deuxième ligne. L'import crée **libellés et ancres, aucun objet implicite**. Dans **Objets**, choisir une ancre de début ou de fin et le mode Instants exacts pour relier une barre ou un jalon.

Lors d'une actualisation, les identités d'activité et d'ancre restent liées à leur position et à la colonne de date. Les objets et leurs options sont conservés. Une date référencée disparue ou un changement du nombre d'activités avec objets existants fait refuser l'import entier. Cette restriction de changement de structure est originale ; elle évite une réaffectation silencieuse.

Les systèmes numériques Excel 1900 et 1904 sont distingués par le classeur. Le 29 février 1900 fictif est refusé. Le collage de tableau accepte des dates ISO ; les valeurs numériques Excel sont prises en charge par le lien au classeur. Seules les valeurs enregistrées sont lues, avec macros et mises à jour de liens externes désactivées.

## Modèle et limites

`ChartSpec.Gantt` contient `StartDate`, `EndDate`, `Culture`, `ScaleMode`, `Scales`, `WeekendDays`, `WeekendDisplay`, `WeekStartsOn`, `Activities`, `MilestoneRow`, `Items`, `DateAnchors`, `Shades` et `Links`. Ce modèle ne se combine pas avec `Rows`, `Series`, `Appearance` ni les modèles d'autres graphiques. Les IDs sont enregistrés pour conserver les références ; l'éditeur propose les activités, objets et ancres par leur libellé.

Les objets portent `Kind`, `RowId`, `StartDate`, `EndDate`, `DateMode`, `StartAnchor`, `EndAnchor`, `Label`, `Color` et `ShowDates`. Une ancre contient `Id`, `RowId`, `Index` et `Date`. Un lien contient `FromId`, `ToId`, `FromEdge` et `ToEdge` (`Start`/`End`). Les dates des ancres remplacent les dates explicites de leurs extrémités.

Limites originales : calendrier grégorien, plage de 36 600 jours, 100 activités ordinaires, 1 000 objets, 32 dates par ligne Excel. Les petites dimensions et densités incompatibles avec des textes lisibles sont refusées. Les textes peuvent être abrégés visuellement ; les originaux restent enregistrés. Le placement utilise des couloirs, sans reproduire le placement manuel de think-cell.

Formats disponibles : `y`/`yy`/`yyyy`, `q`/`qq`/`qqq`, `m` à `mmmmm`, `d` à `ddddd`, `w` à `wwww`, littéraux entre apostrophes. `qq`/`qqq` donnent les trimestres romains minuscules/majuscules ; `y` est une extension originale. Les échappements et formats de référence ne sont pas tous reproduits.

Restent absents : poignées de dates et ancrages entre objets sur diapositive, déplacements liés, date du jour dynamique, calendriers fiscaux, regroupements de libellés sur plusieurs activités, dispositions manuelles persistantes et suivi des retouches natives. Le lien Excel reste manuel. Le calendrier complet, les interactions et Undo/Redo ne sont pas déclarés équivalents à think-cell.

## Sources et validation

Sources primaires : [manuel Gantt actuel](https://www.thinkcell.com/en/resources/manual/gantt), [bornes et jalons du manuel espagnol](https://www.thinkcell.com/es/resources/manual/gantt), [connecteurs de dépendance](https://www.thinkcell.com/en/resources/kb/0155), [apparence du calendrier depuis v14](https://www.thinkcell.com/en/resources/kb/0244). Les versions linguistiques divergent sur les libellés masqués ; le modèle suit la conservation décrite dans le manuel anglais actuel.

Les observations statiques privées de noms natifs ne prouvent aucun algorithme ni comportement d'exécution. Aucune DLL, ressource ou portion de code propriétaire n'est distribuée. Les choix du moteur sont originaux. Tests reproductibles : `Build.ps1 -Test`, puis `Test-Gantt.ps1` avec PowerPoint fermé. Résultats et périmètres : [VALIDATION.md](VALIDATION.md) et [UI-VALIDATION.md](UI-VALIDATION.md). **Aucune parité globale n'est revendiquée.**
