# Interactions PowerPoint : provenance et validation

## Connecteurs natifs

L'éditeur agit sur les véritables connecteurs PowerPoint du groupe graphique. Le rôle `WATERFALL_CONNECTOR`, l'index de primitive `OSC_ITEM` et l'identifiant du connecteur dans le modèle permettent de retrouver chaque forme. La session conserve également l'identifiant PowerPoint de la forme ; un nom identique ne suffit pas à identifier un connecteur.

L'activation sélectionne un connecteur et expose ses poignées natives. Un déplacement d'extrémité recherche une base, un cumul ou un bord de segment voisin, puis modifie les contraintes du modèle. Les cellules `e` restent des inconnues calculées. Une cible ambiguë ouvre un choix explicite. Le déplacement des deux extrémités en même temps est refusé. La suppression d'un connecteur met à jour le modèle et recalcule les sommes.

Un minuteur Win32 de thread (`SetTimer(NULL, 0, 200, callback)`) observe la géométrie sur le STA Office. Il ne dépend pas d'un HWND de document. Après détection d'un changement, deux observations stables supplémentaires permettent de l'appliquer automatiquement. Le maintien du bouton gauche suspend l'analyse automatique ; après relâchement, la stabilité est vérifiée. Ce mécanisme ne représente pas un événement Office de fin de glissement et ne prouve pas encore tous les gestes interactifs. Le bouton d'application force aussi l'analyse du déplacement courant.

Les menus natifs Office des groupes, formes et connecteurs incluent un sous-menu OpenCharts filtré sur la sélection. Le ruban est invalidé après les changements de sélection, les actions d'édition et la fin avant sauvegarde. Un menu géré reste également accessible par le callback `WindowBeforeRightClick` lorsque l'hôte le livre. Le menu contextuel propose l'édition des données, l'édition des connecteurs, leur suppression et la fin de session. L'éditeur exige un groupe sans rotation, directement sur la diapositive. La parité complète du menu, de la rotation, des conflits et des gestes de think-cell n'est pas établie.

Aucune forme temporaire de poignée n'est ajoutée au document. Une première version fondée sur des connecteurs auxiliaires a été abandonnée après observation de leur persistance lors d'un `SaveAs` automatisé. Ses tests de nettoyage ont été retirés du livrable : ils ne valident pas cette architecture.

## Origine du pont d'événements

`src/PowerPointEvents.cs` est une implémentation C# originale utilisant `System.Runtime.InteropServices.ComEventsHelper`. Elle ne copie aucun code décompilé, actif graphique ou bibliothèque propriétaire. Elle ne nécessite pas de référence de compilation à l'assembly PIA PowerPoint.

Les signatures ont été vérifiées indépendamment dans les métadonnées publiques installées : l'assembly Microsoft.Office.Interop.PowerPoint 15.0 et la bibliothèque Office16 `MSPPT.OLB`, chargée avec `LoadTypeLibEx(REGKIND_NONE)`. Cette lecture de métadonnées ne lance pas PowerPoint. L'interface source est `EApplication`, IID `914934C2-5A91-11CF-8700-00AA0060263B`.

| Événement Office | DISPID | Callback du pont |
| --- | ---: | --- |
| WindowSelectionChange | 2001 | `Action<object>` |
| WindowBeforeRightClick | 2002 | `BeforeRightClick(object, ref bool)` |
| PresentationBeforeSave | 2018 | `BeforePresentationAction(object, ref bool)` |
| PresentationBeforeClose | 2025 | `BeforePresentationAction(object, ref bool)` |
| WindowDeactivate | 2010 | `Action<object, object>` |

Le booléen par référence conserve le contrat d'annulation Office. Les délégués restent fortement référencés jusqu'au détachement. Une erreur d'abonnement déclenche un rollback ; ses erreurs éventuelles restent visibles. `Dispose` doit être appelé sur le thread de création, bloque les callbacks tardifs et conserve les abonnements dont le détachement échoue pour permettre une nouvelle tentative. Le pont ne libère pas le RCW d'application appartenant à son appelant.

La classe COM `Connect` dérive de `StandardOleMarshalObject` pour que les appels COM externes passent par le STA de création du complément. Le marshaler CLR par défaut laisse les objets gérés exposés à COM être appelés depuis d'autres threads ; cette différence peut créer un minuteur sans boucle de messages et rendre ses observations inactives. Le contrat de thread reste vérifié lors du démarrage, de l'arrêt et de la libération du minuteur.

Les changements de sélection peuvent résulter du code lui-même : le contrôleur protège ses opérations contre les callbacks réentrants. Les événements de sauvegarde et de fermeture terminent la session ; une erreur remontée à ces callbacks peut annuler l'opération concernée. Ce mécanisme n'est pas une garantie couvrant toute écriture de fichier.

## Sauvegardes : observation et limite

Un probe distinct, exécuté sur l'Office installé le 15 octobre 2026, a enregistré les callbacks suivants. La vérification statique confirme également que `PresentationSave` porte le DISPID 2005 et reçoit seulement une présentation, sans argument d'annulation. Cet événement a été observé par le probe ; il n'est pas un abonnement supplémentaire du pont actuel.

| Appel COM testé | Callbacks observés pendant l'appel |
| --- | --- |
| `SaveAs(path, 24)` | Aucun |
| `Save()` après marquage du document comme modifié | 2018, puis 2005 |
| `SaveCopyAs(path, 24)` | Aucun |
| `Close()` | 2025 |

Ces observations caractérisent les appels et l'installation testés, pas toutes les versions d'Office ni les commandes du menu Fichier. Le probe relève les callbacks synchrones ; il ne démontre pas l'absence universelle de callbacks différés. Les contrôles d'interaction séparés confirment le transport des événements de sélection et la livraison de `BeforeSave` lors de `Save()`.

Les documents Microsoft décrivent `PresentationBeforeSave` avec une annulation possible et mentionnent l'apparition du dialogue Enregistrer sous. Ils ne garantissent pas sa livraison lors d'un `SaveAs` direct. `PresentationSave` est décrit comme précédant l'enregistrement, mais sa signature ne permet pas d'annuler celui-ci. La documentation de `SaveCopyAs` ne promet pas de notification. L'architecture n'utilise donc aucun de ces événements comme condition nécessaire à l'élimination de formes auxiliaires.

Une sauvegarde automatisée peut encore intervenir avant la validation d'un déplacement par le minuteur. L'absence de formes auxiliaires évite leur persistance ; elle ne prouve pas, à elle seule, la synchronisation immédiate de chaque modification géométrique avec les métadonnées. Avant une sauvegarde automatisée, appeler **Terminer les connecteurs** (ou `EndConnectorEditing(null)` depuis le client COM) puis sauvegarder. Les sauvegardes après validation, puis réouverture, font l'objet de tests distincts.

L'identifiant du minuteur est attribué par Windows. L'arrêt utilise cet identifiant exact et conserve une référence forte au callback natif. Les messages déjà en file après l'arrêt sont ignorés lorsqu'aucune inscription courante ne leur correspond. Si Windows réutilise un identifiant, un ancien message peut avancer une observation de la nouvelle session ; il ne rappelle pas un ancien délégué retiré. Les ticks servent à observer l'état courant, pas à mesurer une durée exacte. La boucle de messages Office reste nécessaire.

## Portée des tests

`tests/PowerPointEventsTests.cs` contient 55 vérifications gérées reproductibles sans Office. Il se compile avec `src/PowerPointEvents.cs` afin d'accéder au constructeur injectable interne. Il vérifie les identifiants, les callbacks, la transmission du booléen par référence, l'identité des délégués, le rollback, les erreurs de détachement, les nouvelles tentatives, la suppression des callbacks tardifs et la contrainte de thread. Il ne prouve pas à lui seul le transport COM réel ni le déclenchement des événements par PowerPoint.

`tests/PowerPointTimerTests.cs` vérifie le minuteur de thread sur un STA local pompé : 16 contrôles couvrent arrêt/reprise, durée de vie après GC, callbacks retirés, erreurs, contrainte de thread et isolation de deux minuteurs. Aucune fenêtre ni instance Office n’est nécessaire. Ces contrôles ne prouvent pas que la boucle de messages d’Office distribue les ticks ; le test d’interaction réel doit le vérifier.

`tests/InteractionOffice.cs`, exécuté par `Test-Interactions.ps1`, couvre le transport COM et les formes natives dans PowerPoint : sélection d'un enfant de groupe, rattachement d'extrémité avec recalcul de `e`, conservation de son caractère calculé, sauvegarde/réouverture, absence de formes auxiliaires et suppression d'un connecteur avec nouvelle sommation. Le passage final a validé 34 assertions sur Office x64 16.0.20430.20146 : le total initial passe de −9 à +9 sur application explicite, puis à +5 automatiquement après stabilisation. Il couvre également `SaveCopyAs` pendant une session déjà synchronisée, `Save()` avec fin de session, navigation avant stabilisation sans perte ni retour forcé vers l'ancienne diapositive, suppression du groupe ou de la diapositive, redémarrage après suppression et migration des anciens groupes sans tags de rôle. Les positions sont modifiées via le modèle objet PowerPoint et les commandes du complément via son objet COM : ce test n'injecte pas un glissement réel de souris. Ces appels vérifient aussi la visibilité des commandes contextuelles pour un groupe, un connecteur et une forme ordinaire, ainsi que la suppression par le callback du menu. Un parcours Windows séparé a exercé le glissement réel, l'activation et la fin depuis le menu natif, un choix d'ancrage ambigu et son annulation par Échap ; voir [UI-VALIDATION.md](UI-VALIDATION.md).

Ces probes ne prouvent pas une reproduction exhaustive de l'interface de think-cell. Ils isolent le contrat d'événements Office, la logique du pont et les comportements de reconstruction effectivement exercés.

## Sources Microsoft

- [SetTimer](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-settimer)
- [KillTimer](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-killtimer)
- [StandardOleMarshalObject](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.standardolemarshalobject)
- [Managed and Unmanaged Threading in Windows](https://learn.microsoft.com/en-us/dotnet/standard/threading/managed-and-unmanaged-threading-in-windows)
- [ComEventsHelper](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.comeventshelper)
- [WindowSelectionChange](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.windowselectionchange)
- [WindowBeforeRightClick](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.windowbeforerightclick)
- [PresentationBeforeSave](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationbeforesave)
- [PresentationBeforeClose](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationbeforeclose)
- [WindowDeactivate](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.windowdeactivate)
- [PresentationSave](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentationsave)
- [SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas)
