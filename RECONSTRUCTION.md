# Journal de reconstruction

## Cible et méthode

Référence locale fournie : `setup_think-cell_38718.exe`, version 14.0.38.718, 166 442 072 octets. Signature Authenticode vérifiée valide lors de l'inspection, signataire think-cell Operations GmbH. SHA-256 : `c8e24510f977f0c48516d6f7e62f04f96d6169557aaf662711f82687ee1f9ebf`.

L'installateur n'a pas été exécuté. Inventaire PE réalisé avec REA ; payload MSI extrait par plage de fichier, puis tables Windows Installer ouvertes en lecture seule. Les fichiers de référence restent hors du dépôt.

## Observations confirmées

1. Ressource PE RCDATA 201 : offset 1 276 160, longueur 165 146 624. L'en-tête correspond à un fichier composé OLE ; le lecteur Windows Installer confirme un MSI think-cell 14.0.38.718.
2. Tables MSI : variantes du complément pour x86, x64 et ARM64. Enregistrement Office COM `thinkcell.addin` / `thinkcell.addin.1`, CLSID `{D52B1FA2-1EF8-4035-9DA6-8AD0F40267A1}`, modèle de thread Apartment.
3. DLL x64 : 54 388 920 octets, PE32+, machine x86-64, SHA-256 `9377725e887b6b66ac688ac3994eb73c8138120a794203140f443483a4362eef`. Aucune entrée CLI : aucune métadonnée .NET/CIL disponible.
4. Exports identifiés par lecture PE : `DllGetClassObject`, `DllCanUnloadNow`, un export ordinal non nommé. Imports COM/OLE et fonctions graphiques Windows. Ces imports prouvent des dépendances, pas leur usage dans un workflow donné.
5. Inventaire des ressources de la DLL : 180 ressources ; absence de ressource de type nommé TYPELIB dans cet inventaire. Les payloads non décodés restent opaques.
6. Le MSI contient son propre manuel HTML, un schéma `tcstyle.xsd`, un exemple JSON `sample.ppttc` et des présentations d'exercice. Ils fournissent des références attachées au package inspecté, distinctes de la documentation web actuelle.
7. Le schéma de style embarqué déclare l'espace de noms `https://schemas.think-cell.com/36264/tcstyle`. Le numéro de schéma diffère du numéro de build de l'installateur ; il ne faut pas les assimiler.
8. Les trois présentations d'exemple inspectées contiennent des objets OLE avec des flux `think-cellXML`. Le modèle XML référence des tables, catégories, séries et sources de valeurs par identifiants. Les stockages enfants `think-cellChild*` peuvent contenir des packages XLSB ; ils ne doivent pas être traités comme des classeurs XLSX.
9. Dans ce corpus, les objets `CSequenceChartSE` référencent `CSequenceChartDataTable` ; `table/ocol` ordonne les catégories, `m_cscdser` les séries et chaque catégorie référence ses scalaires. Les sources conservent un code de type : une référence absente ou une valeur non numérique n'est pas assimilée à zéro.
10. Les tags de formes et relations OOXML permettent de relier certains objets du modèle aux graphiques natifs de la présentation. Le seul code de type interne ne suffit pas à déterminer le type : un code observé est partagé par dispersion et bulles. Le convertisseur exige donc un choix explicite du type et du mode.

Ces dernières observations proviennent d'un lecteur original CFB/ZIP/XML et d'un corpus local de trois fichiers. Elles n'établissent pas la compatibilité avec tous les fichiers think-cell. Le lecteur et les tests synthétiques sont distribuables ; les fichiers propriétaires du corpus restent hors du dépôt.

## Références REA

| Observation | Evidence ID |
|---|---|
| Ressources de l'installateur | `ev_b4aba210e3a372c8b2a0ae897eb367c7cce0b3dbce6ee5632ca1279ae26639e2` |
| Classification native x64 | `ev_0315266dff8a1705f4e688dbef7bc22579e1774133c3d0b099b778668ca676be` |
| Ressources de la DLL x64 | `ev_6eeaddf37cb91ce3458eee89571fab90994b13a9fcf236bcf7a56d9c9944b44f` |

Ces identifiants référencent les preuves de la session d'analyse. Les bundles bruts, contenant des chemins locaux et des observations de binaires propriétaires, ne font pas partie des sources publiables.

## Inférences et inconnues

Les tables MSI, exports COM et chaînes d'interface étayent une architecture de complément Office natif. Une partie du modèle de persistance est décodée sur le corpus ci-dessus. Le moteur original de placement, les contraintes de cascade, les règles de migration et les comportements d'actualisation restent à établir. L'analyse Ghidra approfondie est en cours ; aucun résultat non obtenu n'est revendiqué.

L'architecture C# de ce dépôt est un choix de reconstruction : cœur géométrique indépendant, adaptateur Office, groupes de formes natives et modèle de données versionné. Elle ne prétend pas être l'architecture interne de think-cell. Les métadonnées `OSC_*` sont propres à ce projet et ne constituent pas une implémentation du format think-cell.

## Références d'intégration

- [Interface de ruban Office](https://learn.microsoft.com/en-us/office/vba/api/office.iribbonextensibility.getcustomui)
- [Tags PowerPoint](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.tags)
- [Inscription des composants COM .NET](https://learn.microsoft.com/en-us/dotnet/framework/interop/registering-assemblies-with-com)

## Règles de contribution

Les implémentations sont originales. Toute affirmation de parité exige une observation de référence et un test comparatif. Les sources, ressources visuelles, installateurs et DLL de think-cell ne sont pas redistribués avec le projet. Le pseudocode d'analyse n'est pas incorporé dans les sources MIT.
