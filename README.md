# POC ExcelDna — Onglet de ruban personnalisé (sans VBA)

Add-in Excel écrit en **C# / .NET 8** avec [ExcelDna](https://excel-dna.net/) 1.9.
Il ajoute un onglet **« POC ExcelDna »** dans le ruban Excel, avec trois boutons.
Toute la logique est en C# — à **une exception assumée près** : une petite macro VBA
([ChuckMacro.bas](ChuckMacro.bas)) qui sert justement à démontrer comment **consommer un
objet COM C# depuis VBA** (cf. section dédiée plus bas).

## Prérequis

- .NET SDK 8 ou 9
- Microsoft Excel **64 bits** (détecté sur cette machine)

## Organisation du code

Le code est découpé pour qu'**un fichier = un concept/une fonctionnalité ExcelDna**.

> Contrainte ExcelDna : **tous les callbacks du ruban doivent appartenir à l'unique classe
> dérivée de `ExcelRibbon`** (ExcelDna les résout par nom sur cet objet). On utilise donc une
> **classe partielle** `RibbonController` répartie par fonctionnalité (un fragment XML + ses
> callbacks par fichier). Les concepts autonomes sont de **vraies classes séparées** vers
> lesquelles le ruban délègue.

| Fichier | Concept / fonctionnalité ExcelDna |
|---------|-----------------------------------|
| `ExcelDnaPoc.csproj` | Projet .NET 8, référence `ExcelDna.AddIn`, génère le `.xll` |
| `AddIn.cs` | **Cycle de vie** de l'add-in (`IExcelAddIn` : `AutoOpen`/`AutoClose`) |
| `StartupConfig.cs` / `StartupLoader.cs` | **Démarrage configurable** : ouvre des classeurs au chargement selon `startup.json` |
| `WorkbookEventsBinder.cs` / `WorksheetEventsBinder.cs` | **Événements Excel → callbacks C#** (niveaux Workbook et Worksheet) sur les classeurs ouverts |
| `TestAddin.xlsx` | Classeur de test (`BlagueMenuExcel` / `BlagueMenuWinform` / `BlagueMenuWpf`) |
| `RibbonController.cs` | **Ruban** — cœur : assemble le XML CustomUI, capture `IRibbonUI`, câble les délégations |
| `RibbonController.Demonstration.cs` | *partial* — groupe « Demonstration » (commandes simples) |
| `RibbonController.TextTools.cs` | *partial* — groupe « Outils texte » (menu déroulant `<menu>`) |
| `RibbonController.Interactive.cs` | *partial* — groupe « Interactif » (contrôles à état) |
| `RibbonController.TaskPane.cs` | *partial* — adaptateur ruban → `TaskPaneController` |
| `RibbonController.AsyncApi.cs` | *partial* — adaptateur ruban → `ChuckTrigger` |
| `RibbonController.Udf.cs` | *partial* — menu **« UDF Sync »** : insère une formule UDF dans la sélection |
| `RibbonController.ContextMenu.cs` | *partial* — **Solution 2** : menu contextuel CustomUI (`getVisible`) |
| `CellRightClickInterceptor.cs` | **Solution 1** : interception `SheetBeforeRightClick` + popup |
| `AddInServices.cs` | Services partagés (volet, service async) + déclencheur commun `ChuckTrigger` |
| `ChuckCommands.cs` | Macro `[ExcelCommand]` appelée par le popup (Solution 1) |
| `Functions.cs` | **UDF** (`[ExcelFunction]`) appelables depuis une cellule — `=POC.ADDITION(...)` |
| `TaskPaneController.cs` | **Custom Task Pane** : cycle de vie du volet, hébergement WPF |
| `JokeJob.cs` | **Un traitement async indépendant** : `CancellationTokenSource` propre, **parallélisable** |
| `JobRow.xaml` (+`.cs`) | Ligne d'UI d'un traitement : statut + **barre de progression** + **Annuler** |
| `WpfPaneHost.cs` | Pont **WinForms ↔ WPF** (`ElementHost`, ComVisible) |
| `WpfPane.xaml` / `WpfPane.xaml.cs` | **Contenu WPF** du volet (UI + interactions Excel) |
| `ChuckCom.cs` | **Objet COM** (`[ComVisible]` + `ProgId`) consommé depuis VBA — async **fire-and-forget** |
| `ChuckMacro.bas` | **Source** de la macro VBA (texte, diffable, versionnée) |
| `Generate-Xlam.ps1` | Régénère `ChuckMacro.xlam` depuis le `.bas` (Excel caché + injection VBA) |
| `ChuckMacro.xlam` | **Complément VBA** *généré* (non versionné) — chargé au démarrage (`startup.json`) |
| `RibbonController.Vba.cs` | *partial* — bouton « Blague COM (VBA) » → `Application.Run` la macro |

## Contrôles de l'onglet

**Groupe « Demonstration »**
1. **Dire Bonjour** — affiche une boîte de dialogue.
2. **Écrire dans la cellule** — écrit une valeur dans la cellule active via le modèle objet Excel.
3. **Somme sélection** — additionne les cellules numériques sélectionnées.

**Groupe « Outils texte »**
- **Transformer** — bouton avec menu déroulant (`<menu>`) : MAJUSCULES, minuscules, insérer date/heure, effacer.

**Groupe « Interactif »** (contrôles à état, rafraîchis via `IRibbonUI.Invalidate`)
- **Nom** (`editBox`) — mémorise un texte ; **Bonjour {nom}** (`button` dynamique) — libellé recalculé via `getLabel`.
- **Gras** (`checkBox`) / **Italique** (`toggleButton`) — reflètent (`getPressed`) et appliquent le style de la cellule active.
- **Fond** (`dropDown`) — couleur d'arrière-plan ; **Police** (`comboBox` éditable) — police de la cellule.

**Groupe « Volet »**
- **Volet WPF** (`toggleButton`) — affiche/masque un **Custom Task Pane** ancré à droite, dont le
  contenu est du **WPF** : une **liste de traitements** (un par déclenchement), des boutons
  écrire/lire la cellule active, compter la sélection, et un journal.

**Groupe « Reseau »**
- **Blague (async)** (`button`) — appelle l'API `https://api.chucknorris.io/jokes/random` **sans
  bloquer Excel**, lance un **`Task.Delay` asynchrone de 15 s** (barre de progression dans le volet),
  puis écrit la blague dans la **cellule à droite de la cellule sélectionnée** — **après** l'attente.

**Traitements parallèles + annulation individuelle**
- Chaque déclenchement (bouton ruban, menus contextuels, double-clic) crée un **traitement
  indépendant** (`JokeJob`, son propre `CancellationTokenSource`) avec **sa propre ligne** dans le
  volet : statut + **barre de progression** + bouton **Annuler**. Plusieurs traitements tournent donc
  **en parallèle** ; lancer un nouveau traitement **n'interrompt plus** les précédents.
- Le bouton **Annuler** d'une ligne n'annule **que ce traitement-là** (`CancellationToken`) ; la ligne
  disparaît quelques secondes après la fin (ou l'annulation).

**Menu contextuel (clic droit sur une cellule)** — 3 comportements selon le contenu de la cellule :

| Contenu de la cellule | Au clic droit | Implémentation |
|-----------------------|---------------|----------------|
| `BlagueMenuExcel` | entrée **« Blague (async) »** **ajoutée** au menu Excel normal | CustomUI `getVisible` sur un `<button>` du `<contextMenu idMso='ContextMenuCell'>` ([RibbonController.ContextMenu.cs](RibbonController.ContextMenu.cs)) |
| `BlagueMenuWinform` | menu **WinForms** ne contenant **que** « Blague (async) » | interception + `ContextMenuStrip.Show()` ([CellRightClickInterceptor.cs](CellRightClickInterceptor.cs)) |
| `BlagueMenuWpf` | menu **WPF** ne contenant **que** « Blague (async) » | interception + `System.Windows.Controls.ContextMenu` (`IsOpen=true`, `Placement=MousePoint`) ([CellRightClickInterceptor.cs](CellRightClickInterceptor.cs)) |
| autre | menu Excel habituel | — |

Les trois déclenchent le **même** comportement async que le bouton du ruban (via `ChuckTrigger.Fire`).
Chaînes déclencheuses = constantes (`CtxTrigger`, `TriggerWinforms`, `TriggerWpf`) faciles à modifier.

### Notes d'implémentation (pièges rencontrés)
- **CustomUI** : `<contextMenus>` doit être **après** `<ribbon>` (ordre du schéma : `commands`, `ribbon`,
  `backstage`, `contextMenus`). `getVisible` est OK sur un `<button>` de menu contextuel mais **invalide
  sur un `<menuSeparator>`** → Excel rejette alors tout le CustomUI **en silence** (ni ruban ni erreur).
- **Interception** (Winform/Wpf) : abonnement à `Application.SheetBeforeRightClick` **sans aucun interop
  Office** (sous .NET 8, `<COMReference>` n'est pas supporté par `dotnet build`, et le package interop
  échoue au runtime : `FileNotFoundException 'office'`). On passe par les **points de connexion COM**
  (`IConnectionPoint*` du BCL) + un sink IDispatch portant le `DispId(1560)` de `SheetBeforeRightClick`.
- **« Uniquement notre entrée »** : impossible via `CommandBars` (Excel moderne **injecte** des items —
  barre de recherche, options de collage… — hors `CommandBars`), et `CommandBar.ShowPopup()` renvoie
  `E_FAIL` depuis un add-in .NET. La solution qui marche : `Cancel=true` (supprime **tout** le menu Excel,
  même les items injectés) puis afficher **notre propre menu** WinForms/WPF, en différé via `QueueAsMacro`.

## UDF — fonctions de feuille (`=POC.…`)

Le cœur d'ExcelDna : des **fonctions personnalisées** appelables depuis une cellule
([Functions.cs](Functions.cs)). Méthodes marquées `[ExcelFunction]` (avec `Name`, `Description`,
`Category` et `[ExcelArgument]` qui alimentent l'**Assistant de fonction** `fx`) :

| Fonction | Exemple | Résultat |
|----------|---------|----------|
| `POC.ADDITION(a; b)` | `=POC.ADDITION(2;3)` | `5` |
| `POC.BONJOUR(nom)` | `=POC.BONJOUR("Romain")` | `Bonjour Romain !` |
| `POC.SOMMEPERSO(plage)` | `=POC.SOMMEPERSO(H1:H4)` | `35` *(somme des nombres)* |
| `POC.CONCAT(plage; sep)` | `=POC.CONCAT(H1:H4;" \| ")` | `10 \| 20 \| texte \| 5` |
| `POC.DOUBLER(plage)` | `=POC.DOUBLER(H1:H4)` | tableau `20;40;texte;10` *(spill / matricielle)* |
| `POC.MOYENNE(plage)` | `=POC.MOYENNE(L1:L4)` | `2,5` *(`double[,]` — texte → `#VALEUR!`)* |
| `POC.COMPTENB(plage)` | `=POC.COMPTENB(H1:H4)` | `3` *(`object[]` 1D : ligne/colonne)* |
| `POC.INFOPLAGE(plage)` | `=POC.INFOPLAGE(H1:H4)` | `3 nombre(s), 1 texte(s), 0 vide(s)…` |
| `POC.PRODUITSCALAIRE(p1; p2)` | `=POC.PRODUITSCALAIRE(L1:L4;L1:L4)` | `30` *(2 plages ; tailles ≠ → `#VALEUR!`)* |

**Types de paramètres « plage » et conversions** :
- **`object[,]`** (2D, n'importe quelle plage) : reçoit **tout** — nombre → `double`, texte → `string`,
  vide → `ExcelEmpty`, erreur → `ExcelError`, optionnel omis → `ExcelMissing`. On trie au code.
- **`double[,]`** : ExcelDna **convertit en nombres** — vide → `0`, nombre → sa valeur, **texte →
  échec → la fonction renvoie `#VALEUR!`** (la plage doit être numérique/vide).
- **`object[]`** : plage **1D** (une seule ligne ou colonne).
- **Plusieurs plages** : autant de paramètres « plage » que voulu (cf. `POC.PRODUITSCALAIRE`).
- **Renvoyer un `object[,]`** → un tableau qui « spill » (dynamic arrays) ou formule matricielle
  (`Ctrl+Maj+Entrée`). **Renvoyer une erreur** : `return ExcelError.ExcelErrorValue;` → `#VALEUR!`.

> `ExcelAddInExplicitExports=true` → seules les méthodes marquées `[ExcelFunction]`/`[ExcelCommand]`
> sont exportées. (Pistes ExcelDna non encore explorées : UDF **asynchrones**, **RTD** temps réel,
> IntelliSense, API C bas niveau…)

**Menu ruban « UDF Sync »** ([RibbonController.Udf.cs](RibbonController.Udf.cs)) : un bouton à menu
déroulant ; chaque entrée **insère sa formule dans la/les cellule(s) sélectionnée(s)**. La formule est
portée par l'attribut `tag` du `<button>` → **un seul callback** `OnInsertFormula` la lit et l'écrit
(`Selection.Formula`) → pour ajouter une formule au menu, il suffit d'ajouter un `<button tag='=…'/>`.
Les exemples sont **auto-suffisants** : les UDF à plage utilisent des **constantes matricielles**
(`{10;20;30}`) → ils donnent un résultat **depuis n'importe quelle cellule** (pas de dépendance aux
cellules voisines). *(`.Formula` = syntaxe US : `,` sépare les arguments, `;` les lignes d'une
constante matricielle.)*

**Exemples prêts à l'emploi** : le classeur de test [TestAddin.xlsx](TestAddin.xlsx) contient une
feuille **« UDF »** avec des données d'exemple et **toutes les formules** en action :
- **plages 1D** (une colonne : `A2:A5`, `B2:B5`) ;
- **plages 2D** (`A2:B5` = 4 lignes × 2 colonnes) : `SOMMEPERSO`→45, `INFOPLAGE`→« 7 nombres, 1 texte »,
  `CONCAT` (ordre **ligne par ligne** : `10 | 1 | 20 | 2 | …`), et `DOUBLER(A2:B5)` qui renvoie un
  **tableau 2D 4×2** (en G2:H5). → un `object[,]` reçoit bien les **deux dimensions**.

## Démarrage configurable (ouverture de classeurs)

L'add-in possède un point de **démarrage** (`AddIn.AutoOpen` → `StartupLoader.Run`) piloté par
configuration ([StartupConfig.cs](StartupConfig.cs) / [StartupLoader.cs](StartupLoader.cs)) :

- Au chargement, l'add-in lit **`startup.json`** placé **à côté du `.xll`**. Schéma actuel
  (appelé à évoluer — critères, conditions…), avec chemins **relatifs** au dossier du `.xll` :
  ```json
  { "openWorkbooks": [ "TestAddin.xlsx" ] }
  ```
- Chaque classeur listé est ouvert **de façon asynchrone** : l'ouverture est différée via
  `ExcelAsyncUtil.QueueAsMacro`, donc elle ne **bloque pas** `AutoOpen` (le journal montre
  « AutoOpen fin » *avant* « classeur ouvert »).
- **Fonctionne dans les deux modes de lancement** grâce aux chemins relatifs :
  - **F5 (Visual Studio)** : `startup.json` et `TestAddin.xlsx` sont copiés dans le dossier de
    sortie via `CopyToOutputDirectory` (csproj), à côté du `.xll` non packé.
  - **`Launch-AddIn.ps1`** : le script copie ces deux fichiers à côté du `.xll` packé (`publish\`).
- Classeur de test : [TestAddin.xlsx](TestAddin.xlsx) (contient `BlagueMenuExcel`, `BlagueMenuWinform`,
  `BlagueMenuWpf` pour tester les 3 menus contextuels).

### Événements Excel → callbacks C#

Plusieurs événements Excel sont **bindés sur des callbacks C#** (qui écrivent dans le journal
`%TEMP%\ExcelDnaPoc.log`). Tous utilisent la **même technique** : points de connexion COM
(`IConnectionPoint*` du BCL) + un **sink IDispatch** portant le GUID de l'interface source et les
bons `[DispId]` — **sans aucun interop Office**.

**Niveau Application** — global, abonné une fois au chargement de l'add-in
([CellRightClickInterceptor.cs](CellRightClickInterceptor.cs), interface `AppEvents` `00024413-…`) :

| Événement | DISPID | Déclenché par |
|-----------|:------:|---------------|
| `SheetBeforeRightClick` | 1560 | **clic droit** sur une cellule (n'importe quel classeur) — sert à l'interception du menu (Solutions 1/2) |

**Niveau Workbook** — abonné **par classeur ouvert** au démarrage
([WorkbookEventsBinder.cs](WorkbookEventsBinder.cs), interface `WorkbookEvents` `00024412-…`) :

| Événement | DISPID | Déclenché par |
|-----------|:------:|---------------|
| `SheetSelectionChange` | 1558 | la **sélection change** dans une feuille du classeur |
| `SheetChange` | 1564 | le **contenu d'une cellule change** (saisie utilisateur **ou** écriture par code) dans le classeur |
| `Activate` | 304 | le **classeur passe au premier plan** |
| `Deactivate` | 1530 | un **autre classeur devient actif** |
| `BeforeClose` | 1546 | **juste avant la fermeture** du classeur — **annulable** (`ref bool Cancel`) |

**Niveau Worksheet** — abonné **par feuille** des classeurs ouverts
([WorksheetEventsBinder.cs](WorksheetEventsBinder.cs), interface `DocEvents` `00024411-…`) :

| Événement | DISPID | Déclenché par |
|-----------|:------:|---------------|
| `SelectionChange` | 1543 | la **sélection change** dans **cette** feuille précise |
| `Change` | 1545 | une **cellule de cette feuille** est modifiée |
| `Activate` | 304 | on **bascule sur cet onglet** |
| `Deactivate` | 1530 | on **quitte cet onglet** |
| `BeforeDoubleClick` | 1537 | **double-clic** sur une cellule → déclenche la **Blague async** (même `ChuckTrigger.Fire` que les menus / le bouton ruban) ; `Cancel=true` empêche l'entrée en édition |
| `Calculate` | 279 | la feuille est **recalculée** (formules réévaluées) |

> **Workbook vs Worksheet** : les événements `Sheet*` (niveau Workbook) se déclenchent pour **toutes**
> les feuilles du classeur (le paramètre `Sh` indique laquelle) ; les événements Worksheet ne se
> déclenchent que pour **leur** feuille. Donc sélectionner une cellule loggue **deux** lignes :
> `(ws) SelectionChange` *et* `SheetSelectionChange`.
>
> **Pour ajouter un événement** : déclarer un `[DispId(...)]` de plus dans l'interface concernée
> (`IWorkbookEvents` / `IWorksheetEvents`) — les DISPID se récupèrent par réflexion de l'interop Excel.

### Pattern asynchrone (important)
Le modèle objet Excel est **mono-thread (STA)** : interdit d'y toucher depuis le thread de
continuation après un `await`. Deux principes :

**1. Async « jusqu'en haut » + fire-and-forget à la frontière.**
La logique est `async Task` de bout en bout (`JokeJob.RunAsync` ← `ChuckTrigger.RunAsync`).
Le point où l'on « redescend » en synchrone dépend du **type** de point d'entrée :

- **Callbacks Excel/COM** (`onAction` du ruban, sinks d'événements, macros, double-clic) : Excel les
  appelle **synchroniquement** et **n'attend pas** le `Task` → on ne peut **pas** propager l'`async`
  *dans* Excel. Fire-and-forget via **`ChuckTrigger.Fire()`** (`void`, lance le `Task` et **observe les
  exceptions**).
- **Handlers d'événements UI WinForms/WPF** (`Click` des menus contextuels custom) : eux **peuvent**
  être **`async void`** (le seul usage légitime d'`async void`) → l'`async` remonte **jusqu'au
  handler**, qui `await ChuckTrigger.RunSafeAsync()` directement. Pas de fire-and-forget « caché ».

**2. Réécrire dans Excel uniquement sur le thread principal.**
Après un `await` (thread de fond), **ne pas** toucher Excel directement : repasser par
`ExcelAsyncUtil.QueueAsMacro(...)` (voir `JokeJob.Finish`). C'est aussi ce qui rend Excel
réactif pendant l'attente de 15 s.

### Architecture du volet WPF
Un Custom Task Pane Office n'héberge qu'un contrôle WinForms COM-visible. On empile donc :
`CustomTaskPane → WpfPaneHost (WinForms, ComVisible) → ElementHost → WpfPane (WPF/XAML)`.
Fichiers : [WpfPane.xaml](WpfPane.xaml) + [WpfPane.xaml.cs](WpfPane.xaml.cs) (contenu),
[WpfPaneHost.cs](WpfPaneHost.cs) (pont WinForms↔WPF). Activé par `<UseWPF>true</UseWPF>` dans le `.csproj`.

> Astuce : le `editBox`/`comboBox` valident la saisie avec **Entrée**. Les `checkBox`/`toggleButton`
> lisent l'état de la cellule **active au chargement du ruban** ; sélectionne une cellule puis
> rouvre l'onglet (ou appelle `Invalidate`) pour resynchroniser l'état affiché.

## Objet COM (C#) consommé depuis VBA — fire-and-forget

Démontre l'aller-retour complet **ruban (C#) → macro VBA → objet COM (C#)**, où le travail
asynchrone **rend la main à Excel** sans aucun `DoEvents` ni événement.

- [ChuckCom.cs](ChuckCom.cs) : classe `[ComVisible(true)]` + `[ProgId("ExcelDnaPoc.Chuck")]`
  exposant l'interface IDispatch `IChuckCom`. ExcelDna en fait un **serveur COM** :
  `<ExcelAddInComServer>true</ExcelAddInComServer>` (csproj → `ComServer="true"` dans le `.dna`)
  + `ExcelDna.ComInterop.ComServer.DllRegisterServer()` dans `AutoOpen` → la classe devient
  instanciable par `CreateObject("ExcelDnaPoc.Chuck")` (et `DllUnregisterServer()` dans `AutoClose`).
- **Fire-and-forget** : `LancerBlague()` démarre l'appel API async **puis** un `Task.Delay(15 s)`
  async, et **rend la main immédiatement** → la macro VBA se termine, **Excel reste réactif**
  pendant tout l'IO. L'objet écrit la blague (cellule à droite de l'active) quand il a fini,
  via `QueueAsMacro` (thread principal). Aucun thread n'est tenu pendant l'attente.
- [ChuckMacro.bas](ChuckMacro.bas) : **source** (versionnée) de la macro `LancerBlagueAsync`
  = `CreateObject(...)` + appel.
- `ChuckMacro.xlam` : **complément VBA généré** depuis le `.bas` (artefact, non versionné — `.gitignore`).
  Il est **ouvert automatiquement au démarrage** (listé dans [startup.json](startup.json)) → ses
  macros sont disponibles globalement pour `Application.Run`, **sans aucune manipulation**.
- Bouton ruban **« Blague COM (VBA) »** ([RibbonController.Vba.cs](RibbonController.Vba.cs)) →
  `Application.Run("LancerBlagueAsync")`.

**Utilisation** : sélectionner une cellule, cliquer **« Blague COM (VBA) »** dans le groupe
**« VBA + COM »**. Excel reste utilisable ; la blague apparaît **~16 s** plus tard dans la cellule
à droite (l'IO async ne tient aucun thread).

### Génération du `.xlam` (automatique)

Le `.xlam` est **régénéré depuis `ChuckMacro.bas` à chaque build et à chaque lancement** par
[Generate-Xlam.ps1](Generate-Xlam.ps1) (Excel caché → `VBProject.VBComponents.Import` → `SaveAs`
format 55 `.xlam`, via `%TEMP%` puis `Copy-Item` car le SaveAs direct vers OneDrive échoue) :

- **Build VS / VS Code / `dotnet build`** : cible MSBuild `GenerateChuckXlam` (`AfterTargets="Build"`),
  puis copie à côté du `.xll` (sortie non packée, pour F5).
- **`Launch-AddIn.ps1`** : appelle le même script, puis copie vers `publish\` (à côté du `.xll` packé).

> **Prérequis** : Excel installé **et** « Accès approuvé au modèle objet du projet VBA »
> (Fichier → Options → Centre de gestion de la confidentialité → Paramètres des macros). Si absent
> (ex. build CI sans Excel), la génération est **ignorée avec un avertissement** — le build ne casse pas
> (`ContinueOnError`), mais le bouton VBA ne fonctionnera pas tant que le `.xlam` n'est pas généré.

> **Test express** (sans le `.xlam`) : fenêtre **Exécution** (`Ctrl`+`G`) du VBE →
> `CreateObject("ExcelDnaPoc.Chuck").LancerBlague` puis `Entrée`.

## Compiler

```powershell
dotnet build -c Debug
```

Le `.xll` est généré dans `bin\x64\Debug\net8.0-windows\` :
`ExcelDnaPoc-AddIn64.xll` (version 64 bits, à charger dans Excel 64 bits).

Une version « packée » autonome (tout dans un seul `.xll`, plus simple à distribuer)
est aussi produite dans `bin\x64\Debug\net8.0-windows\publish\ExcelDnaPoc-AddIn64-packed.xll`.

## Charger dans Excel

**Recommandé — script de lancement** (charge l'add-in via COM `RegisterXLL`, sans prompt de sécurité) :

```powershell
.\Launch-AddIn.ps1
```

> ⚠️ Passer le `.xll` en argument à `excel.exe` (ou double-clic) ne le charge que pour
> la session **et** déclenche une boîte « problème de sécurité potentiel » qu'il faut
> valider (*Activer ce complément*). Si elle n'est pas validée, l'onglet n'apparaît pas.
> C'est pourquoi le script utilise plutôt l'API COM `RegisterXLL`.

**Manuellement (chargement persistant)** : Excel → *Fichier* → *Options* → *Compléments* →
*Gérer : Compléments Excel* → *Atteindre…* → *Parcourir…* → sélectionner le `.xll`.
Excel demandera d'activer le complément (cocher la case).

> Si Excel bloque le fichier (zone non sûre / OneDrive), clic droit sur le `.xll`
> → *Propriétés* → cocher *Débloquer*.

## Déboguer

**Visual Studio** : ouvrir `ExcelDnaPoc.sln`, vérifier que le profil **« Excel »** est sélectionné,
puis **F5** (le profil `launchSettings.json` est généré par `RunExcelDnaSetDebuggerOptions`).

**Visual Studio Code** : config fournie dans `.vscode/` (nécessite l'extension **C# Dev Kit**) —
- **« Excel : lancer + debug »** (F5) : build, lance Excel avec le `.xll` non packé et attache le
  débogueur ; les points d'arrêt dans le C# sont actifs.
- **« Excel : attacher »** : repli — attacher à une instance Excel déjà lancée (ex. via
  `Launch-AddIn.ps1`), en choisissant le process `EXCEL.EXE`.

Dans les deux IDE :
- au **1er lancement**, Excel affiche le prompt « Activer le complément » (chargement par `/x`) → cliquer
  *Activer* (ou ajouter le dossier de sortie aux **Emplacements approuvés** d'Excel) ;
- **fermer Excel avant de recompiler** (le `.xll` chargé est verrouillé) — à l'arrêt du débogage,
  l'IDE ferme Excel automatiquement.

## Développement : recharge rapide

Après modification du code, relancez `dotnet build` puis rechargez l'add-in dans Excel
(décocher / recocher dans la liste des compléments, ou redémarrer Excel) — ou utilisez
`Launch-AddIn.ps1` (chargement par COM `RegisterXLL`, sans prompt de sécurité).
