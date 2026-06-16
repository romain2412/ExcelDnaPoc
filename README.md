# POC ExcelDna — Onglet de ruban personnalisé (sans VBA)

Add-in Excel écrit en **C# / .NET 8** avec [ExcelDna](https://excel-dna.net/) 1.9.
Il ajoute un onglet **« POC ExcelDna »** dans le ruban Excel, avec trois boutons.
Aucune ligne de VBA — toute la logique est en C#.

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
| `RibbonController.cs` | **Ruban** — cœur : assemble le XML CustomUI, capture `IRibbonUI`, câble les délégations |
| `RibbonController.Demonstration.cs` | *partial* — groupe « Demonstration » (commandes simples) |
| `RibbonController.TextTools.cs` | *partial* — groupe « Outils texte » (menu déroulant `<menu>`) |
| `RibbonController.Interactive.cs` | *partial* — groupe « Interactif » (contrôles à état) |
| `RibbonController.TaskPane.cs` | *partial* — adaptateur ruban → `TaskPaneController` |
| `RibbonController.AsyncApi.cs` | *partial* — adaptateur ruban → `JokeApiService` |
| `RibbonController.ContextMenu.cs` | *partial* — **Solution 2** : menu contextuel CustomUI (`getVisible`) |
| `CellRightClickInterceptor.cs` | **Solution 1** : interception `SheetBeforeRightClick` + popup |
| `AddInServices.cs` | Services partagés (volet, service async) + déclencheur commun `ChuckTrigger` |
| `ChuckCommands.cs` | Macro `[ExcelCommand]` appelée par le popup (Solution 1) |
| `TaskPaneController.cs` | **Custom Task Pane** : cycle de vie du volet, hébergement WPF |
| `JokeApiService.cs` | **Async + `QueueAsMacro` + annulation** (`CancellationToken`) |
| `WpfPaneHost.cs` | Pont **WinForms ↔ WPF** (`ElementHost`, ComVisible) |
| `WpfPane.xaml` / `WpfPane.xaml.cs` | **Contenu WPF** du volet (UI + interactions Excel) |

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
  contenu est du **WPF**. Boutons : écrire/lire la cellule active, compter la sélection ; avec journal.

**Groupe « Reseau »**
- **Blague (async)** (`button`) — appelle l'API `https://api.chucknorris.io/jokes/random` **sans
  bloquer Excel**, lance un **`Task.Delay` asynchrone de 45 s** avec **barre de progression**
  (compte à rebours) dans le volet WPF, puis écrit la blague dans la **cellule active de la feuille
  affichée au moment du clic** — **après** l'attente. Pendant ces 45 s, Excel reste **pleinement
  utilisable** (saisie, navigation, etc.).
- **Annuler** (`button`, `getEnabled`) — actif uniquement pendant une opération ; interrompt
  proprement l'attente en cours via un `CancellationToken` (la blague n'est alors pas écrite).

**Menu contextuel (clic droit sur une cellule)** — 3 comportements selon le contenu de la cellule :

| Contenu de la cellule | Au clic droit | Implémentation |
|-----------------------|---------------|----------------|
| `BlagueMenuExcel` | entrée **« Blague (async) »** **ajoutée** au menu Excel normal | CustomUI `getVisible` sur un `<button>` du `<contextMenu idMso='ContextMenuCell'>` ([RibbonController.ContextMenu.cs](RibbonController.ContextMenu.cs)) |
| `BlagueMenuWinform` | menu **WinForms** ne contenant **que** « Blague (async) » | interception + `ContextMenuStrip.Show()` ([CellRightClickInterceptor.cs](CellRightClickInterceptor.cs)) |
| `BlagueMenuWpf` | menu **WPF** ne contenant **que** « Blague (async) » | interception + `System.Windows.Controls.ContextMenu` (`IsOpen=true`, `Placement=MousePoint`) ([CellRightClickInterceptor.cs](CellRightClickInterceptor.cs)) |
| autre | menu Excel habituel | — |

Les trois déclenchent le **même** comportement async que le bouton du ruban (via `ChuckTrigger.Run`).
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

### Pattern asynchrone (important)
Le modèle objet Excel est **mono-thread (STA)** : interdit d'y toucher depuis le thread de
continuation après un `await`. Le pattern correct :
1. **Au clic** (thread principal Excel) : capturer la cible (`ActiveCell`) et lancer l'appel en
   *fire-and-forget* (`_ = FetchAndWriteJokeAsync(target)`), sans bloquer l'UI.
2. **Après l'`await`** (thread de fond) : récupérer/parser la réponse — **ne pas** toucher Excel ici.
3. **Réécriture** : `ExcelAsyncUtil.QueueAsMacro(() => target.Value2 = texte)` re-marshale l'écriture
   sur le thread principal d'Excel. Voir `OnChuckNorrisClick` / `FetchAndWriteJokeAsync` dans
   [RibbonController.cs](RibbonController.cs).

### Architecture du volet WPF
Un Custom Task Pane Office n'héberge qu'un contrôle WinForms COM-visible. On empile donc :
`CustomTaskPane → WpfPaneHost (WinForms, ComVisible) → ElementHost → WpfPane (WPF/XAML)`.
Fichiers : [WpfPane.xaml](WpfPane.xaml) + [WpfPane.xaml.cs](WpfPane.xaml.cs) (contenu),
[WpfPaneHost.cs](WpfPaneHost.cs) (pont WinForms↔WPF). Activé par `<UseWPF>true</UseWPF>` dans le `.csproj`.

> Astuce : le `editBox`/`comboBox` valident la saisie avec **Entrée**. Les `checkBox`/`toggleButton`
> lisent l'état de la cellule **active au chargement du ruban** ; sélectionne une cellule puis
> rouvre l'onglet (ou appelle `Invalidate`) pour resynchroniser l'état affiché.

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

## Développement : recharge rapide

Après modification du code, relancez `dotnet build` puis rechargez l'add-in dans Excel
(décocher / recocher dans la liste des compléments, ou redémarrer Excel).
