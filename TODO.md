# TODO

## Rendre les fichiers `.xlsx` diffables dans Git (à explorer)

Aujourd'hui `TestAddin.xlsx` est versionné **en binaire** (choix assumé) : `git diff` affiche
seulement « Binary files differ », car un `.xlsx` est une archive ZIP de fichiers XML.

Piste pour rendre le diff lisible (sans changer le format du fichier) : configurer un
**`textconv`** Git qui décompresse le `.xlsx` et expose son XML/texte interne au moment du diff.

Esquisse :

1. `.gitattributes` à la racine :
   ```
   *.xlsx diff=xlsx
   ```
2. Driver de diff (config locale, non versionnée — à scripter dans un setup) :
   ```
   git config diff.xlsx.textconv "<commande qui dezippe et dump le XML/texte>"
   git config diff.xlsx.binary true
   ```
   Exemples de `textconv` possibles :
   - `unzip -p` sur les parties pertinentes (`xl/worksheets/*.xml`, `xl/sharedStrings.xml`) ;
   - un petit script Python (openpyxl) qui sort le contenu des cellules en texte/CSV ;
   - un outil dédié (ex. `xlsx2csv`).

Limites à garder en tête :
- le diff restera verbeux (XML), et le bruit de re-sauvegarde (métadonnées, `calcChain`,
  ordre de compression) subsistera ;
- le `textconv` doit être installé/configuré sur chaque poste (prévoir un script de setup).

Alternative déjà identifiée (non retenue pour l'instant) : ne pas versionner le `.xlsx` et le
**générer** à la volée (comme `startup.json`).
