# Utilisation de Git pour la gestion des sources VBA

## Installation

Lancer

    initGit.bat

## Sources

[lien](https://medium.com/@ivanzhd/use-git-hooks-for-version-control-of-excel-vba-code-261fda960fc5)

## Principe

Exploit Git hooks written in Python to dump the Excel VBA modules content into text files each time when you execute git commit command. This method is the most transparent and sustainable as it only uses basic tools and methods of Git, Python and VBA.

## Requirement

- Python
- oletools
`pip install -U oletools`

## ➕📒 Ajout d'un fichier excel si multiple fichiers

1. pre-commit.py:  replace `vba_path = 'src.vba'` with `vba_path = workbook_path + '.vba'`
2. edit your 'hook/pre-commit' file (hook) to have a `git add -- ./<workbookname>.vba`
3. ajouter le repertoire `[excel sheet].vba` à créer dans le fichier `initGit.bat`
4. `git add .` (+commit  si besoin)
5. run `initGit.bat` (cmd)
6. run `initGitFolderForSheets.sh` (gitbash)
pour creer les repertoire necessaire pour stocker les vba

### Exemple

pour ajouter le fichier "test/excel with space.xlsx"

1. pre-commit.py:  replace `vba_path = 'src.vba'` with `vba_path = workbook_path + '.vba'`
2. edit your 'hook/pre-commit' file (hook) to have a `git add -- ./"test/excel with space.xlsx".vba`
3. ajouter le repertoire `"test/excel with space.xlsx".vba` à créer dans le fichier `initGit.bat`

    ```shell
    mkdir -p ./"poub/GS 2021 Tour 1/Tour 1 Homme Séries 3 et 4 NET.xlsx".vba
    touch ./"Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.9.xlsm".vba/1
    git add  ./"Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.9.xlsm".vba/1
    ```

4. `git add .` (+commit  si besoin)
5. run `initGit.bat` (cmd)
6. run `initGitFolderForSheets.sh` (gitbash)
