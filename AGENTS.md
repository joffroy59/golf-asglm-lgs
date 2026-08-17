# Repository Guidelines

## Project Structure & Module Organization
This repository centers on the Excel workbook `Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm`. Source-controlled VBA modules are exported into `vba/Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm.vba/` as `.bas` and `.cls` files. Helper documentation lives in [`readme.md`](./readme.md), release notes in [`changelog.md`](./changelog.md), and Git/VBA setup notes in [`git.md`](./git.md). Batch utilities such as `init.bat`, `initYear.bat`, `initGit.bat`, and `backup.bat` manage yearly setup, hook installation, and export backups. Sample FFG exports are stored under `fichier exemple nom export FFG/`.

## Build, Test, and Development Commands
Use Windows batch scripts from the repository root:

- `initGit.bat`: installs the Git hooks from `.hooks/` into `.git/hooks/`.
- `initYear.bat`: creates a new `..\ASGLM <YEAR>\LGS` workspace and seeds it with tour folders and workbook copies.
- `init.bat`: initializes the current `LGS` folder with `T1` to `T7`, `Finale`, `Backup`, and `Poub`.
- `backup.bat`: archives `2d. Extraction XLS globale.xls` into a dated `backup_YYYY-MM-DD` folder.

After editing the workbook in Excel, make a normal `git commit`; the pre-commit hook runs `olevba` export automatically and stages `vba/` plus `git.log`.

## Year Initialization Procedure
To prepare a new season, first create the parent folder `..\ASGLM <YEAR>` if it does not already exist. From the repository root, run `initYear.bat`, enter the target year when prompted, and let the script create `..\ASGLM <YEAR>\LGS`. It copies `init.bat` and `backup.bat`, then generates the working tree with `T1` to `T7`, `Finale`, `Backup`, and `Poub`, plus one workbook copy per tour. After setup, confirm that `readme.md`, `fichier exemple nom export FFG/`, and files such as `Calcul La Grande Semaine - STROKEPLAY - T1 - HOMME_OU_DAME_v2.12.xlsm` are present in the new `LGS` folder before starting imports.

## Coding Style & Naming Conventions
Keep VBA modules focused by responsibility, following the existing pattern: `import.bas`, `process.bas`, `lib*.bas`, and `ThisWorkbook.cls`. Use 4-space indentation in VBA, preserve French business labels already used in the workbook UI, and prefer descriptive PascalCase or camelCase procedure names over abbreviations. Do not rename exported module files unless the workbook module name changes too.

## Testing Guidelines
There is no automated test suite in this repository. Validate changes by opening the workbook in Excel, importing the sample file `fichier exemple nom export FFG/2d. Extraction XLS globale.xls`, and checking the Homme and Dame result tabs. For import or cleanup changes, verify both single-tour import and multi-tour import flows described in `readme.md`.

## Commit & Pull Request Guidelines
Follow the lightweight history style already used here: short, version-oriented messages such as `fix: update RMS import parsing` or `docs: clarify initYear setup`. Keep commits scoped to one workbook or workflow change. Pull requests should include a brief summary, affected tour/import scenarios, manual validation steps, and screenshots only when worksheet layout or visible output changes.

## Security & Configuration Tips
Git hook export requires Python and `oletools` (`pip install -U oletools`). Avoid committing live player data or year-specific working folders outside the repository template. Keep sample exports anonymized when adding new fixtures.
