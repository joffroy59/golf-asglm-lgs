# Repository Guidelines

## Project Structure
The Excel workbook `Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm` is the score-calculation source of truth. Exported VBA source lives in `vba/Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm.vba/` as `.bas` and `.cls` files. Batch utilities at the root initialize seasons and backups. Sample RMS/FFG exports live in `fichier exemple nom export FFG/`.

`app/` contains the dependency-free browser dashboard for tracking seasons. Keep its metadata separate from Excel calculations and macros. For the complete application architecture and data flow, see [app/architecture.mmd](app/architecture.mmd).

## Build, Test, and Development
Run these Windows batch files from the repository root:

- `initGit.bat` installs the VBA-export Git hooks.
- `initYear.bat` creates a new yearly LGS workspace.
- `init.bat` seeds the current folder with `T1` through `T7`, `Finale`, `Backup`, and `Poub`.
- `backup.bat` archives `2d. Extraction XLS globale.xls` into a dated folder.

Open `app/index.html` in a modern browser to run the season dashboard; no server or package installation is required. The pre-commit hook needs Python and `oletools` (`pip install -U oletools`) to export VBA before a normal `git commit`.

## Year Initialization
Create `..\ASGLM <YEAR>` if necessary, then run `initYear.bat` and enter the year. It creates `..\ASGLM <YEAR>\LGS` with tour folders and workbook copies. Verify that `readme.md`, the sample export folder, and a workbook such as `... - T1 - ...xlsm` are present before importing results.

Create the same year in `app/index.html` and use the matching LGS folder as its reference. Export its JSON archive into that folder before moving to another browser or computer.

## Style and Testing
Use 4-space indentation in VBA. Keep existing French business labels, use descriptive procedure names, and do not rename exported module files unless the workbook module name changes. For the browser app, use semantic HTML, plain CSS, and vanilla JavaScript; preserve its local-first behavior.

There is no automated suite. Test VBA changes by importing `fichier exemple nom export FFG/2d. Extraction XLS globale.xls` and checking Homme and Dame results. Test dashboard changes by creating a season, updating a tour, then exporting and importing its JSON.

## Commits and Pull Requests
Use short scoped messages, for example `fix: update RMS import parsing` or `docs: clarify initYear setup`. Describe affected tours and manual checks in pull requests; include screenshots only for visible worksheet or dashboard changes. Do not commit live player data or year-specific working folders.
