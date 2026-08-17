# Copilot Instructions for golf-asglm-lgs

## Project Overview

**La Grande Semaine (LGS)** is a golf competition tracking system with two independent components:

1. **Excel/VBA Engine** - The calculation source of truth for scores and rankings (`Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm`)
2. **Browser Dashboard** - A local-first web app (`app/`) for multi-year season management, RMS export tracking, and validation status

The Excel workbook contains player data imports and score calculations. The app provides a UI to organize seasons across years, link folder structures, and manage RMS export files without modifying Excel calculation logic.

## Build, Test, and Development

### Setup: VBA Export Git Hooks

VBA modules live in exported `.bas` and `.cls` files. The pre-commit hook automatically extracts them on each commit:

```bash
initGit.bat                    # Install Git hooks (run once from repo root)
pip install -U oletools       # Required dependency for hook
```

The hook runs [`.hooks/pre-commit.py`](<.hooks/pre-commit.py>) which uses `oletools` to dump VBA from `Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm` into `.bas`/`.cls` files in `vba/Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm.vba/`.

### Year Initialization

```bash
initYear.bat                   # Create new season workspace
# Prompts for year, creates ../ASGLM <YEAR>/LGS with tour folders T1-T7, Finale, Backup, Poub
```

### Season Setup in Dashboard

1. Open [`app/index.html`](<app/index.html>) in a modern browser (no server required)
2. Click "Nouvelle saison" and enter the year
3. Click "Lier le dossier LGS" and select the `../ASGLM <YEAR>/LGS` folder
4. The app scans for tour folders and caches Excel file references locally

### Testing

There is no automated test suite. Validate changes manually:

- **VBA Changes**: Import sample export [`fichier exemple nom export FFG/2d. Extraction XLS globale.xls`](<fichier exemple nom export FFG>) into a test workbook and check results in Homme and Dame worksheets
- **Dashboard Changes**: Create a season in the browser, update a tour status/file, export JSON, then import it and verify data persists
- **Year Initialization**: Run `initYear.bat`, verify folder structure in `../ASGLM <YEAR>/LGS`, and check that workbook copies exist

### Backup Utility

```bash
backup.bat                     # Archives "2d. Extraction XLS globale.xls" to dated subfolder
```

Run from within a tour directory (T1-T7 or Finale) to archive export files.

## Architecture and Key Concepts

### Excel Workbook Structure

- **`ThisWorkbook.cls`** - Event handlers for UI interactions
- **`import.bas`** - Parses RMS/FFG export files; handles column mapping and data extraction
- **`integration.bas`** - Merges imported data into scoresheet; calculates nets and rankings
- **`clean.bas`** - Removes imports; resets sheets for fresh data
- **`initConstante.bas`** - Initializes constants (tour names, column indices, version)
- **`import_lib.bas`** - Helper functions for file I/O and data transformation
- **`lib.bas`** - Utility procedures (string manipulation, array operations)
- **`lib_excel.bas`** - Excel API wrappers (range navigation, formatting)
- **`process.bas`** - Orchestrates multi-tour import workflows
- **`History.bas`** - Tracks calculation version and audit logs
- **`tools.bas`** - Miscellaneous debugging and diagnostic procedures
- **`mock.bas`** - Test data generators for development

Each tour workbook (T1-T7, Finale) is a copy of the main template. They share the same module names and structure.

### Browser Dashboard Structure

- **[`index.html`](<app/index.html>)** - Semantic markup for season selector, tour grid, notes, dialogs
- **[`app.js`](<app/app.js>)** - State management (localStorage-based), event listeners, UI updates
- **[`styles.css`](<app/styles.css>)** - Plain CSS; no build step or framework
- **[`historical-data.js`](<app/historical-data.js>)** - Pre-linked 2023, 2024, 2025 folder mappings

The app stores all data in `localStorage` under key `lgs-season-manager-v1`. It exports seasons as `lgs-saison-<year>.json` files for backup and cross-device import.

## Key Conventions

### Language and Naming

- Keep existing **French business labels** (tour names, worksheet tabs, RMS field names)
- Use **English for code logic** (variable names, procedures, comments)
- VBA procedure names: descriptive, camelCase or PascalCase (not abbreviated)
- Don't rename exported module files unless the Excel workbook module name changes

### VBA Style

- Use **4-space indentation** throughout
- Declare all variables explicitly (`Option Explicit` at module top)
- Use `ByRef` for mutable parameters; `ByVal` for immutable
- Comment complex loops and non-obvious logic

### Browser App

- Use **semantic HTML** (`<section>`, `<article>`, `<button>`, etc.)
- Keep it **dependency-free** — plain CSS and vanilla JavaScript only
- Preserve **local-first** behavior — no server, all data in browser storage
- File operations (link folder, add Excel file) use browser File Access API; write-access is local only

### Version Management

- Workbook version lives in `init.bat` as `VERSION=2.12` and in VBA as constants in `initConstante.bas`
- Update both when releasing a new version
- Commit version bumps separately from feature work

### Metadata vs. Calculation

Keep **Excel calculations** (scores, rankings, nets) separate from **app metadata** (season notes, tour status, file references). The app never modifies Excel files directly; it only reads folder contents and caches file names. Calculation updates require manual RMS import into the workbook.

### Commits and Pull Requests

- Use short, scoped messages: `fix: update RMS import parsing`, `docs: clarify initYear setup`
- Describe affected tours and manual checks in PR descriptions
- Include **screenshots only for visible worksheet or dashboard changes**
- Don't commit live player data or year-specific working folders (they live outside the repo at `../ASGLM <YEAR>/LGS/`)

## Important Files and Paths

- **Main Workbook** — [`Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm`](<Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm>)
- **VBA Export Directory** — [`vba/Calcul La Grande Semaine - STROKEPLAY - Tn - HOMME_OU_DAME_v2.12.xlsm.vba/`](<vba>)
- **Sample Exports** — [`fichier exemple nom export FFG/`](<fichier exemple nom export FFG>)
- **Dashboard** — [`app/`](<app>)
- **Pre-commit Hook** — [`.hooks/pre-commit.py`](<.hooks/pre-commit.py>)
- **Documentation** — [`AGENTS.md`](<AGENTS.md>) (repository guidelines), [`git.md`](<git.md>) (VBA export details)
