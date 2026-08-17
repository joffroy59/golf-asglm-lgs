# LGS Season Manager

Small local web application to track La Grande Semaine across multiple years.

## Run locally

Open `app/index.html` in a modern browser. No installation, server, or dependency is required.

## Initialize a new year

1. Create the parent directory `..\ASGLM <year>`.
2. Run `initYear.bat` from the repository root and enter the year when prompted.
3. In the dashboard, create the same season and link `..\ASGLM <year>\LGS`.

See the [main project README](../readme.md) for the full RMS import procedure.

## Data

The application stores season data in the browser's local storage. Use **Exporter la saison** to save `lgs-saison-<year>.json` in the corresponding `ASGLM <year>/LGS` directory. Import that file to restore the season on another computer or browser.

Deleting a season never deletes Excel files. To prevent accidental removal of the local dashboard record, type the displayed year in the confirmation dialog.

To link historic Excel data, select the relevant season then choose **Lier le dossier LGS**. In Microsoft Edge or Google Chrome, select `ASGLM <year>/LGS`; the application reads the `T1` to `T7` and `Finale` folders and records the spreadsheet names it finds. No file is uploaded; writing occurs only after an explicit **Ajouter fichier XLS** action.

After linking a folder, **Ajouter fichier XLS** is enabled on each tour. Select the day's `.xls` or `.xlsx` export and it is copied into the matching `T1` to `T7` or `Finale` directory. Existing files are preserved; duplicate names receive a number such as ` (2)`.

For the pre-linked 2023, 2024, and 2025 archives, **Ouvrir le fichier RMS** opens the catalogued local file directly. For any other season, link its folder first; access granted by the browser is only retained for the current session.

The 2023, 2024, and 2025 seasons are pre-linked to the export inventories currently stored in their `ASGLM <year>/LGS` folders. Use **Lier le dossier LGS** again when a folder changes or when using a different copy of an archive.

## Scope

This first version tracks each season's eight events (`Tour 1` to `Tour 7` and `Finale`), RMS export file references, validation status, and season notes. Excel/VBA calculation remains the source of truth for scores and rankings.
