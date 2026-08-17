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

To link historic Excel data, select the relevant season then choose **Lier le dossier LGS**. In Microsoft Edge or Google Chrome, select `ASGLM <year>/LGS`; the application reads the `T1` to `T7` and `Finale` folders and records the spreadsheet names it finds. The files remain read-only and are never uploaded.

## Scope

This first version tracks each season's eight events (`Tour 1` to `Tour 7` and `Finale`), RMS export file references, validation status, and season notes. Excel/VBA calculation remains the source of truth for scores and rankings.
