# LGS Season Manager

Small local web application to track La Grande Semaine across multiple years.

> **Architecture Overview**: See [architecture.mmd](./architecture.mmd) for a complete flowchart of the application workflow, data flow, and user interactions.

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

### Source modes

- **Mode local**: select **Lier le dossier LGS** to choose `ASGLM <year>/LGS` from the local filesystem (Edge/Chrome).
- **Mode Dropbox cle navigateur**: select **Connecter Dropbox**, provide a Dropbox access token, and set the season folder path (for example `/ASGLM 2026/LGS`).
- **Mode Dropbox serveur**: select **Configurer serveur Dropbox** and set the backend URL and season path. This mode requires a dedicated server implementation.

Security note: in browser-key mode, the Dropbox token is kept only for the current browser session and is never exported in the season JSON file.

In both modes, the app scans `T1` to `T6` and `Finale` and records detected spreadsheets.

**Ajouter fichier XLS** is available on every tour. In local mode, files are copied into the selected `LGS` folder. In Dropbox browser-key mode, files are uploaded to the configured Dropbox tour folder. Existing files are preserved via automatic rename when needed.

For the pre-linked 2023, 2024, and 2025 archives, **Ouvrir le fichier RMS** opens the catalogued local file directly. For any other season, link its folder first; access granted by the browser is only retained for the current session.

The 2023, 2024, and 2025 seasons are pre-linked to export inventories. Use **Lier le dossier LGS** (local) or **Analyser Dropbox** (Dropbox browser-key mode) when a folder changes or when using a different copy of an archive.

## Scope

This version tracks each season's seven events (`Tour 1` to `Tour 6` and `Finale`), RMS export file references, validation status, and season notes. Excel/VBA calculation remains the source of truth for scores and rankings.

## Application Flow

The dashboard follows this sequence:

1. **Initialization** — Load seasons from browser storage or create new ones
2. **Source Linking** — Link a local `ASGLM <year>/LGS` folder, connect Dropbox with a browser-managed key, or use a Dropbox backend server
3. **File Discovery** — Scan tour folders (`T1`–`T6`, `Finale`) for export files
4. **Data Parsing** — Extract tour information and standings from Excel files
5. **Rendering** — Display standings grouped by series and score type (Brut/Net)
6. **User Interactions** — Navigate between tours, update status, manage files, and export seasons

For a detailed visual breakdown, see [architecture.mmd](./architecture.mmd).
