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
- **Mode Dropbox**: select **Configurer Dropbox**, set:
  - backend URL (for example `https://lgs-api.example.com/api/dropbox`)
  - proxy access key (shared by admin)
  - season folder path (for example `/ASGLM 2019/ASGLM 2026/LGS`).

Security note: Dropbox credentials are kept on the server only (environment variable).
The browser only receives a proxy access key (via shared link or prompt), never the Dropbox token.

In both modes, the app scans `T1` to `T6` and `Finale` and records detected spreadsheets.

**Ajouter fichier XLS** is available on every tour. In local mode, files are copied into the selected `LGS` folder. In Dropbox mode, files are uploaded through the server to the configured Dropbox tour folder. Existing files are preserved via automatic rename when needed.

## Dropbox server setup

Run a small proxy server from this repository:

```bash
set DROPBOX_ACCESS_TOKEN=<your_dropbox_access_token>
set LGS_PROXY_ACCESS_KEY=<shared_proxy_key>
set ALLOWED_ORIGINS=https://joffroy59.github.io,file://
set DROPBOX_ALLOWED_ROOT=/ASGLM
set ENABLE_DROPBOX_UPLOAD=true
set PORT=8787
node server/dropbox-proxy.js
```

Then in the app, choose **Mode Dropbox** and set:

- Server URL: `http://localhost:8787/api/dropbox` (or your hosted server URL)
- Proxy access key: value of `LGS_PROXY_ACCESS_KEY`
- Dropbox path: `/ASGLM 2019/ASGLM 2026/LGS` (or other season root)

For the pre-linked 2023, 2024, and 2025 archives, **Ouvrir le fichier RMS** opens the catalogued local file directly. For any other season, link its folder first; access granted by the browser is only retained for the current session.

The 2023, 2024, and 2025 seasons are pre-linked to export inventories. Use **Lier le dossier LGS** (local) or **Analyser Dropbox** (Dropbox mode) when a folder changes or when using a different copy of an archive.

## Scope

This version tracks each season's seven events (`Tour 1` to `Tour 6` and `Finale`), RMS export file references, validation status, and season notes. Excel/VBA calculation remains the source of truth for scores and rankings.

## Application Flow

The dashboard follows this sequence:

1. **Initialization** — Load seasons from browser storage or create new ones
2. **Source Linking** — Link a local `ASGLM <year>/LGS` folder or connect Dropbox
3. **File Discovery** — Scan tour folders (`T1`–`T6`, `Finale`) for export files
4. **Data Parsing** — Extract tour information and standings from Excel files
5. **Rendering** — Display standings grouped by series and score type (Brut/Net)
6. **User Interactions** — Navigate between tours, update status, manage files, and export seasons

For a detailed visual breakdown, see [architecture.mmd](./architecture.mmd).
