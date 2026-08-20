# Dropbox proxy server

This server keeps Dropbox credentials on the server side and exposes a minimal API for the browser app.

## Endpoints

- `POST /api/dropbox/list-folder`
- `POST /api/dropbox/temporary-link`
- `POST /api/dropbox/download`
- `POST /api/dropbox/upload`

## Run

```bash
set DROPBOX_ACCESS_TOKEN=<your_dropbox_access_token>
set LGS_PROXY_ACCESS_KEY=<shared_proxy_key>
set ALLOWED_ORIGINS=https://joffroy59.github.io
set DROPBOX_ALLOWED_ROOT=/ASGLM
set ENABLE_DROPBOX_UPLOAD=true
set PORT=8787
node server/dropbox-proxy.js
```

## Notes

- Do **not** commit real credentials.
- Restrict `ALLOWED_ORIGINS` to trusted origins (no wildcard).
- Keep `LGS_PROXY_ACCESS_KEY` private and rotate it if leaked.
- Restrict `DROPBOX_ALLOWED_ROOT` to the minimal path scope required.
- Set `ENABLE_DROPBOX_UPLOAD=false` when read-only mode is preferred.
- If hosted publicly, protect access with your reverse proxy/network rules.
