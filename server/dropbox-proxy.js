/**
 * LGS Dropbox proxy server
 * - Keeps Dropbox credentials on server side
 * - Exposes minimal endpoints for the front-end app
 *
 * Usage:
 *   set DROPBOX_ACCESS_TOKEN=<token>
 *   set LGS_PROXY_ACCESS_KEY=<shared_proxy_key>
 *   set ALLOWED_ORIGINS=https://joffroy59.github.io
 *   set DROPBOX_ALLOWED_ROOT=/ASGLM
 *   set ENABLE_DROPBOX_UPLOAD=true
 *   set PORT=8787
 *   node server\\dropbox-proxy.js
 */

const PORT = Number(process.env.PORT || 8787);
const DROPBOX_ACCESS_TOKEN = (process.env.DROPBOX_ACCESS_TOKEN || "").trim();
const PROXY_ACCESS_KEY = (process.env.LGS_PROXY_ACCESS_KEY || "").trim();
const DROPBOX_ALLOWED_ROOT = String(process.env.DROPBOX_ALLOWED_ROOT || "/ASGLM")
  .trim()
  .replace(/\/+$/, "");
const ENABLE_DROPBOX_UPLOAD = String(process.env.ENABLE_DROPBOX_UPLOAD || "false").toLowerCase() === "true";
const ALLOWED_ORIGINS = String(process.env.ALLOWED_ORIGINS || "")
  .split(",")
  .map((value) => value.trim())
  .filter(Boolean);

if (!DROPBOX_ACCESS_TOKEN) {
  console.error("Missing DROPBOX_ACCESS_TOKEN environment variable.");
  process.exit(1);
}
if (!PROXY_ACCESS_KEY) {
  console.error("Missing LGS_PROXY_ACCESS_KEY environment variable.");
  process.exit(1);
}
if (!ALLOWED_ORIGINS.length || ALLOWED_ORIGINS.includes("*")) {
  console.error("ALLOWED_ORIGINS must be explicitly defined and must not contain '*'.");
  process.exit(1);
}
if (!DROPBOX_ALLOWED_ROOT.startsWith("/")) {
  console.error("DROPBOX_ALLOWED_ROOT must start with '/'.");
  process.exit(1);
}

function buildCorsHeaders(origin) {
  const allowOrigin = origin && ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-LGS-Proxy-Key",
    "Access-Control-Max-Age": "86400",
    Vary: "Origin"
  };
}

async function parseJsonBody(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  if (!chunks.length) return {};
  const raw = Buffer.concat(chunks).toString("utf8");
  return JSON.parse(raw);
}

async function callDropboxJson(endpoint, payload) {
  const response = await fetch(`https://api.dropboxapi.com/2${endpoint}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${DROPBOX_ACCESS_TOKEN}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
  const text = await response.text();
  let json = null;
  try { json = text ? JSON.parse(text) : null; } catch (_) { json = null; }
  return { ok: response.ok, status: response.status, text, json };
}

async function callDropboxDownload(path) {
  const response = await fetch("https://content.dropboxapi.com/2/files/download", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${DROPBOX_ACCESS_TOKEN}`,
      "Dropbox-API-Arg": JSON.stringify({ path })
    }
  });
  const buffer = Buffer.from(await response.arrayBuffer());
  return {
    ok: response.ok,
    status: response.status,
    headers: response.headers,
    buffer
  };
}

async function callDropboxUpload(path, contentBuffer) {
  const response = await fetch("https://content.dropboxapi.com/2/files/upload", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${DROPBOX_ACCESS_TOKEN}`,
      "Content-Type": "application/octet-stream",
      "Dropbox-API-Arg": JSON.stringify({
        path,
        mode: "add",
        autorename: true,
        mute: false,
        strict_conflict: false
      })
    },
    body: contentBuffer
  });
  const text = await response.text();
  let json = null;
  try { json = text ? JSON.parse(text) : null; } catch (_) { json = null; }
  return { ok: response.ok, status: response.status, text, json };
}

function normalizeDropboxPath(path) {
  const normalized = String(path || "").trim().replace(/\\/g, "/");
  if (!normalized) return "";
  return normalized.startsWith("/") ? normalized : `/${normalized}`;
}

function isPathAllowed(path) {
  const normalized = normalizeDropboxPath(path);
  if (!normalized) return false;
  return normalized === DROPBOX_ALLOWED_ROOT || normalized.startsWith(`${DROPBOX_ALLOWED_ROOT}/`);
}

function sendJson(response, status, payload, corsHeaders) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    ...corsHeaders
  });
  response.end(JSON.stringify(payload));
}

const server = require("http").createServer(async (request, response) => {
  const origin = request.headers.origin;
  const corsHeaders = buildCorsHeaders(origin);

  if (request.method === "OPTIONS") {
    response.writeHead(204, corsHeaders);
    response.end();
    return;
  }

  if (request.method !== "POST") {
    sendJson(response, 405, { error: "METHOD_NOT_ALLOWED" }, corsHeaders);
    return;
  }

  const incomingKey = String(request.headers["x-lgs-proxy-key"] || "").trim();
  if (incomingKey !== PROXY_ACCESS_KEY) {
    sendJson(response, 401, { error: "UNAUTHORIZED" }, corsHeaders);
    return;
  }

  try {
    const body = await parseJsonBody(request);
    const url = new URL(request.url, `http://${request.headers.host}`);
    const pathname = url.pathname.replace(/\/+$/, "");

    if (pathname === "/api/dropbox/list-folder") {
      const path = String(body.path || "").trim();
      if (!isPathAllowed(path)) {
        sendJson(response, 403, { error: "PATH_OUT_OF_SCOPE" }, corsHeaders);
        return;
      }
      const result = await callDropboxJson("/files/list_folder", { path });
      if (!result.ok) {
        const isNotFound = /path\/not_found/i.test(result.text);
        sendJson(response, isNotFound ? 404 : 502, {
          error: isNotFound ? "DROPBOX_PATH_NOT_FOUND" : "DROPBOX_LIST_FAILED",
          details: result.text
        }, corsHeaders);
        return;
      }
      sendJson(response, 200, { entries: result.json?.entries || [] }, corsHeaders);
      return;
    }

    if (pathname === "/api/dropbox/temporary-link") {
      const path = String(body.path || "").trim();
      if (!isPathAllowed(path)) {
        sendJson(response, 403, { error: "PATH_OUT_OF_SCOPE" }, corsHeaders);
        return;
      }
      const result = await callDropboxJson("/files/get_temporary_link", { path });
      if (!result.ok) {
        sendJson(response, 502, { error: "DROPBOX_TEMP_LINK_FAILED", details: result.text }, corsHeaders);
        return;
      }
      sendJson(response, 200, { link: result.json?.link || "" }, corsHeaders);
      return;
    }

    if (pathname === "/api/dropbox/download") {
      const path = String(body.path || "").trim();
      if (!isPathAllowed(path)) {
        sendJson(response, 403, { error: "PATH_OUT_OF_SCOPE" }, corsHeaders);
        return;
      }
      const result = await callDropboxDownload(path);
      if (!result.ok) {
        response.writeHead(502, { "Content-Type": "text/plain; charset=utf-8", ...corsHeaders });
        response.end("DROPBOX_DOWNLOAD_FAILED");
        return;
      }
      const apiArg = result.headers.get("dropbox-api-result");
      let fileName = path.split("/").pop() || "download.bin";
      if (apiArg) {
        try {
          const parsed = JSON.parse(apiArg);
          if (parsed?.name) fileName = parsed.name;
        } catch (_) {
          // Keep fallback file name.
        }
      }
      response.writeHead(200, {
        "Content-Type": result.headers.get("content-type") || "application/octet-stream",
        "Content-Length": String(result.buffer.length),
        "X-Dropbox-File-Name": fileName,
        ...corsHeaders
      });
      response.end(result.buffer);
      return;
    }

    if (pathname === "/api/dropbox/upload") {
      if (!ENABLE_DROPBOX_UPLOAD) {
        sendJson(response, 403, { error: "UPLOAD_DISABLED" }, corsHeaders);
        return;
      }
      const path = String(body.path || "").trim();
      const contentBase64 = String(body.contentBase64 || "");
      if (!isPathAllowed(path)) {
        sendJson(response, 403, { error: "PATH_OUT_OF_SCOPE" }, corsHeaders);
        return;
      }
      if (!path || !contentBase64) {
        sendJson(response, 400, { error: "INVALID_UPLOAD_PAYLOAD" }, corsHeaders);
        return;
      }
      const contentBuffer = Buffer.from(contentBase64, "base64");
      const result = await callDropboxUpload(path, contentBuffer);
      if (!result.ok) {
        sendJson(response, 502, { error: "DROPBOX_UPLOAD_FAILED", details: result.text }, corsHeaders);
        return;
      }
      sendJson(response, 200, { name: result.json?.name || path.split("/").pop() }, corsHeaders);
      return;
    }

    sendJson(response, 404, { error: "NOT_FOUND" }, corsHeaders);
  } catch (error) {
    sendJson(response, 500, { error: "SERVER_ERROR", message: error.message }, corsHeaders);
  }
});

server.listen(PORT, () => {
  console.log(`Dropbox proxy server listening on port ${PORT}`);
  console.log(`Allowed origins: ${ALLOWED_ORIGINS.join(", ")}`);
  console.log(`Allowed Dropbox root: ${DROPBOX_ALLOWED_ROOT}`);
  console.log(`Upload enabled: ${ENABLE_DROPBOX_UPLOAD}`);
});
