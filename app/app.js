const STORAGE_KEY = "lgs-season-manager-v1";
const HISTORICAL_YEARS = [2023, 2024, 2025];
const TOUR_NAMES = ["Tour 1", "Tour 2", "Tour 3", "Tour 4", "Tour 5", "Tour 6", "Finale"];
const SOURCE_MODES = {
  local: "local",
  dropbox: "dropbox"
};
const DROPBOX_TOKEN_SESSION_KEY = "lgs-dropbox-access-token-v1";
const STATUS_LABELS = {
  planned: "A preparer",
  ready: "Export pret",
  imported: "Import realise",
  validated: "Valide"
};

let currentStandingsFile = null;

// Toggle section collapse/expand
function toggleSection(sectionId) {
  const section = document.querySelector(`[data-section="${sectionId}"]`);
  if (section) {
    section.classList.toggle('collapsed');
    localStorage.setItem(`section-collapsed-${sectionId}`, section.classList.contains('collapsed'));
  }
}

// Restore section collapse states from localStorage
function restoreSectionStates() {
  ['season-overview', 'tour-section', 'standings-section', 'statistics-section', 'notes-section', 'setup-section'].forEach(sectionId => {
    const isCollapsed = localStorage.getItem(`section-collapsed-${sectionId}`) === 'true';
    const section = document.querySelector(`[data-section="${sectionId}"]`);
    if (section && isCollapsed) {
      section.classList.add('collapsed');
    }
  });
}

const elements = {
  seasonSelect: document.querySelector("#season-select"),
  seasonTitle: document.querySelector("#season-title"),
  seasonPath: document.querySelector("#season-path"),
  sourceInfo: document.querySelector("#source-info"),
  sourceLocalButton: document.querySelector("#source-local-button"),
  sourceDropboxButton: document.querySelector("#source-dropbox-button"),
  linkFolderButton: document.querySelector("#link-folder-button"),
  tourGrid: document.querySelector("#tour-grid"),
  notes: document.querySelector("#season-notes"),
  progressLabel: document.querySelector("#progress-label"),
  progressBar: document.querySelector("#progress-bar"),
  scanResult: document.querySelector("#scan-result"),
  dialog: document.querySelector("#season-dialog"),
  form: document.querySelector("#season-form"),
  importInput: document.querySelector("#import-input"),
  deleteDialog: document.querySelector("#delete-season-dialog"),
  deleteForm: document.querySelector("#delete-season-form"),
  deleteYearLabel: document.querySelector("#delete-year-label"),
  deleteYearInput: document.querySelector("#delete-year-input"),
  confirmDeleteButton: document.querySelector("#confirm-delete-button"),
  deleteSeasonButton: document.querySelector("#delete-season-button"),
  standingsContainer: document.querySelector("#standings-container"),
  standingsStatus: document.querySelector("#standings-status"),
  refreshStandingsButton: document.querySelector("#refresh-standings-button")
};

function makeSeason(year, directory) {
  const seasonYear = Number(year);
  return {
    id: crypto.randomUUID(),
    year: seasonYear,
    directory,
    sourceMode: SOURCE_MODES.local,
    dropboxPath: `/ASGLM ${seasonYear}/LGS`,
    sourceMessage: "",
    notes: "",
    tours: TOUR_NAMES.map((name, index) => ({
      name,
      number: index + 1,
      status: "planned",
      file: "",
      note: ""
    }))
  };
}

function ensureSeasonDefaults(season) {
  season.sourceMode = season.sourceMode === SOURCE_MODES.dropbox ? SOURCE_MODES.dropbox : SOURCE_MODES.local;
  if (typeof season.dropboxPath !== "string" || !season.dropboxPath.trim()) {
    season.dropboxPath = `/ASGLM ${season.year}/LGS`;
  }
  if (typeof season.sourceMessage !== "string") season.sourceMessage = "";
  if (season.sourceMessage === "Mode Dropbox actif. La connexion Dropbox API sera ajoutee dans une prochaine mise a jour.") {
    season.sourceMessage = "";
  }
}

function loadState() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (parsed?.seasons?.length) return prepareState(parsed);
  } catch (_) {
    // A corrupted local record is replaced with a clean first season.
  }
  return prepareState({ activeId: "", seasons: [] });
}

function prepareState(savedState) {
  savedState.migrations ||= {};
  if (!savedState.migrations.removedUninitialized2026) {
    savedState.seasons = savedState.seasons.filter((season) => !isUninitialized2026(season));
    savedState.migrations.removedUninitialized2026 = true;
  }
  addHistoricalSeasons(savedState);
  savedState.seasons.forEach(ensureSeasonDefaults);
  if (!savedState.seasons.some((season) => season.id === savedState.activeId)) {
    savedState.activeId = savedState.seasons.find((season) => season.year === 2025)?.id || savedState.seasons[0]?.id;
  }
  return savedState;
}

function isUninitialized2026(season) {
  return season.year === 2026
    && season.directory === "..\\ASGLM 2026\\LGS"
    && !season.notes
    && season.tours.every((tour) => !tour.file && !tour.note && tour.status === "planned");
}

function addHistoricalSeasons(savedState) {
  HISTORICAL_YEARS.forEach((year) => {
    if (!savedState.seasons.some((season) => season.year === year)) {
      savedState.seasons.push(makeSeason(year, `..\\ASGLM ${year}\\LGS`));
    }
  });
  Object.entries(window.LGS_HISTORICAL_SEASON_DATA || {}).forEach(([year, catalog]) => {
    const historicSeason = savedState.seasons.find((season) => season.year === Number(year));
    const canApplyCatalog = historicSeason?.tours.every((tour) => (
      !tour.sourceFiles?.length && !tour.file && !tour.note && tour.status === "planned"
    ));
    if (historicSeason && canApplyCatalog && !historicSeason.lastScan && !historicSeason.catalogMessage) {
      historicSeason.directory = catalog.directory;
      historicSeason.tours.forEach((tour) => {
        const files = catalog.tours[tour.number] || [];
        tour.sourceFiles = files;
        tour.file = files[0] || "";
        if (files.length) tour.status = "imported";
      });
      historicSeason.catalogMessage = catalog.message;
    }
  });
  return savedState;
}

let state = loadState();
const linkedFileHandles = new Map();
const linkedDirectoryHandles = new Map();
let dropboxAccessToken = "";
try {
  dropboxAccessToken = sessionStorage.getItem(DROPBOX_TOKEN_SESSION_KEY) || "";
} catch (_) {
  dropboxAccessToken = "";
}


function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function hasDropboxToken() {
  return Boolean(dropboxAccessToken);
}

function setDropboxToken(token) {
  dropboxAccessToken = String(token || "").trim();
  try {
    if (dropboxAccessToken) sessionStorage.setItem(DROPBOX_TOKEN_SESSION_KEY, dropboxAccessToken);
    else sessionStorage.removeItem(DROPBOX_TOKEN_SESSION_KEY);
  } catch (_) {
    // Keep token in-memory only if sessionStorage is unavailable.
  }
}

function ensureDropboxPath(path, seasonYear) {
  const trimmed = String(path || "").trim().replace(/\\/g, "/");
  const defaultPath = `/ASGLM ${seasonYear}/LGS`;
  if (!trimmed) return defaultPath;
  const withLeadingSlash = trimmed.startsWith("/") ? trimmed : `/${trimmed}`;
  return withLeadingSlash.replace(/\/+$/, "");
}

function tourFolderName(tour) {
  return tour.name === "Finale" ? "Finale" : `T${tour.number}`;
}

async function dropboxApiJson(url, payload) {
  if (!hasDropboxToken()) throw new Error("DROPBOX_TOKEN_MISSING");
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${dropboxAccessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    const details = await response.text();
    throw new Error(`DROPBOX_API_ERROR ${response.status}: ${details}`);
  }
  return response.json();
}

async function dropboxListFolder(path) {
  try {
    const result = await dropboxApiJson("https://api.dropboxapi.com/2/files/list_folder", { path });
    return result.entries || [];
  } catch (error) {
    if (/path\/not_found/i.test(String(error.message))) return [];
    throw error;
  }
}

async function dropboxDownload(path) {
  if (!hasDropboxToken()) throw new Error("DROPBOX_TOKEN_MISSING");
  const response = await fetch("https://content.dropboxapi.com/2/files/download", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${dropboxAccessToken}`,
      "Dropbox-API-Arg": JSON.stringify({ path })
    }
  });
  if (!response.ok) {
    const details = await response.text();
    throw new Error(`DROPBOX_DOWNLOAD_ERROR ${response.status}: ${details}`);
  }
  const fileBlob = await response.blob();
  return {
    name: response.headers.get("Dropbox-API-Result") ? JSON.parse(response.headers.get("Dropbox-API-Result")).name : path.split("/").pop(),
    arrayBuffer: () => fileBlob.arrayBuffer()
  };
}

async function dropboxUpload(path, file) {
  if (!hasDropboxToken()) throw new Error("DROPBOX_TOKEN_MISSING");
  const response = await fetch("https://content.dropboxapi.com/2/files/upload", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${dropboxAccessToken}`,
      "Content-Type": "application/octet-stream",
      "Dropbox-API-Arg": JSON.stringify({
        path,
        mode: "add",
        autorename: true,
        mute: false,
        strict_conflict: false
      })
    },
    body: file
  });
  if (!response.ok) {
    const details = await response.text();
    throw new Error(`DROPBOX_UPLOAD_ERROR ${response.status}: ${details}`);
  }
  return response.json();
}

function activeSeason() {
  return state.seasons.find((season) => season.id === state.activeId) || state.seasons[0];
}

function render() {
  const season = activeSeason();
  state.activeId = season.id;
  elements.seasonSelect.replaceChildren(...state.seasons
    .slice()
    .sort((a, b) => b.year - a.year)
    .map((item) => new Option(String(item.year), item.id, false, item.id === season.id)));
  elements.seasonTitle.textContent = `Saison ${season.year}`;
  const isDropboxMode = season.sourceMode === SOURCE_MODES.dropbox;
  elements.seasonPath.textContent = isDropboxMode
    ? `Dropbox : ${season.dropboxPath}`
    : season.directory;
  elements.sourceInfo.textContent = currentSourceInfo(season);
  elements.sourceLocalButton.classList.toggle("active", !isDropboxMode);
  elements.sourceLocalButton.setAttribute("aria-pressed", String(!isDropboxMode));
  elements.sourceDropboxButton.classList.toggle("active", isDropboxMode);
  elements.sourceDropboxButton.setAttribute("aria-pressed", String(isDropboxMode));
  elements.linkFolderButton.disabled = false;
  elements.linkFolderButton.textContent = isDropboxMode
    ? (hasDropboxToken() ? "Analyser Dropbox" : "Connecter Dropbox")
    : "Lier le dossier LGS";
  elements.linkFolderButton.title = isDropboxMode
    ? "Connecter puis analyser le dossier Dropbox de cette saison."
    : "Lier le dossier LGS local pour analyser les fichiers.";
  elements.scanResult.textContent = season.sourceMessage
    || season.catalogMessage
    || (season.lastScan
      ? `Derniere analyse : ${new Date(season.lastScan).toLocaleDateString("fr-FR")}`
      : "Aucun dossier LGS analyse pour cette saison.");
  elements.notes.value = season.notes;
  renderTours(season);
  renderProgress(season);
  elements.deleteSeasonButton.disabled = state.seasons.length === 1;
  elements.deleteSeasonButton.title = elements.deleteSeasonButton.disabled
    ? "Conservez au moins une saison dans l'application."
    : "Une confirmation par annee sera demandee.";
  saveState();
  restoreSectionStates();
}

function renderTours(season) {
  const template = document.querySelector("#tour-template");
  elements.tourGrid.replaceChildren(...season.tours.map((tour) => {
    const card = template.content.firstElementChild.cloneNode(true);
    const pill = card.querySelector(".status-pill");
    card.querySelector(".tour-number").textContent = tour.name === "Finale" ? "CLOTURE" : `TOUR ${tour.number}`;
    card.querySelector(".tour-name").textContent = tour.name;
    pill.textContent = STATUS_LABELS[tour.status];
    pill.dataset.status = tour.status;
    const status = card.querySelector(".tour-status");
    const file = card.querySelector(".tour-file");
    const note = card.querySelector(".tour-note");
    const sourceSummary = card.querySelector(".source-summary");
    const openButton = card.querySelector(".open-rms-button");
    const uploadButton = card.querySelector(".upload-xls-button");
    status.value = tour.status;
    file.value = tour.file;
    note.value = tour.note;
    sourceSummary.textContent = sourceLabel(tour.sourceFiles || []);
    const isDropboxMode = season.sourceMode === SOURCE_MODES.dropbox;
    openButton.disabled = !canOpenTourFile(season, tour);
    openButton.title = openButton.disabled
      ? (isDropboxMode
        ? "Connectez Dropbox et relancez l'analyse pour ouvrir ce fichier."
        : "Liez le dossier LGS pour ouvrir ce fichier.")
      : `Ouvrir ${tour.file}`;
    uploadButton.disabled = isDropboxMode
      ? !hasDropboxToken() || !window.showOpenFilePicker
      : !window.showOpenFilePicker || !window.showDirectoryPicker;
    uploadButton.title = isDropboxMode
      ? (uploadButton.disabled
        ? "Connectez Dropbox pour ajouter un fichier."
        : `Ajouter un fichier dans ${tour.name} sur Dropbox`)
      : (uploadButton.disabled
        ? "Utilisez Microsoft Edge ou Google Chrome pour ajouter un fichier."
        : `Ajouter un fichier dans ${tour.name}`);
    status.addEventListener("change", () => updateTour(tour.number, "status", status.value));
    file.addEventListener("change", () => {
      linkedFileHandles.delete(fileHandleKey(season.id, tour.number));
      updateTour(tour.number, "file", file.value.trim());
    });
    note.addEventListener("change", () => updateTour(tour.number, "note", note.value.trim()));
    openButton.addEventListener("click", () => openRmsFile(season.id, tour.number));
    uploadButton.addEventListener("click", () => addResultFile(season.id, tour.number));
    return card;
  }));
}

function fileHandleKey(seasonId, tourNumber) {
  return `${seasonId}:${tourNumber}`;
}

function canOpenTourFile(season, tour) {
  if (season.sourceMode === SOURCE_MODES.dropbox) {
    return hasDropboxToken() && Boolean(tour.file);
  }
  return linkedFileHandles.has(fileHandleKey(season.id, tour.number)) || Boolean(knownRmsHref(season, tour));
}

function knownRmsHref(season, tour) {
  const catalog = window.LGS_HISTORICAL_SEASON_DATA?.[season.year];
  const knownFiles = catalog?.tours?.[tour.number] || [];
  if (!knownFiles.includes(tour.file)) return "";
  const folderName = tour.name === "Finale" ? "Finale" : `T${tour.number}`;
  const parts = ["..", "..", `ASGLM ${season.year}`, "LGS", folderName, tour.file];
  return parts.map(encodeURIComponent).join("/");
}

async function openRmsFile(seasonId, tourNumber) {
  const handle = linkedFileHandles.get(fileHandleKey(seasonId, tourNumber));
  const season = state.seasons.find((item) => item.id === seasonId);
  const tour = season?.tours.find((item) => item.number === tourNumber);
  if (!tour) return;
  try {
    let url = "";
    if (season.sourceMode === SOURCE_MODES.dropbox) {
      if (!hasDropboxToken()) {
        alert("Connectez Dropbox puis reessayez.");
        return;
      }
      const dropboxFilePath = `${ensureDropboxPath(season.dropboxPath, season.year)}/${tourFolderName(tour)}/${tour.file}`;
      const tempLink = await dropboxApiJson("https://api.dropboxapi.com/2/files/get_temporary_link", { path: dropboxFilePath });
      url = tempLink?.link || "";
    } else {
      url = handle
        ? URL.createObjectURL(await handle.getFile())
        : knownRmsHref(season, tour);
    }
    if (!url) return;
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.target = "_blank";
    anchor.rel = "noopener";
    anchor.click();
    if (handle && season.sourceMode === SOURCE_MODES.local) setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (_) {
    alert("Le fichier RMS ne peut pas etre ouvert. Verifiez la source puis reessayez.");
  }
}

async function addResultFile(seasonId, tourNumber) {
  let root = linkedDirectoryHandles.get(seasonId);
  const season = state.seasons.find((item) => item.id === seasonId);
  const tour = season?.tours.find((item) => item.number === tourNumber);
  if (season?.sourceMode === SOURCE_MODES.dropbox) {
    if (!season || !tour || !window.showOpenFilePicker) {
      alert("Utilisez Microsoft Edge ou Google Chrome pour ajouter un fichier.");
      return;
    }
    if (!hasDropboxToken()) {
      const connected = await connectAndScanDropboxSeason(season);
      if (!connected) return;
    }
    try {
      const [sourceHandle] = await window.showOpenFilePicker({
        types: [{
          description: "Fichiers Excel",
          accept: {
            "application/vnd.ms-excel": [".xls"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"]
          }
        }]
      });
      const sourceFile = await sourceHandle.getFile();
      const folderPath = `${ensureDropboxPath(season.dropboxPath, season.year)}/${tourFolderName(tour)}`;
      const uploadPath = `${folderPath}/${sourceFile.name}`;
      const uploadResult = await dropboxUpload(uploadPath, sourceFile);
      const uploadedName = uploadResult.name || sourceFile.name;
      tour.sourceFiles = [...new Set([...(tour.sourceFiles || []), uploadedName])]
        .sort((first, second) => first.localeCompare(second, "fr"));
      tour.file = uploadedName;
      tour.status = "ready";
      season.lastScan = new Date().toISOString();
      render();
      elements.scanResult.textContent = `${uploadedName} ajoute dans Dropbox (${tourFolderName(tour)}).`;
    } catch (error) {
      if (error.name !== "AbortError") {
        console.error(error);
        alert("L'ajout du fichier XLS sur Dropbox a echoue.");
      }
    }
    return;
  }
  if (!season || !tour || !window.showOpenFilePicker) {
    alert("Utilisez Microsoft Edge ou Google Chrome pour ajouter un fichier.");
    return;
  }
  try {
    if (!root) {
      const linked = await linkSeasonFolder();
      if (!linked) return;
      root = linkedDirectoryHandles.get(seasonId);
    }
    const [sourceHandle] = await window.showOpenFilePicker({
      types: [{
        description: "Fichiers Excel",
        accept: {
          "application/vnd.ms-excel": [".xls"],
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
        }
      }]
    });
    const sourceFile = await sourceHandle.getFile();
    const folderName = tour.name === "Finale" ? "Finale" : `T${tour.number}`;
    const folder = await root.getDirectoryHandle(folderName, { create: true });
    const destinationHandle = await nextAvailableFileHandle(folder, sourceFile.name);
    const destinationName = destinationHandle.name;
    const writable = await destinationHandle.createWritable();
    await writable.write(sourceFile);
    await writable.close();

    tour.sourceFiles = [...new Set([...(tour.sourceFiles || []), destinationName])]
      .sort((first, second) => first.localeCompare(second, "fr"));
    tour.file = destinationName;
    tour.status = "ready";
    linkedFileHandles.set(fileHandleKey(season.id, tour.number), destinationHandle);
    render();
    elements.scanResult.textContent = `${destinationName} ajoute dans ${folderName}.`;
  } catch (error) {
    if (error.name !== "AbortError") alert("L'ajout du fichier XLS a echoue.");
  }
}

async function nextAvailableFileHandle(folder, originalName) {
  const extensionIndex = originalName.lastIndexOf(".");
  const baseName = extensionIndex > 0 ? originalName.slice(0, extensionIndex) : originalName;
  const extension = extensionIndex > 0 ? originalName.slice(extensionIndex) : "";
  for (let copyNumber = 1; ; copyNumber += 1) {
    const candidate = copyNumber === 1 ? originalName : `${baseName} (${copyNumber})${extension}`;
    try {
      await folder.getFileHandle(candidate);
    } catch (error) {
      if (error.name === "NotFoundError") return folder.getFileHandle(candidate, { create: true });
      throw error;
    }
  }
}

function sourceLabel(files) {
  if (!files.length) return "Aucune donnee source detectee";
  if (files.length === 1) return `Donnee source : ${files[0]}`;
  return `${files.length} fichiers sources, dont ${files[0]}`;
}

function currentSourceInfo(season) {
  if (season.sourceMode === SOURCE_MODES.dropbox) {
    return hasDropboxToken()
      ? `Ressource active : Dropbox (${season.dropboxPath})`
      : `Ressource active : Dropbox non connecte (${season.dropboxPath})`;
  }
  const linkedRoot = linkedDirectoryHandles.get(season.id);
  return linkedRoot
    ? `Ressource active : Local (${linkedRoot.name})`
    : "Ressource active : Local (non lie)";
}

function setSourceMode(mode) {
  const season = activeSeason();
  if (season.sourceMode === mode) return;
  season.sourceMode = mode;
  season.sourceMessage = mode === SOURCE_MODES.dropbox
    ? "Mode Dropbox actif. Cliquez sur \"Connecter Dropbox\" pour analyser les fichiers."
    : "Mode local actif. Utilisez \"Lier le dossier LGS\" pour scanner les fichiers.";
  if (mode === SOURCE_MODES.dropbox) {
    linkedDirectoryHandles.delete(season.id);
    season.lastScan = "";
    season.catalogMessage = "";
  }
  render();
}

async function connectAndScanDropboxSeason(season) {
  const existingTokenHint = hasDropboxToken()
    ? "Token deja charge pour cette session navigateur. Laissez vide pour le conserver."
    : "Le token est conserve uniquement pour la session en cours (jamais exporte).";
  const tokenInput = window.prompt(`Access token Dropbox (Scoped App)\n${existingTokenHint}`, "");
  if (tokenInput === null) return false;
  if (tokenInput.trim()) setDropboxToken(tokenInput.trim());
  if (!hasDropboxToken()) {
    alert("Un access token Dropbox est requis.");
    return false;
  }

  const currentPath = ensureDropboxPath(season.dropboxPath, season.year);
  const pathInput = window.prompt("Chemin Dropbox du dossier LGS pour cette saison", currentPath);
  if (pathInput === null) return false;
  season.dropboxPath = ensureDropboxPath(pathInput, season.year);

  try {
    await scanDropboxSeason(season);
    season.sourceMessage = "";
    return true;
  } catch (error) {
    console.error(error);
    season.sourceMessage = "Connexion Dropbox echouee. Verifiez le token et le chemin.";
    alert("Impossible de lire Dropbox. Verifiez le token et le chemin du dossier LGS.");
    return false;
  }
}

async function scanDropboxSeason(season) {
  let detectedCount = 0;
  const basePath = ensureDropboxPath(season.dropboxPath, season.year);
  for (const tour of season.tours) {
    const folderPath = `${basePath}/${tourFolderName(tour)}`;
    const entries = await dropboxListFolder(folderPath);
    const excelEntries = entries
      .filter((entry) => entry[".tag"] === "file" && /\.xls[xm]?$/i.test(entry.name))
      .sort((first, second) => first.name.localeCompare(second.name, "fr"));
    const files = excelEntries.map((entry) => entry.name);
    tour.sourceFiles = files;
    if (files.length) {
      detectedCount += files.length;
      const rmsFile = excelEntries.find((entry) => /extraction/i.test(entry.name)) || excelEntries[0];
      tour.file = rmsFile.name;
      if (files.some((name) => /\.xlsx?$/i.test(name))) tour.status = "imported";
      else if (tour.status === "planned") tour.status = "ready";
    }
  }
  season.lastScan = new Date().toISOString();
  season.catalogMessage = "";
  season.directory = `Dropbox : ${basePath}`;
  render();
  elements.scanResult.textContent = `${detectedCount} fichiers Excel detectes dans Dropbox.`;
}

function renderProgress(season) {
  const complete = season.tours.filter((tour) => tour.status === "validated").length;
  const percent = Math.round((complete / season.tours.length) * 100);
  elements.progressLabel.textContent = `${complete}/${season.tours.length} tours valides`;
  elements.progressBar.style.width = `${percent}%`;
}

function updateTour(number, key, value) {
  const tour = activeSeason().tours.find((item) => item.number === number);
  tour[key] = value;
  render();
}

async function linkSeasonFolder() {
  const season = activeSeason();
  if (season.sourceMode === SOURCE_MODES.dropbox) {
    return connectAndScanDropboxSeason(season);
  }
  if (!window.showDirectoryPicker) {
    alert("Utilisez Microsoft Edge ou Google Chrome pour lier un dossier LGS.");
    return false;
  }
  try {
    const root = await window.showDirectoryPicker({ mode: "readwrite" });
    const requiredFolders = ["T1", "Finale"];
    const hasLgsStructure = await Promise.all(requiredFolders.map(async (name) => {
      try {
        await root.getDirectoryHandle(name);
        return true;
      } catch (_) {
        return false;
      }
    }));
    if (!hasLgsStructure.every(Boolean)) {
      alert("Selectionnez le dossier LGS qui contient T1 a T6 et Finale.");
      return false;
    }
    linkedDirectoryHandles.set(season.id, root);
    let detectedCount = 0;
    for (const tour of season.tours) {
      const folderName = tour.name === "Finale" ? "Finale" : `T${tour.number}`;
      const folder = await root.getDirectoryHandle(folderName);
      linkedFileHandles.delete(fileHandleKey(season.id, tour.number));
      const fileEntries = [];
      for await (const entry of folder.values()) {
        if (entry.kind === "file" && /\.xls[xm]?$/i.test(entry.name)) fileEntries.push(entry);
      }
      fileEntries.sort((first, second) => first.name.localeCompare(second.name, "fr"));
      const files = fileEntries.map((entry) => entry.name);
      tour.sourceFiles = files;
      if (files.length) {
        detectedCount += files.length;
        const rmsFile = fileEntries.find((entry) => /extraction/i.test(entry.name)) || fileEntries[0];
        tour.file = rmsFile.name;
        linkedFileHandles.set(fileHandleKey(season.id, tour.number), rmsFile);
        if (files.some((name) => /\.xlsx?$/i.test(name))) tour.status = "imported";
        else if (tour.status === "planned") tour.status = "ready";
      }
    }
    season.directory = `Dossier lie : ${root.name}`;
    season.lastScan = new Date().toISOString();
    season.sourceMessage = "";
    season.catalogMessage = "";
    render();
    elements.scanResult.textContent = `${detectedCount} fichiers Excel detectes dans ${root.name}.`;
    return true;
  } catch (error) {
    if (error.name !== "AbortError") alert("La lecture du dossier LGS a echoue.");
    return false;
  }
}

function createSeason(event) {
  event.preventDefault();
  const formData = new FormData(elements.form);
  const year = Number(formData.get("year"));
  if (state.seasons.some((season) => season.year === year)) {
    alert(`La saison ${year} existe deja.`);
    return;
  }
  const season = makeSeason(year, String(formData.get("directory")).trim());
  state.seasons.push(season);
  state.activeId = season.id;
  elements.dialog.close();
  elements.form.reset();
  render();
}

function exportSeason() {
  const season = activeSeason();
  const content = JSON.stringify({ version: 1, exportedAt: new Date().toISOString(), season }, null, 2);
  const url = URL.createObjectURL(new Blob([content], { type: "application/json" }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `lgs-saison-${season.year}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function importSeason(event) {
  const [file] = event.target.files;
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const imported = JSON.parse(reader.result).season;
      if (!imported?.year || !Array.isArray(imported.tours) || imported.tours.length !== TOUR_NAMES.length) throw new Error();
      ensureSeasonDefaults(imported);
      const existing = state.seasons.findIndex((season) => season.year === imported.year);
      if (existing >= 0) state.seasons[existing] = imported;
      else state.seasons.push(imported);
      state.activeId = imported.id || crypto.randomUUID();
      if (!imported.id) imported.id = state.activeId;
      render();
    } catch (_) {
      alert("Ce fichier ne contient pas une exportation LGS valide.");
    } finally {
      elements.importInput.value = "";
    }
  };
  reader.readAsText(file);
}

async function findLatestCalculFile() {
  const season = activeSeason();
  if (season.sourceMode === SOURCE_MODES.dropbox) {
    return findLatestCalculFileDropbox(season);
  }
  if (!linkedDirectoryHandles.has(season.id)) {
    console.log("❌ No linked directory for season:", season.id);
    return null;
  }
  
  const root = linkedDirectoryHandles.get(season.id);
  const calcFiles = [];
  // RMS export handles keyed by tour label (T1–T6, Finale)
  const rmsHandles = {};
  
  console.log("🔍 Scanning linked root:", root.name);
  
  // First check the root LGS directory
  console.log("📁 Checking root directory");
  for await (const entry of root.values()) {
    console.log(`   Entry: ${entry.name} (${entry.kind})`);
    
    if (entry.kind === "file" && /\.xls[xm]?$/i.test(entry.name)) {
      console.log(`      ✓ XLS file: ${entry.name}`);
      
      // Check if it matches our pattern
      if (/Calcul La Grande Semaine/i.test(entry.name) && /HOMME_OU_DAME/i.test(entry.name)) {
        console.log(`        🎯 MATCH: Adding to calcFiles`);
        calcFiles.push({ folder: "root", handle: entry, name: entry.name });
      }
    }
  }
  
  // Then check the tour folders
  const tourFolders = ["Finale", "T6", "T5", "T4", "T3", "T2", "T1"];
  
  for (const folderName of tourFolders) {
    try {
      const folder = await root.getDirectoryHandle(folderName);
      console.log(`📁 Checking folder: ${folderName}`);
      const filesInFolder = [];
      
      for await (const entry of folder.values()) {
        console.log(`   Entry: ${entry.name} (${entry.kind})`);
        
        if (entry.kind === "file" && /\.xls[xm]?$/i.test(entry.name)) {
          filesInFolder.push(entry.name);
          console.log(`      ✓ XLS file: ${entry.name}`);
          
          // Check if it matches our pattern
          if (/Calcul La Grande Semaine/i.test(entry.name) && /HOMME_OU_DAME/i.test(entry.name)) {
            console.log(`        🎯 MATCH: Adding to calcFiles`);
            calcFiles.push({ folder: folderName, handle: entry, name: entry.name });
          }

          // Capture RMS export file ("2d. Extraction XLS globale") per tour folder
          if (/2d\. Extraction XLS globale/i.test(entry.name) && !rmsHandles[folderName]) {
            const tourKey = folderName === "Finale" ? "finale" : folderName; // T1..T6, finale
            rmsHandles[tourKey] = entry;
            console.log(`        📅 RMS file captured for ${tourKey}: ${entry.name}`);
          }
        }
      }
      
      if (filesInFolder.length === 0) {
        console.log(`   (empty folder)`);
      }
    } catch (error) {
      console.log(`⚠️  Could not access ${folderName}:`, error.message);
    }
  }
  
  console.log("✅ Total calc files found:", calcFiles.length);
  if (calcFiles.length > 0) {
    console.log("   All files found:", calcFiles.map(f => f.name).join(", "));
    
    // Priority 1: Look for Finale file
    const finaleFile = calcFiles.find(f => f.name.includes("Finale"));
    if (finaleFile) {
      console.log("   Selected: Finale file:", finaleFile.name, "from", finaleFile.folder);
      return { ...finaleFile, rmsHandles };
    }
    
    // Priority 2: Use the last file (highest tour number)
    const lastFile = calcFiles[calcFiles.length - 1];
    console.log("   Selected: Latest tour file:", lastFile.name, "from", lastFile.folder);
    return { ...lastFile, rmsHandles };
  }
  
  return null;
}

async function findLatestCalculFileDropbox(season) {
  if (!hasDropboxToken()) return null;
  const basePath = ensureDropboxPath(season.dropboxPath, season.year);
  const calcFiles = [];
  const rmsPaths = {};
  const tourFolders = ["Finale", "T6", "T5", "T4", "T3", "T2", "T1"];

  const rootEntries = await dropboxListFolder(basePath);
  for (const entry of rootEntries) {
    if (entry[".tag"] !== "file") continue;
    if (!/\.xls[xm]?$/i.test(entry.name)) continue;
    if (/Calcul La Grande Semaine/i.test(entry.name) && /HOMME_OU_DAME/i.test(entry.name)) {
      calcFiles.push({ folder: "root", path: entry.path_lower || entry.path_display, name: entry.name });
    }
  }

  for (const folderName of tourFolders) {
    const folderPath = `${basePath}/${folderName}`;
    const entries = await dropboxListFolder(folderPath);
    for (const entry of entries) {
      if (entry[".tag"] !== "file") continue;
      if (!/\.xls[xm]?$/i.test(entry.name)) continue;
      if (/Calcul La Grande Semaine/i.test(entry.name) && /HOMME_OU_DAME/i.test(entry.name)) {
        calcFiles.push({ folder: folderName, path: entry.path_lower || entry.path_display, name: entry.name });
      }
      if (/2d\. Extraction XLS globale/i.test(entry.name) && !rmsPaths[folderName === "Finale" ? "finale" : folderName]) {
        rmsPaths[folderName === "Finale" ? "finale" : folderName] = entry.path_lower || entry.path_display;
      }
    }
  }

  if (!calcFiles.length) return null;
  const finaleFile = calcFiles.find((item) => item.name.includes("Finale"));
  const selected = finaleFile || calcFiles[calcFiles.length - 1];
  return { ...selected, rmsPaths, source: "dropbox" };
}

async function refreshStandings() {
  const season = activeSeason();
  elements.standingsStatus.textContent = "Chargement en cours...";
  elements.standingsContainer.innerHTML = "";
  
  console.log("=== Starting refreshStandings ===");
  
  try {
    let fileInfo = await findLatestCalculFile();
    
    if (!fileInfo) {
      const hasLinked = season.sourceMode === SOURCE_MODES.dropbox
        ? hasDropboxToken()
        : linkedDirectoryHandles.has(season.id);
      console.log("No file found. Linked:", hasLinked);
      
      if (!hasLinked) {
        console.log("Attempting to auto-link source...");
        elements.standingsStatus.textContent = season.sourceMode === SOURCE_MODES.dropbox
          ? "Connexion Dropbox en cours..."
          : "Liaison du dossier LGS en cours...";
        
        const linked = await linkSeasonFolder();
        if (!linked) {
          elements.standingsStatus.innerHTML = season.sourceMode === SOURCE_MODES.dropbox
            ? "<strong>Aucun fichier trouve.</strong><br>Connectez Dropbox puis relancez le rafraichissement."
            : `<strong>Aucun fichier trouve.</strong><br>La liaison au dossier LGS a ete perdue (rechargement de page?). Cliquez sur "Lier le dossier LGS" pour reconnecter, puis revenez ici.`;
          return;
        }
        
        // Try again after linking
        fileInfo = await findLatestCalculFile();
      }
      
      if (!fileInfo) {
        elements.standingsStatus.textContent = "Aucun fichier Calcul La Grande Semaine trouve. Verifiez que le dossier LGS contient des fichiers *HOMME_OU_DAME*.xlsm. Consultez la console (F12) pour les details.";
        return;
      }
    }
    
    console.log("Reading file:", fileInfo.name, "from folder:", fileInfo.folder);
    const file = season.sourceMode === SOURCE_MODES.dropbox
      ? await dropboxDownload(fileInfo.path)
      : await fileInfo.handle.getFile();
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    
    console.log("Workbook sheets:", workbook.SheetNames);
    
    const standings = {};
    const categorySheets = { "Resultat LGS (HOMME)": "HOMME", "Resultat LGS (DAME)": "DAME" };
    const isFinaleFile = fileInfo.name.toLowerCase().includes("finale");
    
    // Detect if finale has actually been played: check if AF column has any numeric value
    let finaleHasBeenPlayed = false;
    if (isFinaleFile) {
      const checkSheet = workbook.Sheets["Resultat LGS (HOMME)"] || workbook.Sheets["Resultat LGS (DAME)"];
      if (checkSheet) {
        const checkData = XLSX.utils.sheet_to_json(checkSheet, { header: "A", defval: "" });
        finaleHasBeenPlayed = checkData.slice(1).some(row => {
          const val = row.AF;
          return val !== "" && val !== null && !isNaN(parseFloat(val));
        });
      }
      console.log(`Finale file: finale played = ${finaleHasBeenPlayed}`);
    }
    
    let sheetsFound = 0;

    // Extract per-tour competition dates from the RMS export files (column B, row 1)
    // Format: "16.08.2026" → "16 août"  — one file per tour folder
    const tourDateMap = {};
    const rmsSources = season.sourceMode === SOURCE_MODES.dropbox
      ? (fileInfo.rmsPaths || {})
      : (fileInfo.rmsHandles || {});
    for (const [tourKey, rmsSource] of Object.entries(rmsSources)) {
      try {
        const rmsFile = season.sourceMode === SOURCE_MODES.dropbox
          ? await dropboxDownload(rmsSource)
          : await rmsSource.getFile();
        const rmsBuffer = await rmsFile.arrayBuffer();
        const rmsWb = XLSX.read(rmsBuffer, { type: "array" });
        const rmsSheet = rmsWb.Sheets[rmsWb.SheetNames[0]];
        if (rmsSheet) {
          const rmsData = XLSX.utils.sheet_to_json(rmsSheet, { header: "A", defval: "" });
          // Row 0 = headers, Row 1+ = data; column B holds the date string "DD.MM.YYYY"
          const dateRaw = String(rmsData[1]?.B || "").trim();
          const m = dateRaw.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
          if (m) {
            const jsDate = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
            tourDateMap[tourKey] = jsDate.toLocaleDateString("fr-FR", { weekday: "short", day: "2-digit", month: "short" });
            console.log(`📅 ${tourKey}: ${dateRaw} → ${tourDateMap[tourKey]}`);
          }
        }
      } catch (e) {
        console.log(`⚠️  Could not read RMS date for ${tourKey}:`, e.message);
      }
    }

    for (const [sheetName, category] of Object.entries(categorySheets)) {
      if (!workbook.SheetNames.includes(sheetName)) {
        console.log(`Sheet not found: ${sheetName}`);
        continue;
      }
      
      console.log(`Processing sheet: ${sheetName}, Finale file: ${isFinaleFile}, Finale played: ${finaleHasBeenPlayed}`);
      sheetsFound++;
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: "A", defval: "" });
      
      console.log(`Data rows in ${sheetName}:`, data.length);
      
      const bySeriesAndTotal = {};
      let validRecords = 0;
      let skippedRecords = 0;
      
      // Per-tour column mapping: [NET col, BRUT col, label]
      const TOUR_COLS = [
        ["F",  "H",  "T1"],
        ["J",  "L",  "T2"],
        ["N",  "P",  "T3"],
        ["R",  "T",  "T4"],
        ["V",  "X",  "T5"],
        ["Z",  "AB", "T6"],
      ];

      data.forEach((row, rowIndex) => {
        // Skip header rows (row 0 = group headers, row 1 = column labels, row 2 = empty)
        if (rowIndex <= 1) return;

        // Extract name from column B
        const name = String(row.B || "").trim();

        // Series is in column E
        const series = row.E || "";

        const bestNET    = row.AD || "";
        const bestBRUT   = row.AE || "";
        const finalNET   = row.AF || "";
        const finalBRUT  = row.AH || "";
        const totalNET   = row.AJ || "";
        const totalBRUT  = row.AK || "";

        // Skip placeholder/header rows
        if (name === "Nom - Prénom" || name === "Nom - prenom" || !name || !series) {
          skippedRecords++;
          return;
        }

        const seriesKey = String(series).toLowerCase();

        // Build per-tour scores array (only tours that have a numeric score)
        const tourScores = {};
        for (const [netCol, brutCol, label] of TOUR_COLS) {
          const n = parseFloat(row[netCol]);
          const b = parseFloat(row[brutCol]);
          if (!isNaN(n)) tourScores[label] = { NET: n, BRUT: isNaN(b) ? n : b };
        }

        const addRecord = (scoreType, totalRaw, finalRaw, bestRaw) => {
          let total = parseFloat(totalRaw);
          let totalScore = String(totalRaw).trim();
          
          // "En cours" = season in progress, use best-tour score as ranking proxy
          if (isNaN(total)) {
            const best = parseFloat(bestRaw);
            if (isNaN(best)) {
              if (name.toLowerCase().includes("salgado")) console.log(`  ❌ ${scoreType} SKIPPED (no score at all):`, { totalRaw, bestRaw });
              return; // no score at all, skip
            }
            total = best;
            totalScore = String(bestRaw).trim();
          }

          // When finale has been played, skip players without a finale score
          if (isFinaleFile && finaleHasBeenPlayed) {
            const finalVal = parseFloat(finalRaw);
            if (isNaN(finalVal)) {
              skippedRecords++;
              if (name.toLowerCase().includes("salgado")) console.log(`  ❌ ${scoreType} SKIPPED (Finale without final score):`, { finalRaw });
              return;
            }
          }

          if (!bySeriesAndTotal[seriesKey]) bySeriesAndTotal[seriesKey] = [];
          const record = {
            name,
            series: String(series).trim(),
            type: scoreType,
            bestScore: String(bestRaw).trim(),
            finalScore: String(finalRaw).trim(),
            total,
            totalScore,
            inProgress: String(totalRaw).trim() === "En cours",
            tourScores: Object.fromEntries(
              Object.entries(tourScores).map(([t, v]) => [t, v[scoreType]])
            )
          };
          bySeriesAndTotal[seriesKey].push(record);
          validRecords++;
          if (name.toLowerCase().includes("salgado")) console.log(`  ✅ ${scoreType} record added:`, record);
        };

        if (name.toLowerCase().includes("salgado")) {
          console.log(`🔍 SALGADO FOUND in ${sheetName}:`, { name, series: seriesKey, totalNET, totalBRUT, finalNET, finalBRUT, bestNET, bestBRUT, tourScores });
        }

        if (totalNET   !== "") addRecord("NET",  totalNET,  finalNET,  bestNET);
        if (totalBRUT  !== "") addRecord("BRUT", totalBRUT, finalBRUT, bestBRUT);
      });
      
      console.log(`${sheetName}: ${validRecords} valid, ${skippedRecords} skipped, series found: ${Object.keys(bySeriesAndTotal).sort().join(", ")}`);
      standings[category] = bySeriesAndTotal;
    }
    
    if (sheetsFound === 0) {
      elements.standingsStatus.textContent = `Erreur : Les feuilles "Resultat LGS (HOMME)" et "Resultat LGS (DAME)" n'ont pas ete trouvees dans ${fileInfo.name}. Verifiez le nom des onglets du classement.`;
      return;
    }
    
    // Store file handle for opening from standings
    currentStandingsFile = {
      handle: fileInfo.handle,
      name: fileInfo.name
    };
    
    renderStandings(standings, fileInfo.name, isFinaleFile, finaleHasBeenPlayed, tourDateMap);
    renderStatistics(standings, fileInfo.name);
  } catch (error) {
    elements.standingsStatus.textContent = "Erreur : Impossible de lire le fichier Excel. Verifiez qu'il n'est pas ouvert.";
    console.error(error);
  }
}

async function openStandingsFile() {
  if (!currentStandingsFile) {
   alert("Aucun fichier charge. Cliquez sur 'Rafraichir classement' d'abord.");
   return;
  }
  
  try {
   const file = await currentStandingsFile.handle.getFile();
   const blob = new Blob([await file.arrayBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
   const url = URL.createObjectURL(blob);
   const a = document.createElement("a");
   a.href = url;
   a.download = currentStandingsFile.name;
   document.body.appendChild(a);
   a.click();
   document.body.removeChild(a);
   URL.revokeObjectURL(url);
  } catch (error) {
   console.error("Erreur lors de l'ouverture du fichier:", error);
   alert("Impossible d'ouvrir le fichier. Verifiez que le dossier LGS est toujours lie.");
  }
}

function printStandingsPanel(panel, tabLabel, fileName) {
  // Build a print container with title header + cloned panel content
  const season = activeSeason();
  const year = season ? season.year : "";
  const dateStr = new Date().toLocaleDateString("fr-FR", { day: "2-digit", month: "long", year: "numeric" });

  let printArea = document.getElementById("lgs-print-area");
  if (!printArea) {
    printArea = document.createElement("div");
    printArea.id = "lgs-print-area";
    document.body.appendChild(printArea);
  }

  // Header
  printArea.innerHTML = `
    <div class="print-header">
      <div class="print-logo">LGS</div>
      <div class="print-title-block">
        <div class="print-eyebrow">ASGLM — LA GRANDE SEMAINE ${year}</div>
        <div class="print-tab-name">${tabLabel}</div>
        <div class="print-source">${fileName} · Imprimé le ${dateStr}</div>
      </div>
    </div>
    <hr class="print-divider">
  `;

  // Clone the active panel content (strips event listeners, keeps structure)
  const clone = panel.cloneNode(true);
  // Remove the "open file" buttons from the clone — they don't work in print
  clone.querySelectorAll(".open-file-btn").forEach(b => b.remove());
  clone.querySelectorAll(".standings-tab").forEach(b => b.remove());
  clone.querySelectorAll(".tour-subtab").forEach(b => b.remove());
  clone.querySelectorAll(".tour-subtabs").forEach(b => b.remove());
  // Unhide all hidden elements so they appear in print (especially hidden sub-panels)
  clone.querySelectorAll("[hidden]").forEach(el => {
    el.removeAttribute("hidden");
  });
  printArea.appendChild(clone);

  window.print();
}

function shareStandingsPanelToWhatsapp(panel, tabLabel, fileName) {
  // Generate a PDF like printStandingsPanel does
  const season = activeSeason();
  const year = season ? season.year : "";
  const dateStr = new Date().toLocaleDateString("fr-FR", { day: "2-digit", month: "long", year: "numeric" });

  // Build HTML content with title header + cloned panel content
  let printArea = document.getElementById("lgs-print-area");
  if (!printArea) {
    printArea = document.createElement("div");
    printArea.id = "lgs-print-area";
    document.body.appendChild(printArea);
  }

  // Create print header
  const headerHTML = `
    <div class="print-header">
      <div class="print-logo">LGS</div>
      <div class="print-title-block">
        <div class="print-eyebrow">ASGLM — LA GRANDE SEMAINE ${year}</div>
        <div class="print-tab-name">${tabLabel}</div>
        <div class="print-source">${fileName} · Partagé le ${dateStr}</div>
      </div>
    </div>
    <hr class="print-divider">
  `;

  // Clone panel and remove buttons
  const clone = panel.cloneNode(true);
  clone.querySelectorAll(".open-file-btn").forEach(b => b.remove());
  clone.querySelectorAll(".standings-tab").forEach(b => b.remove());
  clone.querySelectorAll(".tour-subtab").forEach(b => b.remove());
  clone.querySelectorAll(".tour-subtabs").forEach(b => b.remove());
  
  // Unhide all hidden elements so they appear in PDF (especially hidden sub-panels)
  clone.querySelectorAll("[hidden]").forEach(el => {
    el.removeAttribute("hidden");
  });

  // Create container for PDF content
  const pdfContent = document.createElement("div");
  pdfContent.innerHTML = headerHTML;
  pdfContent.appendChild(clone);

  // Generate PDF using html2pdf
  const options = {
    margin: 10,
    filename: `LGS-${year}-${tabLabel.replace(/\s+/g, "-")}.pdf`,
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { orientation: "portrait", unit: "mm", format: "a4" }
  };

  // Generate PDF as blob
  html2pdf()
    .set(options)
    .from(pdfContent)
    .outputPdf("blob")
    .then(blob => {
      // Try using Web Share API (works on mobile for WhatsApp)
      if (navigator.share) {
        navigator.share({
          title: `LGS ${year} - ${tabLabel}`,
          text: `Résultats du ${tabLabel}`,
          files: [
            new File([blob], `LGS-${year}-${tabLabel.replace(/\s+/g, "-")}.pdf`, { type: "application/pdf" })
          ]
        }).catch(err => {
          // User cancelled or sharing failed
          console.log("Partage annulé ou échoué:", err);
          fallbackShare(blob, year, tabLabel);
        });
      } else {
        // Fallback: For desktop or browsers without Web Share API
        fallbackShare(blob, year, tabLabel);
      }
    })
    .catch(error => {
      console.error("Erreur lors de la génération du PDF:", error);
      alert("Impossible de générer le PDF pour le partage.");
    });

  function fallbackShare(blob, year, tabLabel) {
    // Create a download link and suggest manual WhatsApp sharing
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `LGS-${year}-${tabLabel.replace(/\s+/g, "-")}.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    alert(
      `Le PDF a été téléchargé.\n\n` +
      `Pour le partager sur WhatsApp:\n` +
      `1. Ouvrez WhatsApp\n` +
      `2. Sélectionnez un contact ou groupe\n` +
      `3. Cliquez sur le + (Joindre un fichier)\n` +
      `4. Sélectionnez le fichier PDF téléchargé`
    );
  }
}

function computeStatistics(standings) {
  const stats = {
    totalPlayers: 0,
    totalCards: 0,
    womenPlayers: 0,
    menPlayers: 0,
    playersByCategory: {},
    cardsByCategory: {},
    playersBySeries: {},
    cardsBySeries: {},
    cardsPerTour: {},
    uniquePlayerNames: new Set(),
    uniqueWomenNames: new Set(),
    uniqueMenNames: new Set()
  };

  const TOUR_NAMES = ["T1", "T2", "T3", "T4", "T5", "T6"];

  // Initialize tour counters
  TOUR_NAMES.forEach(tour => {
    stats.cardsPerTour[tour] = 0;
  });

  // Process each category (HOMME/DAME)
  for (const [category, seriesData] of Object.entries(standings)) {
    stats.playersByCategory[category] = 0;
    stats.cardsByCategory[category] = 0;

    if (!seriesData || Object.keys(seriesData).length === 0) continue;

    // Process each series
    for (const [seriesKey, players] of Object.entries(seriesData)) {
      if (!stats.playersBySeries[seriesKey]) {
        stats.playersBySeries[seriesKey] = { men: 0, women: 0, total: 0 };
        stats.cardsBySeries[seriesKey] = 0;
      }

      if (!Array.isArray(players)) continue;

      // Process each player
      for (const player of players) {
        if (!player.name || !player.name.trim()) continue;

        // Count unique players by name
        stats.uniquePlayerNames.add(player.name);

        // Count by category
        stats.playersByCategory[category]++;
        stats.totalPlayers++;

        // Count by gender (unique names only)
        if (category === "DAME") {
          stats.uniqueWomenNames.add(player.name);
          stats.playersBySeries[seriesKey].women++;
        } else if (category === "HOMME") {
          stats.uniqueMenNames.add(player.name);
          stats.playersBySeries[seriesKey].men++;
        }

        stats.playersBySeries[seriesKey].total++;

        // Count cards played (each player record = one card)
        stats.cardsByCategory[category]++;
        stats.totalCards++;
        stats.cardsBySeries[seriesKey]++;

        // Count cards per tour
        if (player.tourScores && typeof player.tourScores === "object") {
          for (const tour of TOUR_NAMES) {
            if (player.tourScores[tour] !== undefined && player.tourScores[tour] !== null && player.tourScores[tour] !== "") {
              stats.cardsPerTour[tour]++;
            }
          }
        }
      }
    }
  }

  // Count unique players (one person might have multiple cards for NET/BRUT)
  stats.uniquePlayers = stats.uniquePlayerNames.size;
  stats.womenPlayers = stats.uniqueWomenNames.size;
  stats.menPlayers = stats.uniqueMenNames.size;

  return stats;
}

function renderStatistics(standings, fileName) {
  const stats = computeStatistics(standings);
  const container = document.getElementById("statistics-container");
  const statusEl = document.getElementById("statistics-status");

  if (!container) return;

  container.innerHTML = "";
  statusEl.textContent = `Données de : ${fileName}`;

  if (stats.totalCards === 0) {
    container.innerHTML = "<p style='color: #888; font-style: italic; padding: 1rem 0;'>Aucune donnée disponible.</p>";
    return;
  }

  // Create main stat cards
  const mainStatsHtml = `
    <div class="stat-card">
      <div class="stat-label">Joueurs différents</div>
      <div class="stat-value">${stats.uniquePlayers}</div>
      <div class="stat-description">Personnes inscrites à la compétition</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Cartes jouées</div>
      <div class="stat-value">${stats.totalCards}</div>
      <div class="stat-description">Total des cartes (NET + BRUT)</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Femmes</div>
      <div class="stat-value">${stats.womenPlayers}</div>
      <div class="stat-description">Joueuses inscrites</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Hommes</div>
      <div class="stat-value">${stats.menPlayers}</div>
      <div class="stat-description">Joueurs inscrits</div>
    </div>
  `;

  const mainStatsDiv = document.createElement("div");
  mainStatsDiv.className = "statistics-container";
  mainStatsDiv.innerHTML = mainStatsHtml;
  container.appendChild(mainStatsDiv);

  // Create detailed breakdown section
  const detailsDiv = document.createElement("div");
  detailsDiv.className = "statistics-grid";

  // Category breakdown
  const categoryHtml = `
    <div class="statistics-detail">
      <div class="statistics-detail-title" style="cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center;" data-toggle="category">
        <span>Par catégorie</span>
        <span class="statistics-toggle-icon">▼</span>
      </div>
      <div class="statistics-detail-content" data-content="category">
        ${
          Object.entries(stats.playersByCategory)
            .map(([cat, count]) => `
              <div class="statistics-detail-row">
                <div class="statistics-detail-label">${cat}</div>
                <div class="statistics-detail-value">${count} <span style="font-size: 0.8rem; color: #888;">cartes</span></div>
              </div>
            `).join("")
        }
      </div>
    </div>
  `;

  // Tour breakdown
  const tourHtml = `
    <div class="statistics-detail">
      <div class="statistics-detail-title" style="cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center;" data-toggle="tour">
        <span>Cartes par tour</span>
        <span class="statistics-toggle-icon">▼</span>
      </div>
      <div class="statistics-detail-content" data-content="tour">
        ${
          Object.entries(stats.cardsPerTour)
            .map(([tour, count]) => count > 0 ? `
              <div class="statistics-detail-row">
                <div class="statistics-detail-label">${tour}</div>
                <div class="statistics-detail-value">${count}</div>
              </div>
            ` : "")
            .join("")
        }
        ${Object.values(stats.cardsPerTour).every(c => c === 0) ? 
          '<div style="padding: 0.5rem 0; color: #888; font-size: 0.8rem;">Pas de données de tour</div>' : 
          ""}
      </div>
    </div>
  `;

  // Series breakdown
  const seriesHtml = `
    <div class="statistics-detail">
      <div class="statistics-detail-title" style="cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center;" data-toggle="series">
        <span>Répartition par série</span>
        <span class="statistics-toggle-icon">▼</span>
      </div>
      <div class="statistics-detail-content" data-content="series">
        ${
          Object.entries(stats.playersBySeries)
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([seriesKey, data]) => `
              <div>
                <div style="font-weight: 600; margin-top: 0.5rem; margin-bottom: 0.3rem; color: var(--green);">${seriesKey.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ')}</div>
                <div class="statistics-detail-row" style="border: none; padding: 0.2rem 0; font-size: 0.75rem;">
                  <span class="statistics-detail-label">Femmes</span>
                  <span class="statistics-detail-value">${data.women}</span>
                </div>
                <div class="statistics-detail-row" style="border: none; padding: 0.2rem 0; font-size: 0.75rem;">
                  <span class="statistics-detail-label">Hommes</span>
                  <span class="statistics-detail-value">${data.men}</span>
                </div>
                <div class="statistics-detail-row" style="border: none; padding: 0.2rem 0; font-size: 0.75rem;">
                  <span class="statistics-detail-label">Total</span>
                  <span class="statistics-detail-value">${data.total}</span>
                </div>
              </div>
            `).join("")
        }
      </div>
    </div>
  `;

  detailsDiv.innerHTML = categoryHtml + tourHtml + seriesHtml;
  container.appendChild(detailsDiv);
  
  // Add toggle event listeners for collapsible sections
  const toggleButtons = container.querySelectorAll('[data-toggle]');
  toggleButtons.forEach(button => {
    button.addEventListener('click', function() {
      const contentType = this.getAttribute('data-toggle');
      const content = container.querySelector(`[data-content="${contentType}"]`);
      const icon = this.querySelector('.statistics-toggle-icon');
      
      content.classList.toggle('statistics-detail-collapsed');
      icon.textContent = content.classList.contains('statistics-detail-collapsed') ? '▶' : '▼';
    });
  });
}


function renderStandings(standings, fileName, isFinaleFile = false, finaleHasBeenPlayed = false, tourDateMap = {}) {
  console.log("renderStandings called, standings keys:", Object.keys(standings));
  for (const [cat, seriesData] of Object.entries(standings)) {
    const seriesKeys = Object.keys(seriesData);
    console.log(`  ${cat}: ${seriesKeys.length} series, total players:`,
      seriesKeys.reduce((n, k) => n + seriesData[k].length, 0));
    if (seriesKeys.length > 0) {
      const sample = seriesData[seriesKeys[0]][0];
      console.log(`  sample player:`, sample?.name, "tourScores:", sample?.tourScores);
    }
  }

  const availableTours = ["T1","T2","T3","T4","T5","T6"];
  const toursWithData = availableTours.filter(t =>
    Object.values(standings).some(cat =>
      Object.values(cat).some(players =>
        players.some(p => p.tourScores && p.tourScores[t] !== undefined)
      )
    )
  );
  console.log("toursWithData:", toursWithData);
  // Helper: build display label with date suffix when available
  function withDate(id, baseLabel, includeDate = false) {
    if (!includeDate) return baseLabel;
    const d = tourDateMap[id];
    return d ? `${baseLabel} · ${d}` : baseLabel;
  }

  const tabs = [
    { id: "best", label: "Top 10" },
    // Always show all T1–T6 tabs + Finale (even without data yet)
    ...availableTours.map(t => ({ id: t, label: withDate(t, t, true) })),
    { id: "finale", label: withDate("finale", "Finale", true) },
    { id: "all",  label: "Tout" }
  ];

  // --- Tab bar + empty panels (append panels to DOM now) ---
  const tabBar = document.createElement("div");
  tabBar.className = "standings-tabs";
  elements.standingsContainer.appendChild(tabBar);

  // Track active tab label for the PDF title
  let activeTabLabel = tabs[0]?.label ?? "Top 10";

  const panels = {};
  tabs.forEach((tab, i) => {
    const panel = document.createElement("div");
    panel.className = "standings-tab-panel";
    panel.hidden = i !== 0;
    panels[tab.id] = panel;
    elements.standingsContainer.appendChild(panel);

    const btn = document.createElement("button");
    btn.className = "standings-tab" + (i === 0 ? " active" : "");
    btn.textContent = tab.label;
    btn.addEventListener("click", () => {
      tabBar.querySelectorAll(".standings-tab").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      Object.values(panels).forEach(p => { p.hidden = true; });
      panel.hidden = false;
      activeTabLabel = tab.label;
      rebuildSectionNav(panel);
    });
    tabBar.appendChild(btn);
  });

  // PDF export button (always visible, exports the currently active tab)
  const pdfBtn = document.createElement("button");
  pdfBtn.className = "standings-tab standings-tab-pdf";
  pdfBtn.textContent = "📄 PDF";
  pdfBtn.title = "Exporter l'onglet actif en PDF";
  pdfBtn.addEventListener("click", () => {
    const activePanel = Object.values(panels).find(p => !p.hidden);
    if (!activePanel) return;
    printStandingsPanel(activePanel, activeTabLabel, fileName);
  });
  tabBar.appendChild(pdfBtn);

  // WhatsApp share button (always visible, shares the currently active tab)
  const whatsappBtn = document.createElement("button");
  whatsappBtn.className = "standings-tab standings-tab-whatsapp";
  whatsappBtn.textContent = "📱 Partager";
  whatsappBtn.title = "Partager l'onglet actif sur WhatsApp";
  whatsappBtn.addEventListener("click", () => {
    const activePanel = Object.values(panels).find(p => !p.hidden);
    if (!activePanel) return;
    shareStandingsPanelToWhatsapp(activePanel, activeTabLabel, fileName);
  });
  tabBar.appendChild(whatsappBtn);

  // Section nav — one per panel, prepended at top before content is filled
  // rebuildSectionNav fills the nav inside the given panel
  function getOrCreateNav(panel) {
    let nav = panel.querySelector(":scope > .standings-section-nav");
    if (!nav) {
      nav = document.createElement("nav");
      nav.className = "standings-section-nav";
      panel.prepend(nav);
    }
    return nav;
  }

  function rebuildSectionNav(panel) {
    const sectionNav = getOrCreateNav(panel);
    sectionNav.innerHTML = "";

    function isVisible(el) {
      let cur = el.parentElement;
      while (cur && cur !== panel) {
        if (cur.hidden) return false;
        cur = cur.parentElement;
      }
      return true;
    }

    // Collect all visible cat-headings (HOMME / DAME)
    const catHeadings = [...panel.querySelectorAll(".cat-heading")].filter(isVisible);
    if (catHeadings.length === 0) return;

    catHeadings.forEach(catEl => {
      const catLabel = catEl.dataset.category || catEl.textContent.trim();

      // Outer sex group
      const catGroup = document.createElement("div");
      catGroup.className = "section-nav-cat";

      const catBtn = document.createElement("span");
      catBtn.className = "section-nav-cat-label";
      catBtn.textContent = catLabel;
      catGroup.appendChild(catBtn);

      // Collect all score-type-headers that belong to this category
      // (siblings after catEl, until the next catEl)
      const scoreHeaders = [];
      let node = catEl.nextElementSibling;
      while (node) {
        if (node.classList && node.classList.contains("cat-heading")) break;
        // Could be nested in sub-panels — search within
        const headers = node.classList && node.classList.contains("score-type-header")
          ? [node]
          : [...node.querySelectorAll(".score-type-header[id]")];
        headers.filter(isVisible).forEach(h => scoreHeaders.push(h));
        node = node.nextElementSibling;
      }

      const groupsWrap = document.createElement("div");
      groupsWrap.className = "section-nav-groups";

      scoreHeaders.forEach(h => {
        const row = document.createElement("div");
        row.className = "section-nav-group";

        // Check if this header is inside a tour subtab and extract tour info
        let tourLabel = null;
        let tourPanel = h.closest(".standings-tab-panel");
        if (tourPanel) {
          // Look for tour identifier in panel structure
          const tourMatch = tourPanel.id && tourPanel.id.match(/tour-(\w+)/);
          if (tourMatch) {
            tourLabel = tourMatch[1].toUpperCase();
          }
        }

        // Create non-clickable tour label if this is a tour-specific panel
        if (tourLabel) {
          const tourLabelSpan = document.createElement("span");
          tourLabelSpan.className = "section-nav-tour-label";
          tourLabelSpan.textContent = tourLabel;
          row.appendChild(tourLabelSpan);
        }

        // Create non-clickable label for NET/BRUT
        const headerLabel = document.createElement("span");
        headerLabel.className = "section-nav-header";
        if (/NET/i.test(h.textContent)) headerLabel.classList.add("section-nav-net");
        if (/BRUT/i.test(h.textContent)) headerLabel.classList.add("section-nav-brut");
        headerLabel.textContent = h.textContent;
        row.appendChild(headerLabel);

        // Series chips under this header
        let sib = h.nextElementSibling;
        while (sib) {
          if (sib.classList && sib.classList.contains("score-type-header")) break;
          if (sib.classList && sib.classList.contains("cat-heading")) break;
          const groups = sib.classList && sib.classList.contains("series-group")
            ? [sib]
            : [...sib.querySelectorAll(".series-group[id]")];
          groups.filter(isVisible).forEach(sg => {
            const title = sg.querySelector(".series-title");
            if (!title || !sg.id) return;
            const a = document.createElement("a");
            a.href = "#" + sg.id;
            a.className = "section-nav-link section-nav-series";
            a.textContent = title.textContent;
            a.addEventListener("click", e => {
              e.preventDefault();
              sg.scrollIntoView({ behavior: "smooth", block: "start" });
            });
            row.appendChild(a);
          });
          sib = sib.nextElementSibling;
        }

        groupsWrap.appendChild(row);
      });

      if (scoreHeaders.length > 0) {
        catGroup.appendChild(groupsWrap);
        sectionNav.appendChild(catGroup);
      }
    });
  }
  // Helper: build an inner sub-tab switcher inside a tour panel
  // subViews: [{ label, buildFn }] where buildFn(container) fills that sub-view
  function makeTourSubTabs(panel, subViews) {
    const bar = document.createElement("div");
    bar.className = "tour-subtabs";
    panel.appendChild(bar);

    const subPanels = subViews.map((sv, i) => {
      const sp = document.createElement("div");
      sp.hidden = i !== 0;
      panel.appendChild(sp);

      const btn = document.createElement("button");
      btn.className = "tour-subtab" + (i === 0 ? " active" : "");
      btn.textContent = sv.label;
      btn.addEventListener("click", () => {
        bar.querySelectorAll(".tour-subtab").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        subPanels.forEach(p => { p.hidden = true; });
        sp.hidden = false;
        rebuildSectionNav(sp);
      });
      bar.appendChild(btn);

      // Fill the sub-panel now
      sv.buildFn(sp);
      return sp;
    });

    // Initialize nav for first sub-panel
    if (subPanels[0]) rebuildSectionNav(subPanels[0]);
  }


  function sortedWithTies(players, limit) {
    const sorted = [...players].sort((a, b) => {
      const d = a.total - b.total;
      return d !== 0 ? d : a.name.localeCompare(b.name);
    });
    if (sorted.length <= limit) return sorted;
    const cutScore = sorted[limit - 1].total;
    // keep top N, then extend for any tied at the cut
    const base = sorted.slice(0, limit);
    const extra = sorted.slice(limit).filter(p => p.total === cutScore);
    return base.concat(extra);
  }

  function trueRank(player, allSorted) {
    return allSorted.filter(p => p.total < player.total).length + 1;
  }

  function makeOpenFileBtn() {
    const btn = document.createElement("button");
    btn.className = "open-file-btn";
    btn.textContent = "📄";
    btn.title = "Ouvrir le fichier Excel";
    btn.addEventListener("click", () => openStandingsFile());
    return btn;
  }

  function makePlayerRow(player, rank, cols, bestScores = {}) {
    const row = document.createElement("div");
    row.className = "player-row";
    row.style.gridTemplateColumns = `2rem 1.5fr ${cols.map(() => "1fr").join(" ")}`;

    const rankCell = document.createElement("div");
    rankCell.className = "rank";
    
    // Add medal emoji for top 3 with rank number
    let rankDisplay = String(rank);
    let medal = "";
    if (rank === 1) {
      medal = "🥇";
      rankCell.classList.add("rank-gold");
    } else if (rank === 2) {
      medal = "🥈";
      rankCell.classList.add("rank-silver");
    } else if (rank === 3) {
      medal = "🥉";
      rankCell.classList.add("rank-bronze");
    }
    
    rankCell.textContent = rankDisplay + (medal ? " " + medal : "");
    row.appendChild(rankCell);

    const nameCell = document.createElement("div");
    nameCell.className = "name";
    nameCell.textContent = player.name;
    row.appendChild(nameCell);

    for (let colIndex = 0; colIndex < cols.length; colIndex++) {
      const col = cols[colIndex];
      const cell = document.createElement("div");
      cell.className = "score-cell";
      if (col.bold) cell.style.fontWeight = "700";
      const val = col.value(player);
      
      // Check if this score is the best in its column
      const isBest = bestScores[colIndex] !== undefined && val === bestScores[colIndex];
      const star = isBest ? "⭐ " : "";
      
      // Only show the value, not the label (label goes in header row)
      cell.innerHTML = `<div class="score-value">${star}${val !== undefined && val !== "" ? val : "—"}</div>`;
      row.appendChild(cell);
    }
    return row;
  }

  function makeColumnHeader(cols) {
    const header = document.createElement("div");
    header.className = "column-header";
    header.style.gridTemplateColumns = `2rem 1.5fr ${cols.map(() => "1fr").join(" ")}`;
    
    // Rank column header (empty)
    const rankHeader = document.createElement("div");
    rankHeader.className = "header-cell";
    header.appendChild(rankHeader);
    
    // Name column header (empty)
    const nameHeader = document.createElement("div");
    nameHeader.className = "header-cell";
    header.appendChild(nameHeader);
    
    // Score column headers
    for (const col of cols) {
      const cell = document.createElement("div");
      cell.className = "header-cell";
      cell.textContent = col.label;
      header.appendChild(cell);
    }
    return header;
  }

  function compactSeriesLabel(seriesName) {
    const raw = String(seriesName || "").trim();
    const numberMatch = raw.match(/(\d+)/);
    if (numberMatch) return `Serie${numberMatch[1]}`;
    return raw;
  }

  function makeSeriesGroup(seriesName, players, cols, sectionLabel, combinedRankingGroup = null) {
    const group = document.createElement("div");
    group.className = "series-group";
    // Unique anchor id: section label + series name
    const anchorBase = (sectionLabel ? sectionLabel + "-" + seriesName : seriesName)
      .toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
    group.id = "sg-" + anchorBase;

    const titleRow = document.createElement("div");
    titleRow.style.cssText = "display:flex;align-items:center;gap:.5rem;margin-bottom:.5rem";
    const title = document.createElement("div");
    title.className = "series-title";
    title.textContent = compactSeriesLabel(seriesName);
    titleRow.appendChild(title);
    titleRow.appendChild(makeOpenFileBtn());
    group.appendChild(titleRow);

    // Add column header row
    group.appendChild(makeColumnHeader(cols));

    // Calculate best score for each column
    const bestScores = {};
    cols.forEach((col, colIndex) => {
      const scores = players.map(p => {
        const val = col.value(p);
        return typeof val === 'number' ? val : null;
      }).filter(v => v !== null);
      if (scores.length > 0) {
        // For golf, lower is better (unless it's a ranking/position column)
        bestScores[colIndex] = Math.min(...scores);
      }
    });

    // Use combinedRankingGroup for ranking if provided (for series 3-4), otherwise use current series players
    const rankingGroup = combinedRankingGroup || players;

    players.forEach(player => {
      group.appendChild(makePlayerRow(player, trueRank(player, rankingGroup), cols, bestScores));
    });

    // Back-to-top link
    const topLink = document.createElement("a");
    topLink.className = "back-to-top";
    topLink.href = "#";
    topLink.textContent = "↑ Haut";
    topLink.addEventListener("click", e => {
      e.preventDefault();
      // Scroll to show the full header of current tab panel with all context
      const nearestPanel = group.closest(".standings-tab-panel");
      if (nearestPanel) {
        const panelRect = nearestPanel.getBoundingClientRect();
        // Scroll with -200px offset to show "CLASSEMENT EN COURS", "Résultats du jour" header, description, and tabs
        window.scrollTo({
          top: window.scrollY + panelRect.top - 200,
          behavior: "smooth"
        });
      }
    });
    group.appendChild(topLink);

    return group;
  }

  // Helper: check if two series should have shared ranking (series 3-4 per tournament rules)
  function shouldGroupSeries(s1, s2) {
    const n1 = String(s1).toLowerCase().trim();
    const n2 = String(s2).toLowerCase().trim();
    // Extract numeric part from series names (e.g., "serie 3" -> 3)
    const num1 = parseInt(n1.match(/\d+/)?.[0]) || null;
    const num2 = parseInt(n2.match(/\d+/)?.[0]) || null;
    // Series 3 and 4 share common ranking
    return (num1 === 3 && num2 === 4) || (num1 === 4 && num2 === 3);
  }

  function getSeriesNumber(seriesName) {
    const match = String(seriesName || "").toLowerCase().match(/\d+/);
    return match ? parseInt(match[0]) : null;
  }

  function appendTypeSection(panel, label, seriesNames, seriesData, scoreType, getPlayers, cols) {
    let added = false;
    const processedSeries = new Set();

    for (const sn of seriesNames) {
      if (processedSeries.has(sn)) continue; // Already processed as part of a group

      // Check if this is series 3 or 4 and has a pair
      const seriesNum = getSeriesNumber(sn);
      let groupedSeries = [sn];
      if ((seriesNum === 3 || seriesNum === 4)) {
        // Look for the paired series (3-4)
        const otherNum = seriesNum === 3 ? 4 : 3;
        const otherSeries = seriesNames.find(s => getSeriesNumber(s) === otherNum);
        if (otherSeries) {
          // Combine series 3 and 4 for ranking
          groupedSeries = [sn, otherSeries].sort();
          processedSeries.add(otherSeries);
        }
      }

      // Get all players from grouped series
      const allGroupPlayers = [];
      for (const seriesKey of groupedSeries) {
        const arr = seriesData[seriesKey];
        if (!arr) { console.warn("appendTypeSection: no data for series", seriesKey); continue; }
        const players = getPlayers(arr, scoreType);
        if (players && players.length > 0) {
          allGroupPlayers.push(...players.map(p => ({ ...p, _seriesKey: seriesKey })));
        }
      }

      if (allGroupPlayers.length === 0) continue;

      if (!added) {
        const h = document.createElement("div");
        h.className = "score-type-header";
        // Assign a stable anchor id from the label
        h.id = "section-" + label.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
        h.textContent = label;
        panel.appendChild(h);
        added = true;
      }

      // Sort combined group for shared ranking
      const sortedGroup = [...allGroupPlayers].sort((a, b) => {
        const d = a.total - b.total;
        return d !== 0 ? d : a.name.localeCompare(b.name);
      });

      // Render each series separately but with shared ranking
      for (const seriesKey of groupedSeries) {
        const seriesPlayers = sortedGroup.filter(p => p._seriesKey === seriesKey)
          .map(p => {
            // Remove the temporary _seriesKey property
            const { _seriesKey, ...clean } = p;
            return clean;
          });

        if (seriesPlayers.length === 0) continue;

        console.log(`📊 Serie ${seriesKey} ${scoreType} (${label}):`, {
          totalPlayers: seriesPlayers.length,
          sharedRanking: groupedSeries.length > 1,
          top: seriesPlayers.map(p => ({ name: p.name, total: p.total }))
        });

        panel.appendChild(makeSeriesGroup(seriesKey, seriesPlayers, cols, label, sortedGroup));
      }

      processedSeries.add(sn);
    }
    if (!added) console.log(`appendTypeSection: nothing rendered for "${label}", scoreType=${scoreType}, seriesNames=`, seriesNames);
  }

  // --- Fill panels ---
  const categories = ["HOMME", "DAME"];
  const scoreTypes = ["NET", "BRUT"];

  for (const category of categories) {
    if (!standings[category] || Object.keys(standings[category]).length === 0) continue;
    const seriesData = standings[category];
    const seriesNames = Object.keys(seriesData).sort();

    const catHeading = () => {
      const h = document.createElement("h3");
      h.className = "cat-heading";
      h.dataset.category = category;
      h.textContent = category;
      return h;
    };

    // ── Classement tab (top 10, best + finale + total) ──
    panels["best"].appendChild(catHeading());
    for (const scoreType of scoreTypes) {
      const totalCols = [
        { label: "Meilleur tour", value: p => p.bestScore },
        ...(isFinaleFile && finaleHasBeenPlayed ? [{ label: "Finale", value: p => p.finalScore }] : []),
        { label: "Total LGS", value: p => p.inProgress ? `${p.totalScore} ⏳` : p.totalScore, bold: true }
      ];
      appendTypeSection(
        panels["best"], `${scoreType}`, seriesNames, seriesData, scoreType,
        (arr, t) => sortedWithTies(arr.filter(p => p.type === t), 10),
        totalCols
      );
    }

    // ── Per-tour tabs (sub-tabs: Top 10 / Tous) ──
    for (const tourId of availableTours) {
      if (!panels[tourId]) continue;
      const hasData = toursWithData.includes(tourId);
      const tourLabel = withDate(tourId, tourId);

      if (!hasData) {
        // Empty tour — show placeholder (only once, not per category)
        if (category === categories[0]) {
          const msg = document.createElement("p");
          msg.style.cssText = "color:#888;font-style:italic;padding:1rem 0";
          msg.textContent = "Pas encore de données pour ce tour.";
          panels[tourId].appendChild(msg);
        }
        continue;
      }

      panels[tourId].appendChild(catHeading());
      const tourCols = [
        { label: tourLabel, value: p => p.total, bold: true },
        { label: "Total LGS", value: p => p.inProgress ? `${p.totalScore} ⏳` : p.totalScore, bold: true }
      ];

      const getTourPlayers = (arr, t, limit) => {
        const raw = arr
          .filter(p => p.type === t && p.tourScores && p.tourScores[tourId] !== undefined)
          .map(p => ({ ...p, total: p.tourScores[tourId] }));
        return limit ? sortedWithTies(raw, limit) : [...raw].sort((a, b) => {
          const d = a.total - b.total; return d !== 0 ? d : a.name.localeCompare(b.name);
        });
      };

      makeTourSubTabs(panels[tourId], [
        {
          label: "Top 10",
          buildFn: sp => {
            for (const scoreType of scoreTypes) {
              appendTypeSection(
               sp, `${scoreType}`, seriesNames, seriesData, scoreType,
                (arr, t) => getTourPlayers(arr, t, 10), tourCols
              );
            }
          }
        },
        {
          label: "Tous",
          buildFn: sp => {
            for (const scoreType of scoreTypes) {
              appendTypeSection(
               sp, `${scoreType}`, seriesNames, seriesData, scoreType,
                (arr, t) => getTourPlayers(arr, t, 0), tourCols
              );
            }
          }
        }
      ]);
    }

    // ── Finale tab ──
    if (panels["finale"]) {
      if (!isFinaleFile || !finaleHasBeenPlayed) {
        // Show placeholder only once (first category)
        if (category === categories[0]) {
          const msg = document.createElement("p");
          msg.style.cssText = "color:#888;font-style:italic;padding:1rem 0";
          msg.textContent = isFinaleFile
            ? "La finale n'a pas encore été jouée."
            : "Le fichier de la finale n'est pas encore disponible.";
          panels["finale"].appendChild(msg);
        }
      } else {
      panels["finale"].appendChild(catHeading());
      const finaleLabel = withDate("finale", "Finale");
      const finalCols = [{ label: finaleLabel, value: p => p.finalScore, bold: true }];

      const getFinalePlayers = (arr, t, limit) => {
        const raw = arr
          .filter(p => p.type === t && p.finalScore && !isNaN(parseFloat(p.finalScore)))
          .map(p => ({ ...p, total: parseFloat(p.finalScore) }));
        return limit ? sortedWithTies(raw, limit) : [...raw].sort((a, b) => {
          const d = a.total - b.total; return d !== 0 ? d : a.name.localeCompare(b.name);
        });
      };

      makeTourSubTabs(panels["finale"], [
        {
          label: "Top 10",
          buildFn: sp => {
            for (const scoreType of scoreTypes) {
              appendTypeSection(
               sp, `${scoreType}`, seriesNames, seriesData, scoreType,
                (arr, t) => getFinalePlayers(arr, t, 10), finalCols
              );
            }
          }
        },
        {
          label: "Tous",
          buildFn: sp => {
            for (const scoreType of scoreTypes) {
              appendTypeSection(
               sp, `${scoreType}`, seriesNames, seriesData, scoreType,
                (arr, t) => getFinalePlayers(arr, t, 0), finalCols
              );
            }
          }
        }
      ]);
    } // end else (finale played)
    } // end if panels["finale"]

    // ── Tout tab (all players, no limit) ──
    panels["all"].appendChild(catHeading());
    for (const scoreType of scoreTypes) {
      const totalCols = [
        { label: "Meilleur tour", value: p => p.bestScore },
        ...(isFinaleFile && finaleHasBeenPlayed ? [{ label: "Finale", value: p => p.finalScore }] : []),
        { label: "Total LGS", value: p => p.inProgress ? `${p.totalScore} ⏳` : p.totalScore, bold: true }
      ];
      appendTypeSection(
        panels["all"], `${scoreType}`, seriesNames, seriesData, scoreType,
        (arr, t) => [...arr.filter(p => p.type === t)].sort((a, b) => {
          const d = a.total - b.total;
          return d !== 0 ? d : a.name.localeCompare(b.name);
        }),
        totalCols
      );
    }
  }

  // Build nav for all simple panels (best, all, empty tour panels — no sub-tabs)
  // Tour panels with sub-tabs handle their own nav inside makeTourSubTabs
  ["best", "all", ...availableTours, "finale"].forEach(id => {
    const p = panels[id];
    if (!p) return;
    // Only rebuild if panel has no tour-subtabs (those handle it themselves)
    if (!p.querySelector(".tour-subtabs")) rebuildSectionNav(p);
  });
}

document.querySelector("#new-season-button").addEventListener("click", () => {
  const nextYear = Math.max(...state.seasons.map((season) => season.year)) + 1;
  document.querySelector("#season-year").value = nextYear;
  document.querySelector("#season-directory").value = `..\\ASGLM ${nextYear}\\LGS`;
  elements.dialog.showModal();
});
document.querySelector("#cancel-dialog-button").addEventListener("click", () => elements.dialog.close());
elements.form.addEventListener("submit", createSeason);
elements.seasonSelect.addEventListener("change", () => { state.activeId = elements.seasonSelect.value; render(); });
elements.sourceLocalButton.addEventListener("click", () => setSourceMode(SOURCE_MODES.local));
elements.sourceDropboxButton.addEventListener("click", () => setSourceMode(SOURCE_MODES.dropbox));
elements.notes.addEventListener("change", () => { activeSeason().notes = elements.notes.value; saveState(); });
document.querySelector("#export-button").addEventListener("click", exportSeason);
elements.linkFolderButton.addEventListener("click", linkSeasonFolder);
elements.importInput.addEventListener("change", importSeason);
elements.deleteSeasonButton.addEventListener("click", () => {
  const season = activeSeason();
  if (state.seasons.length === 1) return;
  elements.deleteYearLabel.textContent = `Saisissez ${season.year} pour confirmer`;
  elements.deleteYearInput.value = "";
  elements.confirmDeleteButton.disabled = true;
  elements.deleteDialog.showModal();
});
elements.deleteYearInput.addEventListener("input", () => {
  elements.confirmDeleteButton.disabled = Number(elements.deleteYearInput.value) !== activeSeason().year;
});
document.querySelector("#cancel-delete-button").addEventListener("click", () => elements.deleteDialog.close());
elements.deleteForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const season = activeSeason();
  if (Number(elements.deleteYearInput.value) !== season.year) return;
  state.seasons = state.seasons.filter((item) => item.id !== season.id);
  state.activeId = state.seasons[0].id;
  elements.deleteDialog.close();
  render();
});
elements.refreshStandingsButton.addEventListener("click", refreshStandings);

render();
restoreSectionStates();
