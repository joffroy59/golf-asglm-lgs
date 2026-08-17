const STORAGE_KEY = "lgs-season-manager-v1";
const HISTORICAL_YEARS = [2023, 2024, 2025];
const TOUR_NAMES = ["Tour 1", "Tour 2", "Tour 3", "Tour 4", "Tour 5", "Tour 6", "Tour 7", "Finale"];
const STATUS_LABELS = {
  planned: "A preparer",
  ready: "Export pret",
  imported: "Import realise",
  validated: "Valide"
};

const elements = {
  seasonSelect: document.querySelector("#season-select"),
  seasonTitle: document.querySelector("#season-title"),
  seasonPath: document.querySelector("#season-path"),
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
  deleteSeasonButton: document.querySelector("#delete-season-button")
};

function makeSeason(year, directory) {
  return {
    id: crypto.randomUUID(),
    year: Number(year),
    directory,
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

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
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
  elements.seasonPath.textContent = season.directory;
  elements.scanResult.textContent = season.catalogMessage
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
    status.value = tour.status;
    file.value = tour.file;
    note.value = tour.note;
    sourceSummary.textContent = sourceLabel(tour.sourceFiles || []);
    openButton.disabled = !canOpenTourFile(season, tour);
    openButton.title = openButton.disabled
      ? "Liez le dossier LGS pour ouvrir ce fichier."
      : `Ouvrir ${tour.file}`;
    status.addEventListener("change", () => updateTour(tour.number, "status", status.value));
    file.addEventListener("change", () => {
      linkedFileHandles.delete(fileHandleKey(season.id, tour.number));
      updateTour(tour.number, "file", file.value.trim());
    });
    note.addEventListener("change", () => updateTour(tour.number, "note", note.value.trim()));
    openButton.addEventListener("click", () => openRmsFile(season.id, tour.number));
    return card;
  }));
}

function fileHandleKey(seasonId, tourNumber) {
  return `${seasonId}:${tourNumber}`;
}

function canOpenTourFile(season, tour) {
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
    const url = handle
      ? URL.createObjectURL(await handle.getFile())
      : knownRmsHref(season, tour);
    if (!url) return;
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.target = "_blank";
    anchor.rel = "noopener";
    anchor.click();
    if (handle) setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (_) {
    alert("Le fichier RMS ne peut plus etre ouvert. Reliez le dossier LGS puis reessayez.");
  }
}

function sourceLabel(files) {
  if (!files.length) return "Aucune donnee locale liee";
  if (files.length === 1) return `Donnee liee : ${files[0]}`;
  return `${files.length} fichiers Excel lies, dont ${files[0]}`;
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
  if (!window.showDirectoryPicker) {
    alert("Utilisez Microsoft Edge ou Google Chrome pour lier un dossier LGS.");
    return;
  }
  try {
    const root = await window.showDirectoryPicker({ mode: "read" });
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
      alert("Selectionnez le dossier LGS qui contient T1 a T7 et Finale.");
      return;
    }

    const season = activeSeason();
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
    season.catalogMessage = "";
    render();
    elements.scanResult.textContent = `${detectedCount} fichiers Excel detectes dans ${root.name}.`;
  } catch (error) {
    if (error.name !== "AbortError") alert("La lecture du dossier LGS a echoue.");
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

document.querySelector("#new-season-button").addEventListener("click", () => {
  const nextYear = Math.max(...state.seasons.map((season) => season.year)) + 1;
  document.querySelector("#season-year").value = nextYear;
  document.querySelector("#season-directory").value = `..\\ASGLM ${nextYear}\\LGS`;
  elements.dialog.showModal();
});
document.querySelector("#cancel-dialog-button").addEventListener("click", () => elements.dialog.close());
elements.form.addEventListener("submit", createSeason);
elements.seasonSelect.addEventListener("change", () => { state.activeId = elements.seasonSelect.value; render(); });
elements.notes.addEventListener("change", () => { activeSeason().notes = elements.notes.value; saveState(); });
document.querySelector("#export-button").addEventListener("click", exportSeason);
document.querySelector("#link-folder-button").addEventListener("click", linkSeasonFolder);
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

render();
