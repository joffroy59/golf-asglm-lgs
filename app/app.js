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
  deleteSeasonButton: document.querySelector("#delete-season-button"),
  standingsContainer: document.querySelector("#standings-container"),
  standingsStatus: document.querySelector("#standings-status"),
  refreshStandingsButton: document.querySelector("#refresh-standings-button")
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
const linkedDirectoryHandles = new Map();


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
    const uploadButton = card.querySelector(".upload-xls-button");
    status.value = tour.status;
    file.value = tour.file;
    note.value = tour.note;
    sourceSummary.textContent = sourceLabel(tour.sourceFiles || []);
    openButton.disabled = !canOpenTourFile(season, tour);
    openButton.title = openButton.disabled
      ? "Liez le dossier LGS pour ouvrir ce fichier."
      : `Ouvrir ${tour.file}`;
    uploadButton.disabled = !window.showOpenFilePicker || !window.showDirectoryPicker;
    uploadButton.title = uploadButton.disabled
      ? "Utilisez Microsoft Edge ou Google Chrome pour ajouter un fichier."
      : `Ajouter un fichier dans ${tour.name}`;
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

async function addResultFile(seasonId, tourNumber) {
  let root = linkedDirectoryHandles.get(seasonId);
  const season = state.seasons.find((item) => item.id === seasonId);
  const tour = season?.tours.find((item) => item.number === tourNumber);
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
      alert("Selectionnez le dossier LGS qui contient T1 a T7 et Finale.");
      return false;
    }

    const season = activeSeason();
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
  if (!linkedDirectoryHandles.has(season.id)) {
    console.log("❌ No linked directory for season:", season.id);
    return null;
  }
  
  const root = linkedDirectoryHandles.get(season.id);
  const calcFiles = [];
  
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
  const tourFolders = ["Finale", "T7", "T6", "T5", "T4", "T3", "T2", "T1"];
  
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
      return finaleFile;
    }
    
    // Priority 2: Use the last file (highest tour number)
    const lastFile = calcFiles[calcFiles.length - 1];
    console.log("   Selected: Latest tour file:", lastFile.name, "from", lastFile.folder);
    return lastFile;
  }
  
  return null;
}

async function refreshStandings() {
  elements.standingsStatus.textContent = "Chargement en cours...";
  elements.standingsContainer.innerHTML = "";
  
  console.log("=== Starting refreshStandings ===");
  
  try {
    const season = activeSeason();
    let fileInfo = await findLatestCalculFile();
    
    if (!fileInfo) {
      const hasLinked = linkedDirectoryHandles.has(season.id);
      console.log("No file found. Linked:", hasLinked);
      
      if (!hasLinked) {
        // Try to automatically link the folder
        console.log("Attempting to auto-link folder...");
        elements.standingsStatus.textContent = "Liaison du dossier LGS en cours...";
        
        const linked = await linkSeasonFolder();
        if (!linked) {
          elements.standingsStatus.innerHTML = `<strong>Aucun fichier trouve.</strong><br>La liaison au dossier LGS a ete perdue (rechargement de page?). Cliquez sur "Lier le dossier LGS" pour reconnecter, puis revenez ici.`;
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
    
    const file = await fileInfo.handle.getFile();
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    
    console.log("Workbook sheets:", workbook.SheetNames);
    
    const standings = {};
    const categorySheets = { "Resultat LGS (HOMME)": "HOMME", "Resultat LGS (DAME)": "DAME" };
    const isFinaleFile = fileInfo.name.toLowerCase().includes("finale");
    
    let sheetsFound = 0;
    for (const [sheetName, category] of Object.entries(categorySheets)) {
      if (!workbook.SheetNames.includes(sheetName)) {
        console.log(`Sheet not found: ${sheetName}`);
        continue;
      }
      
      console.log(`Processing sheet: ${sheetName}, Finale file: ${isFinaleFile}`);
      sheetsFound++;
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: "A", defval: "" });
      
      console.log(`Data rows in ${sheetName}:`, data.length);
      
      const bySeriesAndTotal = {};
      let validRecords = 0;
      let skippedRecords = 0;
      
      data.forEach((row, rowIndex) => {
        // Skip first row (metadata)
        if (rowIndex === 0) return;
        
        // Extract name from column B
        const name = String(row.B || "").trim();
        
        // Series is in column E
        const series = row.E || "";
        
        const dayScore = row.AD || "";
        const dayScore2 = row.AE || "";
        const finalScore = row.AF || "";
        const finalScore2 = row.AH || "";
        const totalScore = row.AJ || "";
        const totalScore2 = row.AK || "";
        
        // Skip placeholder/header rows
        if (name === "Nom - Prénom" || name === "Nom - prenom" || !name || !series || !totalScore) {
          skippedRecords++;
          return;
        }
        
        // For Finale file: exclude players without final day score (show "En cours")
        if (isFinaleFile) {
          const hasFinalScore = finalScore && finalScore !== "En cours" && finalScore !== "";
          if (!hasFinalScore) {
            skippedRecords++;
            console.log(`Skipping ${name} (Finale): no final day score`);
            return;
          }
        }
        
        validRecords++;
        
        // Use series as key
        const seriesKey = String(series).toLowerCase();
        if (!bySeriesAndTotal[seriesKey]) bySeriesAndTotal[seriesKey] = [];
        
        bySeriesAndTotal[seriesKey].push({
          name,
          series: String(series).trim(),
          dayScore: String(dayScore).trim(),
          dayScore2: String(dayScore2).trim(),
          finalScore: String(finalScore).trim(),
          finalScore2: String(finalScore2).trim(),
          total: parseFloat(totalScore) || 0,
          totalScore: String(totalScore).trim(),
          totalScore2: String(totalScore2).trim()
        });
      });
      
      console.log(`${sheetName}: ${validRecords} valid, ${skippedRecords} skipped, series found: ${Object.keys(bySeriesAndTotal).sort().join(", ")}`);
      standings[category] = bySeriesAndTotal;
    }
    
    if (sheetsFound === 0) {
      elements.standingsStatus.textContent = `Erreur : Les feuilles "Resultat LGS (HOMME)" et "Resultat LGS (DAME)" n'ont pas ete trouvees dans ${fileInfo.name}. Verifiez le nom des onglets du classement.`;
      return;
    }
    
    renderStandings(standings, fileInfo.name);
  } catch (error) {
    elements.standingsStatus.textContent = "Erreur : Impossible de lire le fichier Excel. Verifiez qu'il n'est pas ouvert.";
    console.error(error);
  }
}

function renderStandings(standings, fileName) {
  elements.standingsContainer.innerHTML = "";
  elements.standingsStatus.textContent = `Donnees de : ${fileName}`;
  
  const categories = ["HOMME", "DAME"];
  
  for (const category of categories) {
    if (!standings[category] || Object.keys(standings[category]).length === 0) continue;
    
    const categoryDiv = document.createElement("div");
    categoryDiv.className = "standings-category";
    
    const categoryTitle = document.createElement("h3");
    categoryTitle.style.marginTop = "1.5rem";
    categoryTitle.style.marginBottom = "0.8rem";
    categoryTitle.textContent = category;
    categoryDiv.appendChild(categoryTitle);
    
    const seriesData = standings[category];
    const seriesNames = Object.keys(seriesData).sort();
    
    for (const seriesName of seriesNames) {
      const players = seriesData[seriesName];
      players.sort((a, b) => a.total - b.total);
      const top5 = players.slice(0, 5);
      
      const seriesGroup = document.createElement("div");
      seriesGroup.className = "series-group";
      
      const seriesTitle = document.createElement("div");
      seriesTitle.className = "series-title";
      seriesTitle.textContent = `Serie ${seriesName}`;
      seriesGroup.appendChild(seriesTitle);
      
      top5.forEach((player, index) => {
        const row = document.createElement("div");
        row.className = "player-row";
        
        const rankCell = document.createElement("div");
        rankCell.className = "rank";
        rankCell.textContent = String(index + 1);
        
        const nameCell = document.createElement("div");
        nameCell.className = "name";
        nameCell.textContent = player.name;
        
        const dayScoreCell = document.createElement("div");
        dayScoreCell.className = "score-cell";
        dayScoreCell.innerHTML = `<div class="score-label">Meilleur tour</div><div class="score-value">${player.dayScore}${player.dayScore2 ? " " + player.dayScore2 : ""}</div>`;
        
        const finalScoreCell = document.createElement("div");
        finalScoreCell.className = "score-cell";
        finalScoreCell.innerHTML = `<div class="score-label">Finale</div><div class="score-value">${player.finalScore}${player.finalScore2 ? " " + player.finalScore2 : ""}</div>`;
        
        const totalCell = document.createElement("div");
        totalCell.className = "score-cell";
        totalCell.style.fontWeight = "700";
        totalCell.innerHTML = `<div class="score-label">Total LGS</div><div class="score-value">${player.totalScore}${player.totalScore2 ? " " + player.totalScore2 : ""}</div>`;
        
        row.appendChild(rankCell);
        row.appendChild(nameCell);
        row.appendChild(dayScoreCell);
        row.appendChild(finalScoreCell);
        row.appendChild(totalCell);
        
        seriesGroup.appendChild(row);
      });
      
      categoryDiv.appendChild(seriesGroup);
    }
    
    elements.standingsContainer.appendChild(categoryDiv);
  }
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
elements.refreshStandingsButton.addEventListener("click", refreshStandings);

render();
