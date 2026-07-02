const fs = require('fs');
const path = require('path');

let STATE_FILE_PATH = path.join(__dirname, '..', 'indexer_state.json');
let LEDGER_FILE_PATH = path.join(__dirname, '..', 'uploaded_files.ledger');
let UNPARSEABLE_LEDGER_PATH = path.join(__dirname, '..', 'unparseable_files.ledger');

if (process.pkg) {
  const execDir = path.dirname(process.execPath);
  STATE_FILE_PATH = path.join(execDir, 'indexer_state.json');
  LEDGER_FILE_PATH = path.join(execDir, 'uploaded_files.ledger');
  UNPARSEABLE_LEDGER_PATH = path.join(execDir, 'unparseable_files.ledger');
}

try {
  const electron = require('electron');
  const app = electron.app || (electron.remote && electron.remote.app);
  if (app) {
    const userData = app.getPath('userData');
    STATE_FILE_PATH = path.join(userData, 'indexer_state.json');
    LEDGER_FILE_PATH = path.join(userData, 'uploaded_files.ledger');
    UNPARSEABLE_LEDGER_PATH = path.join(userData, 'unparseable_files.ledger');
  }
} catch (e) {
  // Ignore, running outside Electron
}

const DEFAULT_STATE = {
  folders: [],
  indexingStatus: 'idle',
  schedulerStatus: 'inactive',
  stats: {
    totalFilesFound: 0,
    filesIndexed: 0,
    filesFailed: 0,
    filesSkipped: 0,
    filesIndexedThisSession: 0,
    currentFilePath: '',
    speed: 0,
  },
  logs: [],
  recentErrors: []
};

let currentState = null;
const uploadedSet = new Set();
const unparseableSet = new Set();
let pendingLedgerLines = [];
let pendingUnparseableLines = [];
let saveTimer = null;
let ledgerMigrated = false;

function normalizePath(filePath) {
  return path.normalize(filePath).toLowerCase();
}

function loadLedgerFromFile() {
  if (!fs.existsSync(LEDGER_FILE_PATH)) return;

  const buf = fs.readFileSync(LEDGER_FILE_PATH, 'utf8');
  let start = 0;
  for (let i = 0; i < buf.length; i++) {
    if (buf[i] === '\n') {
      const line = buf.slice(start, i).trim();
      if (line) uploadedSet.add(line);
      start = i + 1;
    }
  }
  if (start < buf.length) {
    const line = buf.slice(start).trim();
    if (line) uploadedSet.add(line);
  }
}

function loadUnparseableLedgerFromFile() {
  if (!fs.existsSync(UNPARSEABLE_LEDGER_PATH)) return;

  const buf = fs.readFileSync(UNPARSEABLE_LEDGER_PATH, 'utf8');
  let start = 0;
  for (let i = 0; i < buf.length; i++) {
    if (buf[i] === '\n') {
      const line = buf.slice(start, i).trim();
      if (line) unparseableSet.add(line);
      start = i + 1;
    }
  }
  if (start < buf.length) {
    const line = buf.slice(start).trim();
    if (line) unparseableSet.add(line);
  }
}

function migrateUploadedFilesFromJson() {
  if (ledgerMigrated || !currentState) return;
  ledgerMigrated = true;

  const legacy = currentState.uploadedFiles;
  if (!legacy || typeof legacy !== 'object') return;

  const paths = Object.keys(legacy);
  if (paths.length === 0) {
    delete currentState.uploadedFiles;
    return;
  }

  const lines = [];
  for (const filePath of paths) {
    const normalized = normalizePath(filePath);
    if (!uploadedSet.has(normalized)) {
      uploadedSet.add(normalized);
      lines.push(normalized);
    }
  }

  if (lines.length > 0) {
    fs.appendFileSync(LEDGER_FILE_PATH, `${lines.join('\n')}\n`, 'utf8');
  }

  delete currentState.uploadedFiles;
  console.log(`Migrated ${paths.length} uploaded file records to ledger.`);
}

function flushLedgerSync() {
  if (pendingLedgerLines.length > 0) {
    fs.appendFileSync(LEDGER_FILE_PATH, `${pendingLedgerLines.join('\n')}\n`, 'utf8');
    pendingLedgerLines = [];
  }
  if (pendingUnparseableLines.length > 0) {
    fs.appendFileSync(UNPARSEABLE_LEDGER_PATH, `${pendingUnparseableLines.join('\n')}\n`, 'utf8');
    pendingUnparseableLines = [];
  }
}

async function flushLedgerAsync() {
  if (pendingLedgerLines.length > 0) {
    const lines = `${pendingLedgerLines.join('\n')}\n`;
    pendingLedgerLines = [];
    await fs.promises.appendFile(LEDGER_FILE_PATH, lines, 'utf8');
  }
  if (pendingUnparseableLines.length > 0) {
    const lines = `${pendingUnparseableLines.join('\n')}\n`;
    pendingUnparseableLines = [];
    await fs.promises.appendFile(UNPARSEABLE_LEDGER_PATH, lines, 'utf8');
  }
}

function writeStateToDiskSync() {
  flushLedgerSync();
  fs.writeFileSync(STATE_FILE_PATH, JSON.stringify(currentState), 'utf8');
}

let saveInProgress = false;
let saveQueued = false;

async function writeStateToDiskAsync() {
  await flushLedgerAsync();
  await fs.promises.writeFile(STATE_FILE_PATH, JSON.stringify(currentState), 'utf8');
}

function scheduleSave(delayMs = 2000) {
  if (saveTimer) return;
  saveTimer = setTimeout(() => {
    saveTimer = null;
    if (saveInProgress) {
      saveQueued = true;
      return;
    }
    saveInProgress = true;
    writeStateToDiskAsync()
      .catch(err => console.error('Error saving state file:', err))
      .finally(() => {
        saveInProgress = false;
        if (saveQueued) {
          saveQueued = false;
          scheduleSave(0);
        }
      });
  }, delayMs);
}

function flushStateNow() {
  if (saveTimer) {
    clearTimeout(saveTimer);
    saveTimer = null;
  }
  scheduleSave(0);
}

function loadState() {
  if (currentState) return currentState;

  try {
    if (fs.existsSync(STATE_FILE_PATH)) {
      const rawData = fs.readFileSync(STATE_FILE_PATH, 'utf8');
      currentState = JSON.parse(rawData);

      currentState.folders = currentState.folders || [];
      currentState.indexingStatus = currentState.indexingStatus || 'idle';
      currentState.schedulerStatus = currentState.schedulerStatus || 'inactive';
      currentState.stats = { ...DEFAULT_STATE.stats, ...(currentState.stats || {}) };
      currentState.logs = currentState.logs || [];
      currentState.recentErrors = currentState.recentErrors || [];
    } else {
      currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
      writeStateToDiskSync();
    }
  } catch (err) {
    console.error('Error loading state file, initializing default state:', err);
    currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
  }

  loadLedgerFromFile();
  loadUnparseableLedgerFromFile();
  migrateUploadedFilesFromJson();
  return currentState;
}

function saveState() {
  try {
    if (!currentState) {
      currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
    }
    if (saveTimer) {
      clearTimeout(saveTimer);
      saveTimer = null;
    }
    writeStateToDiskSync();
  } catch (err) {
    console.error('Error saving state file:', err);
  }
}

function getPublicState() {
  const s = loadState();
  const { uploadedFiles, ...publicState } = s;
  return {
    ...publicState,
    uploadedFilesCount: uploadedSet.size,
    unparseableFilesCount: unparseableSet.size
  };
}

function getIndexingStatus() {
  if (!currentState) loadState();
  return currentState.indexingStatus;
}

function addLog(message) {
  const timestamp = new Date().toLocaleTimeString();
  const logLine = `[${timestamp}] ${message}`;

  if (!currentState) loadState();

  currentState.logs.push(logLine);

  // Bulk-trim when over limit — O(n) shift() on every entry is expensive
  // at high indexing rates. Splice 50 at once to amortise the cost.
  if (currentState.logs.length > 200) {
    currentState.logs.splice(0, 50);
  }

  console.log(logLine);
}

function addErrorLog(filePath, errorMessage) {
  if (!currentState) loadState();
  const timestamp = new Date().toLocaleTimeString();
  currentState.recentErrors.unshift({ timestamp, filePath, error: errorMessage });
  if (currentState.recentErrors.length > 50) {
    currentState.recentErrors.pop();
  }
  scheduleSave(500);
}

function getFolders() {
  if (!currentState) loadState();
  return currentState.folders;
}

function addFolder(folderPath, type = 'local', description = '', collectionId = null) {
  if (!currentState) loadState();

  const normalizedPath = path.normalize(folderPath).trim();

  const exists = currentState.folders.some(
    f => path.normalize(f.path).toLowerCase() === normalizedPath.toLowerCase()
  );

  if (!exists) {
    currentState.folders.push({
      path: folderPath,
      type,
      description: description || path.basename(folderPath) || folderPath,
      collectionId,
      isPublic: true,
      allowedUsers: [],
      addedAt: new Date().toISOString()
    });
    addLog(`Added location to index queue: ${folderPath} (${type})`);
    saveState();
    return true;
  }
  return false;
}

function removeFolder(folderPath) {
  if (!currentState) loadState();

  const normalizedPath = path.normalize(folderPath).toLowerCase().trim();
  const originalLength = currentState.folders.length;

  currentState.folders = currentState.folders.filter(
    f => path.normalize(f.path).toLowerCase() !== normalizedPath
  );

  if (currentState.folders.length < originalLength) {
    addLog(`Removed location from index queue: ${folderPath}`);
    saveState();
    return true;
  }
  return false;
}

function updateFolderPermissions(folderPath, isPublic, allowedUsers = []) {
  if (!currentState) loadState();
  const normalizedPath = path.normalize(folderPath).trim().toLowerCase();

  const folder = currentState.folders.find(
    f => path.normalize(f.path).trim().toLowerCase() === normalizedPath
  );

  if (folder) {
    folder.isPublic = isPublic;
    folder.allowedUsers = Array.isArray(allowedUsers) ? allowedUsers : [];
    saveState();
    return true;
  }
  return false;
}

function isFileUploaded(filePath) {
  if (!currentState) loadState();
  return uploadedSet.has(normalizePath(filePath));
}

function markFileUploaded(filePath) {
  if (!currentState) loadState();
  const normalized = normalizePath(filePath);
  if (uploadedSet.has(normalized)) return;
  uploadedSet.add(normalized);
  pendingLedgerLines.push(normalized);
}

function getUploadedCount() {
  if (!currentState) loadState();
  return uploadedSet.size;
}

function isFileUnparseable(filePath) {
  if (!currentState) loadState();
  return unparseableSet.has(normalizePath(filePath));
}

function markFileUnparseable(filePath) {
  if (!currentState) loadState();
  const normalized = normalizePath(filePath);
  if (unparseableSet.has(normalized)) return;
  unparseableSet.add(normalized);
  pendingUnparseableLines.push(normalized);
  flushStateNow();
}

function getUnparseableCount() {
  if (!currentState) loadState();
  return unparseableSet.size;
}

function updateIndexingStatus(status) {
  if (!currentState) loadState();
  currentState.indexingStatus = status;
  flushStateNow();
}

function updateSchedulerStatus(status) {
  if (!currentState) loadState();
  currentState.schedulerStatus = status;
  flushStateNow();
}

function getStats() {
  if (!currentState) loadState();
  return currentState.stats;
}

function updateStats(newStats, options = {}) {
  const { persist = true, immediate = false } = options;
  if (!currentState) loadState();
  currentState.stats = {
    ...currentState.stats,
    ...newStats
  };
  if (!persist) return;
  if (immediate) {
    flushStateNow();
  } else {
    scheduleSave();
  }
}

function resetProgress() {
  if (!currentState) loadState();

  uploadedSet.clear();
  pendingLedgerLines = [];
  unparseableSet.clear();
  pendingUnparseableLines = [];
  if (fs.existsSync(LEDGER_FILE_PATH)) {
    fs.writeFileSync(LEDGER_FILE_PATH, '', 'utf8');
  }
  if (fs.existsSync(UNPARSEABLE_LEDGER_PATH)) {
    fs.writeFileSync(UNPARSEABLE_LEDGER_PATH, '', 'utf8');
  }

  currentState.indexingStatus = 'idle';
  currentState.stats = { ...DEFAULT_STATE.stats };
  currentState.logs = [];
  currentState.recentErrors = [];
  addLog('Indexer progress reset. Uploaded files log cleared.');
  saveState();
}

function clearAll() {
  currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
  uploadedSet.clear();
  pendingLedgerLines = [];
  unparseableSet.clear();
  pendingUnparseableLines = [];
  if (fs.existsSync(LEDGER_FILE_PATH)) {
    fs.writeFileSync(LEDGER_FILE_PATH, '', 'utf8');
  }
  if (fs.existsSync(UNPARSEABLE_LEDGER_PATH)) {
    fs.writeFileSync(UNPARSEABLE_LEDGER_PATH, '', 'utf8');
  }
  addLog('Indexer state fully reset. All folders and progress cleared.');
  saveState();
}

loadState();
let startupStateChanged = false;
if (currentState && (currentState.indexingStatus === 'scanning' || currentState.indexingStatus === 'uploading')) {
  currentState.indexingStatus = 'paused';
  startupStateChanged = true;
}
if (startupStateChanged) {
  saveState();
}

module.exports = {
  loadState,
  getPublicState,
  getIndexingStatus,
  saveState,
  getFolders,
  addFolder,
  removeFolder,
  isFileUploaded,
  markFileUploaded,
  isFileUnparseable,
  markFileUnparseable,
  getUploadedCount,
  getUnparseableCount,
  updateIndexingStatus,
  updateSchedulerStatus,
  getStats,
  updateStats,
  resetProgress,
  clearAll,
  addLog,
  addErrorLog,
  updateFolderPermissions
};
