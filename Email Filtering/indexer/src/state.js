const fs = require('fs');
const path = require('path');

let STATE_FILE_PATH = path.join(__dirname, '..', 'indexer_state.json');

if (process.pkg) {
  // When bundled with pkg, __dirname is read-only inside the snapshot.
  // We must write to the folder containing the executable instead.
  STATE_FILE_PATH = path.join(path.dirname(process.execPath), 'indexer_state.json');
}

try {
  const electron = require('electron');
  const app = electron.app || (electron.remote && electron.remote.app);
  if (app) {
    STATE_FILE_PATH = path.join(app.getPath('userData'), 'indexer_state.json');
  }
} catch (e) {
  // Ignore, running outside Electron
}

const DEFAULT_STATE = {
  folders: [], // Array of { path, type: 'local'|'network'|'collection', description, addedAt }
  uploadedFiles: {}, // Hash map of { [filePath]: timestamp }
  indexingStatus: 'idle', // 'idle' | 'scanning' | 'uploading' | 'paused' | 'completed'
  schedulerStatus: 'inactive', // 'inactive' | 'active'
  stats: {
    totalFilesFound: 0,
    filesIndexed: 0,
    filesFailed: 0,
    currentFilePath: '',
    speed: 0, // emails/sec
  },
  logs: [], // Array of log lines shown in UI console
  recentErrors: [] // Array of { timestamp, filePath, error }
};

let currentState = null;

function loadState() {
  try {
    if (fs.existsSync(STATE_FILE_PATH)) {
      const rawData = fs.readFileSync(STATE_FILE_PATH, 'utf8');
      currentState = JSON.parse(rawData);
      
      // Ensure defaults exist for missing keys (backwards compatibility)
      currentState.folders = currentState.folders || [];
      currentState.uploadedFiles = currentState.uploadedFiles || {};
      currentState.indexingStatus = currentState.indexingStatus || 'idle';
      currentState.schedulerStatus = currentState.schedulerStatus || 'inactive';
      currentState.stats = currentState.stats || { ...DEFAULT_STATE.stats };
      currentState.logs = currentState.logs || [];
      currentState.recentErrors = currentState.recentErrors || [];
    } else {
      currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
      saveState();
    }
  } catch (err) {
    console.error('Error loading state file, initializing default state:', err);
    currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
  }
  return currentState;
}

function saveState() {
  try {
    if (!currentState) {
      currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
    }
    fs.writeFileSync(STATE_FILE_PATH, JSON.stringify(currentState, null, 2), 'utf8');
  } catch (err) {
    console.error('Error saving state file:', err);
  }
}

function addLog(message) {
  const timestamp = new Date().toLocaleTimeString();
  const logLine = `[${timestamp}] ${message}`;
  
  if (!currentState) loadState();
  
  currentState.logs.push(logLine);
  
  // Cap logs at 500 lines to prevent memory bloat
  if (currentState.logs.length > 500) {
    currentState.logs.shift();
  }
  
  // Print to terminal console too
  console.log(logLine);
}

function addErrorLog(filePath, errorMessage) {
  if (!currentState) loadState();
  const timestamp = new Date().toLocaleTimeString();
  currentState.recentErrors.unshift({ timestamp, filePath, error: errorMessage });
  if (currentState.recentErrors.length > 50) {
    currentState.recentErrors.pop();
  }
  saveState();
}

function getFolders() {
  if (!currentState) loadState();
  return currentState.folders;
}

function addFolder(folderPath, type = 'local', description = '', collectionId = null) {
  if (!currentState) loadState();
  
  const normalizedPath = path.normalize(folderPath).trim();
  
  // Check for duplicates
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
  const normalized = path.normalize(filePath).toLowerCase();
  return !!currentState.uploadedFiles[normalized];
}

function markFileUploaded(filePath) {
  if (!currentState) loadState();
  const normalized = path.normalize(filePath).toLowerCase();
  currentState.uploadedFiles[normalized] = Date.now();
}

function updateIndexingStatus(status) {
  if (!currentState) loadState();
  currentState.indexingStatus = status;
  saveState();
}

function updateSchedulerStatus(status) {
  if (!currentState) loadState();
  currentState.schedulerStatus = status;
  saveState();
}

function getStats() {
  if (!currentState) loadState();
  return currentState.stats;
}

function updateStats(newStats) {
  if (!currentState) loadState();
  currentState.stats = {
    ...currentState.stats,
    ...newStats
  };
  saveState();
}

function resetProgress() {
  if (!currentState) loadState();
  
  currentState.uploadedFiles = {};
  currentState.indexingStatus = 'idle';
  currentState.stats = {
    totalFilesFound: 0,
    filesIndexed: 0,
    filesFailed: 0,
    currentFilePath: '',
    speed: 0
  };
  currentState.logs = [];
  currentState.recentErrors = [];
  addLog('Indexer progress reset. Uploaded files log cleared.');
  saveState();
}

function clearAll() {
  currentState = JSON.parse(JSON.stringify(DEFAULT_STATE));
  addLog('Indexer state fully reset. All folders and progress cleared.');
  saveState();
}

// Initialization: If the app was closed abruptly while running, reset statuses on startup
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
  saveState,
  getFolders,
  addFolder,
  removeFolder,
  isFileUploaded,
  markFileUploaded,
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
