require('dotenv').config();
const { MeiliSearch } = require('meilisearch');
const path = require('path');
const { scanDirectory } = require('./scanner');
const { parseEmailFile } = require('./parser');
const state = require('./state');

// Initialize Meilisearch Client
const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

// Ensure filterable attributes are set for search scopes
emailIndex.updateFilterableAttributes([
  'indexedRootPath',
  'indexedRootType',
  'collectionId'
]).catch(err => console.error("Failed to set filterable attributes:", err));

let activeIndexerPromise = null;
let schedulerInterval = null;
const SCHEDULER_INTERVAL_MS = 15 * 60 * 1000; // 15 minutes

let currentRunStats = {
  indexedInSession: 0,
  startTime: 0
};

/**
 * Main indexer runner logic.
 * Walks through all registered folders, parses new emails, and uploads in batches.
 */
async function runIndexing() {
  state.updateIndexingStatus('scanning');
  state.addLog('Starting scan across all target locations...');
  
  const folders = state.getFolders();
  if (folders.length === 0) {
    state.addLog('No target folders configured. Indexer stopped.');
    state.updateIndexingStatus('idle');
    return;
  }
  
  // 1. Gather all files
  let allFiles = [];
  const seenPaths = new Set();
  
  for (const folder of folders) {
    state.addLog(`Scanning directory: ${folder.path}...`);
    try {
      const files = scanDirectory(folder.path);
      for (const file of files) {
        if (!seenPaths.has(file)) {
          seenPaths.add(file);
          allFiles.push({ filePath: file, folder });
        }
      }
    } catch (err) {
      state.addLog(`Failed scanning directory ${folder.path}: ${err.message}`);
    }
  }
  
  const totalFiles = allFiles.length;
  state.updateStats({ totalFilesFound: totalFiles });
  state.addLog(`Found ${totalFiles} total email files (.msg/.eml) across target locations.`);
  
  state.updateIndexingStatus('uploading');
  currentRunStats.indexedInSession = 0;
  currentRunStats.startTime = Date.now();
  
  let batch = [];
  let batchFilePaths = [];
  const BATCH_SIZE = 100;
  
  for (let i = 0; i < allFiles.length; i++) {
    // Check if user paused or stopped the indexer
    const currentState = state.loadState();
    if (currentState.indexingStatus === 'paused' || currentState.indexingStatus === 'idle') {
      state.addLog('Indexing paused by user.');
      // If we have remaining items in the batch, upload them first
      if (batch.length > 0) {
        await uploadBatch(batch, batchFilePaths);
      }
      return;
    }
    
    const { filePath, folder } = allFiles[i];
    
    // Check if already indexed
    if (state.isFileUploaded(filePath)) {
      continue;
    }
    
    state.updateStats({ currentFilePath: filePath });
    
    try {
      const parsedEmail = await parseEmailFile(filePath);
      
      // Meilisearch requires a unique identifier 'id'.
      // We can generate a clean ID using a simple hash of the file path.
      // Meilisearch primary key constraints: alphanumeric, hyphens, and underscores.
      const rawId = Buffer.from(filePath).toString('base64');
      const safeId = rawId.replace(/[^a-zA-Z0-9_-]/g, 'x').substring(0, 64);
      
      batch.push({
        id: safeId,
        ...parsedEmail,
        indexedRootPath: folder.path,
        indexedRootType: folder.type || 'local',
        collectionId: folder.collectionId || null
      });
      batchFilePaths.push(filePath);
      
      if (batch.length >= BATCH_SIZE) {
        await uploadBatch(batch, batchFilePaths);
        batch = [];
        batchFilePaths = [];
      }
    } catch (err) {
      const stats = state.getStats();
      state.updateStats({ filesFailed: stats.filesFailed + 1 });
      state.addErrorLog(filePath, err.message);
      state.addLog(`[Error] ${path.basename(filePath)}: ${err.message}`);
    }
  }
  
  // Upload remaining batch items
  if (batch.length > 0) {
    await uploadBatch(batch, batchFilePaths);
  }
  
  // Finalize indexing status
  const finalState = state.loadState();
  if (finalState.indexingStatus === 'uploading') {
    state.updateIndexingStatus('completed');
    state.addLog('Indexing process completed successfully!');
    state.updateStats({ currentFilePath: '' });
  }
}

/**
 * Uploads a batch of emails to Meilisearch and saves the progress to state.
 */
async function uploadBatch(emailBatch, paths) {
  try {
    state.addLog(`Uploading batch of ${emailBatch.length} emails to Meilisearch...`);
    
    // Send to Meilisearch
    await emailIndex.addDocuments(emailBatch);
    
    // Mark files as uploaded
    for (const filePath of paths) {
      state.markFileUploaded(filePath);
    }
    
    // Update stats
    currentRunStats.indexedInSession += emailBatch.length;
    const elapsedSeconds = (Date.now() - currentRunStats.startTime) / 1000;
    const speed = Math.round(currentRunStats.indexedInSession / (elapsedSeconds || 1));
    
    const stats = state.getStats();
    state.updateStats({
      filesIndexed: stats.filesIndexed + emailBatch.length,
      speed: speed
    });
    
    state.addLog(`Successfully uploaded batch. Total indexed this session: ${currentRunStats.indexedInSession}`);
    state.saveState();
  } catch (err) {
    state.addLog(`Meilisearch upload batch failed: ${err.message}`);
    const stats = state.getStats();
    state.updateStats({
      filesFailed: stats.filesFailed + emailBatch.length
    });
    state.saveState();
    throw err;
  }
}

function start() {
  const currentState = state.loadState();
  if (currentState.indexingStatus === 'scanning' || currentState.indexingStatus === 'uploading') {
    state.addLog('Indexer is already running.');
    return;
  }
  
  state.addLog('Starting indexer job...');
  activeIndexerPromise = runIndexing()
    .catch(err => {
      state.addLog(`Critical error in indexer runner: ${err.message}`);
      state.updateIndexingStatus('idle');
    })
    .finally(() => {
      activeIndexerPromise = null;
    });
}

function pause() {
  const currentState = state.loadState();
  if (currentState.indexingStatus === 'scanning' || currentState.indexingStatus === 'uploading') {
    state.addLog('Pausing indexer job...');
    state.updateIndexingStatus('paused');
  } else {
    state.addLog('Indexer is not running.');
  }
}

function reset() {
  pause();
  state.resetProgress();
}

function startScheduler() {
  const currentState = state.loadState();
  if (currentState.schedulerStatus === 'active') {
    state.addLog('Live Scheduler is already running.');
    return;
  }
  
  state.updateSchedulerStatus('active');
  state.addLog(`Live Scheduler activated. Delta-sync will run every 15 minutes.`);
  
  // Optionally start immediately if idle
  if (currentState.indexingStatus !== 'scanning' && currentState.indexingStatus !== 'uploading') {
    start();
  }
  
  schedulerInterval = setInterval(() => {
    const st = state.loadState();
    if (st.indexingStatus !== 'scanning' && st.indexingStatus !== 'uploading') {
      state.addLog('Live Scheduler triggered: Starting periodic delta-sync...');
      start();
    }
  }, SCHEDULER_INTERVAL_MS);
}

function stopScheduler() {
  const currentState = state.loadState();
  if (currentState.schedulerStatus === 'active') {
    if (schedulerInterval) {
      clearInterval(schedulerInterval);
      schedulerInterval = null;
    }
    state.updateSchedulerStatus('inactive');
    state.addLog('Live Scheduler deactivated.');
  }
}

module.exports = {
  start,
  pause,
  reset,
  startScheduler,
  stopScheduler
};
