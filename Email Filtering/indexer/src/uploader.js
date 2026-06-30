const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });
const { MeiliSearch } = require('meilisearch');
const { scanDirectory } = require('./scanner');
const { parseEmailFile } = require('./parser');
const state = require('./state');

// Initialize Meilisearch Client
const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

// Ensure filterable and searchable attributes are set for search scopes
emailIndex.updateFilterableAttributes([
  'indexedRootPath',
  'indexedRootType',
  'collectionId',
  'hasAttachments',
  'sentAt',
  'isPublic',
  'allowedUsers'
]).catch(err => console.error("Failed to set filterable attributes:", err));

// Declare searchable attributes with priority order
// filePath MUST be searchable so job numbers embedded in path are found
emailIndex.updateSearchableAttributes([
  'subject',
  'sender',
  'recipients',
  'cc',
  'bcc',
  'body',
  'filePath'
]).catch(err => console.error("Failed to set searchable attributes:", err));

let activeIndexerPromise = null;
let schedulerInterval = null;
const SCHEDULER_INTERVAL_MS = 15 * 60 * 1000; // 15 minutes

let currentRunStats = {
  indexedInSession: 0,
  startTime: 0
};

/**
 * Main indexer runner logic.
 * Walks through registered folders (or specific target paths), parses new emails, and uploads in batches.
 */
async function runIndexing(targetPaths = []) {
  state.updateIndexingStatus('scanning');
  let folders = state.getFolders();
  
  // Filter for specific target paths if provided
  if (targetPaths && targetPaths.length > 0) {
    folders = folders.filter(f => targetPaths.includes(f.path));
    if (folders.length === 0) {
      state.addLog('No matching target folders found. Indexer stopped.');
      state.updateIndexingStatus('idle');
      return;
    }
    state.addLog(`Starting targeted scan across ${folders.length} selected locations...`);
  } else {
    if (folders.length === 0) {
      state.addLog('No target folders configured. Indexer stopped.');
      state.updateIndexingStatus('idle');
      return;
    }
    state.addLog('Starting scan across all target locations...');
  }
  
  // 1. Gather all files
  let allFiles = [];
  const seenPaths = new Set();
  
  for (const folder of folders) {
    state.addLog(`Scanning directory: ${folder.path}...`);
    try {
      const files = await scanDirectory(folder.path);
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
      
      const rawId = Buffer.from(filePath).toString('base64');
      const safeId = rawId.replace(/[^a-zA-Z0-9_-]/g, 'x').substring(0, 64);
      
      batch.push({
        id: safeId,
        ...parsedEmail,
        body: (parsedEmail.body || '').substring(0, 50000),
        indexedRootPath: folder.path,
        indexedRootType: folder.type || 'local',
        collectionId: folder.type === 'collection' ? (folder.description || folder.collectionId) : (folder.collectionId || null),
        isPublic: folder.isPublic !== false,
        allowedUsers: (folder.allowedUsers || []).map(u => u.toLowerCase())
      });
      batchFilePaths.push(filePath);
    } catch (err) {
      const stats = state.getStats();
      state.updateStats({ filesFailed: stats.filesFailed + 1 });
      state.addErrorLog(filePath, err.message);
      state.addLog(`[Error] ${path.basename(filePath)}: ${err.message}`);
    }
    
    // Perform upload outside the parsing try/catch to ensure we can clear the batch even on failure
    if (batch.length >= BATCH_SIZE) {
      try {
        await uploadBatch(batch, batchFilePaths);
      } catch (uploadErr) {
        // Error is already logged inside uploadBatch, but we MUST clear the batch
        // to prevent an infinite loop of failing uploads causing an OOM crash.
        console.error("Batch upload failed, clearing batch to continue...", uploadErr.message);
      } finally {
        batch = [];
        batchFilePaths = [];
      }
    }
  }
  
  // Upload remaining batch items
  if (batch.length > 0) {
    try {
      await uploadBatch(batch, batchFilePaths);
    } catch (uploadErr) {
      console.error("Final batch upload failed:", uploadErr.message);
    }
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
    await emailIndex.addDocuments(emailBatch, { primaryKey: 'id' });
    
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
    state.addErrorLog('Meilisearch Batch Upload', err.message);
    const stats = state.getStats();
    state.updateStats({
      filesFailed: stats.filesFailed + emailBatch.length
    });
    state.saveState();
    throw err;
  }
}

function start(targetPaths = []) {
  const currentState = state.loadState();
  if (currentState.indexingStatus === 'scanning' || currentState.indexingStatus === 'uploading') {
    state.addLog('Indexer is already running.');
    return;
  }
  
  state.addLog('Starting indexer job...');
  activeIndexerPromise = runIndexing(targetPaths)
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

function startScheduler(isInit = false) {
  const currentState = state.loadState();
  if (currentState.schedulerStatus === 'active' && !isInit) {
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

function reset() {
  state.resetProgress();
}

module.exports = {
  start,
  pause,
  reset,
  startScheduler,
  stopScheduler
};
