const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });
const { MeiliSearch } = require('meilisearch');
const { scanDirectory } = require('./scanner');
const { parseInWorker, WORKER_COUNT } = require('./parsePool');
const state = require('./state');

// ── Meilisearch client ────────────────────────────────────────────────────────
const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

emailIndex.updateFilterableAttributes([
  'indexedRootPath',
  'indexedRootType',
  'collectionId',
  'hasAttachments',
  'sentAt',
  'isPublic',
  'allowedUsers'
]).catch(err => console.error('Failed to set filterable attributes:', err));

emailIndex.updateSearchableAttributes([
  'subject',
  'sender',
  'recipients',
  'cc',
  'bcc',
  'body',
  'filePath'
]).catch(err => console.error('Failed to set searchable attributes:', err));

// ── Tuning constants ──────────────────────────────────────────────────────────
// BATCH_SIZE 500: Railway Pro Meilisearch has 24 GB RAM — 500 emails (~25 MB
// worst-case) is comfortably within limits and cuts API round-trips 5x vs 100.
const BATCH_SIZE = 250;

// CONCURRENCY: number of parse jobs running in parallel. Matches worker pool
// size so we saturate all workers without building an unbounded queue.
const CONCURRENCY = WORKER_COUNT;

// How often to yield to the Node.js event loop (keeps HTTP endpoints responsive)
const YIELD_EVERY = 50;

// How often to flush stats to disk (less I/O than every 25 files)
const STATS_FLUSH_EVERY = 100;

// How often to log a "still working" heartbeat (every N files indexed)
const HEARTBEAT_EVERY = 500;
// ─────────────────────────────────────────────────────────────────────────────

let activeIndexerPromise = null;
let schedulerInterval = null;
const SCHEDULER_INTERVAL_MS = 15 * 60 * 1000;

let currentRunStats = {
  indexedInSession: 0,
  skippedInSession: 0,
  startTime: 0
};

function yieldToEventLoop() {
  return new Promise(resolve => setImmediate(resolve));
}

function isIndexingStopped() {
  const status = state.getIndexingStatus();
  return status === 'paused' || status === 'idle';
}

// ── Core indexer ──────────────────────────────────────────────────────────────

async function runIndexing(targetPaths = []) {
  state.updateIndexingStatus('scanning');
  let folders = state.getFolders();

  if (targetPaths && targetPaths.length > 0) {
    folders = folders.filter(f => targetPaths.includes(f.path));
    if (folders.length === 0) {
      state.addLog('No matching target folders found. Indexer stopped.');
      state.updateIndexingStatus('idle');
      return;
    }
    state.addLog(`Starting targeted scan across ${folders.length} selected location(s)...`);
  } else {
    if (folders.length === 0) {
      state.addLog('No target folders configured. Indexer stopped.');
      state.updateIndexingStatus('idle');
      return;
    }
    state.addLog(`Starting scan across all target locations... (${CONCURRENCY} parallel parsers, batch size ${BATCH_SIZE})`);
  }

  state.updateStats({
    totalFilesFound: 0,
    filesSkipped: 0,
    filesIndexedThisSession: 0,
    currentFilePath: ''
  }, { immediate: true });
  state.updateIndexingStatus('uploading');

  currentRunStats.indexedInSession = 0;
  currentRunStats.skippedInSession = 0;
  currentRunStats.startTime = Date.now();

  let totalFiles = 0;
  let filesSkipped = 0;

  // Pending parsed emails waiting to be batched
  let batch = [];
  let batchFilePaths = [];

  // In-flight parse promises (sliding window of CONCURRENCY width)
  const inFlight = new Map(); // filePath → { promise, folder }

  /**
   * Drains inFlight promises: waits for the next settled one,
   * adds it to the batch and flushes when BATCH_SIZE is reached.
   */
  async function drainOne() {
    if (inFlight.size === 0) return;

    // Race all in-flight promises to get the first settled one
    const settled = await Promise.race(
      [...inFlight.entries()].map(([fp, { promise, folder }]) =>
        promise.then(
          result => ({ fp, result, folder, error: null }),
          err    => ({ fp, result: null, folder, error: err })
        )
      )
    );

    inFlight.delete(settled.fp);

    if (settled.error) {
      const stats = state.getStats();
      state.updateStats({ filesFailed: stats.filesFailed + 1 });
      state.addErrorLog(settled.fp, settled.error.message);
      state.addLog(`[Error] ${path.basename(settled.fp)}: ${settled.error.message}`);
    } else {
      const parsedEmail = settled.result;
      const folder = settled.folder;
      const rawId = Buffer.from(settled.fp).toString('base64');
      const safeId = rawId.replace(/[^a-zA-Z0-9_-]/g, 'x').substring(0, 64);

      batch.push({
        id: safeId,
        // Parser already applies toSearchableText internally — no need to double-process
        subject:    parsedEmail.subject,
        sender:     parsedEmail.sender,
        recipients: parsedEmail.recipients,
        cc:         parsedEmail.cc,
        bcc:        parsedEmail.bcc,
        sentAt:     parsedEmail.sentAt,
        body:       parsedEmail.body,
        hasAttachments: parsedEmail.hasAttachments,
        filePath:   parsedEmail.filePath,
        comment:    parsedEmail.comment || '',
        indexedRootPath: folder.path,
        indexedRootType: folder.type || 'local',
        collectionId: folder.type === 'collection'
          ? (folder.description || folder.collectionId)
          : (folder.collectionId || null),
        isPublic:     folder.isPublic !== false,
        allowedUsers: (folder.allowedUsers || []).map(u => u.toLowerCase())
      });
      batchFilePaths.push(settled.fp);
    }

    if (batch.length >= BATCH_SIZE) {
      await flushBatch();
    }
  }

  async function flushBatch() {
    if (batch.length === 0) return;
    const b = batch;
    const bfp = batchFilePaths;
    batch = [];
    batchFilePaths = [];
    try {
      await uploadBatch(b, bfp);
    } catch (uploadErr) {
      console.error('Batch upload failed, continuing...', uploadErr.message);
    }
  }

  // ── Main scan loop ────────────────────────────────────────────────────────

  for (const folder of folders) {
    state.addLog(`Scanning: ${folder.path}`);

    try {
      for await (const filePath of scanDirectory(folder.path)) {
        totalFiles++;

        // Periodic event-loop yield (keeps server responsive)
        if (totalFiles % YIELD_EVERY === 0) {
          await yieldToEventLoop();
        }

        // Periodic stats flush (avoids hammering disk)
        if (totalFiles % STATS_FLUSH_EVERY === 0) {
          state.updateStats({ totalFilesFound: totalFiles, filesSkipped });
        }

        // Heartbeat log so the dashboard shows life
        if (totalFiles % HEARTBEAT_EVERY === 0) {
          const elapsed = Math.round((Date.now() - currentRunStats.startTime) / 1000);
          const rate = elapsed > 0
            ? Math.round(currentRunStats.indexedInSession / elapsed)
            : 0;
          state.addLog(`Progress: ${totalFiles} found | ${currentRunStats.indexedInSession} indexed | ${rate} emails/s | ${inFlight.size} parsing`);
        }

        if (isIndexingStopped()) {
          state.addLog('Indexing paused by user.');
          // Drain remaining in-flight jobs before stopping
          while (inFlight.size > 0) await drainOne();
          await flushBatch();
          state.updateStats({ totalFilesFound: totalFiles, filesSkipped }, { immediate: true });
          return;
        }

        if (state.isFileUploaded(filePath)) {
          filesSkipped++;
          currentRunStats.skippedInSession++;
          continue; // no yield here — just skip cheaply
        }

        state.updateStats({ currentFilePath: filePath }, { persist: false });

        // Start parse job (non-blocking — goes into the in-flight pool)
        const parsePromise = parseInWorker(filePath);
        inFlight.set(filePath, { promise: parsePromise, folder });

        // If we've filled the concurrency window, wait for one to finish
        if (inFlight.size >= CONCURRENCY) {
          await drainOne();
        }
      }
    } catch (err) {
      state.addLog(`Failed scanning ${folder.path}: ${err.message}`);
    }
  }

  // Drain all remaining in-flight parse jobs
  while (inFlight.size > 0) {
    await drainOne();
  }

  // Final flush for any remaining batch
  await flushBatch();

  state.updateStats({
    totalFilesFound: totalFiles,
    filesSkipped,
    currentFilePath: ''
  }, { immediate: true });

  state.addLog(`Scan complete. ${totalFiles} email files found.`);

  if (state.getIndexingStatus() === 'uploading') {
    const finalStats = state.getStats();
    state.updateIndexingStatus('completed');
    state.addLog(
      `Indexing finished! ${finalStats.filesIndexed || 0} indexed, ` +
      `${finalStats.filesFailed || 0} errors, ` +
      `${filesSkipped} already-indexed skipped.`
    );
    state.updateStats({ currentFilePath: '', speed: 0 }, { immediate: true });
  }
}

// ── Batch upload ──────────────────────────────────────────────────────────────

async function uploadBatch(emailBatch, paths) {
  try {
    state.addLog(`Uploading batch of ${emailBatch.length} emails to Meilisearch...`);

    await emailIndex.addDocuments(emailBatch, { primaryKey: 'id' });

    for (const filePath of paths) {
      state.markFileUploaded(filePath);
    }

    currentRunStats.indexedInSession += emailBatch.length;
    const elapsedSeconds = (Date.now() - currentRunStats.startTime) / 1000;
    const speed = Math.round(currentRunStats.indexedInSession / (elapsedSeconds || 1));

    const stats = state.getStats();
    state.updateStats({
      filesIndexed: stats.filesIndexed + emailBatch.length,
      filesIndexedThisSession: currentRunStats.indexedInSession,
      speed
    }, { immediate: true });

    state.addLog(`Batch uploaded. Session total: ${currentRunStats.indexedInSession} (${speed} emails/s)`);
  } catch (err) {
    state.addLog(`Meilisearch upload failed: ${err.message}`);
    state.addErrorLog('Meilisearch Batch Upload', err.message);
    const stats = state.getStats();
    state.updateStats({ filesFailed: stats.filesFailed + emailBatch.length }, { immediate: true });
    throw err;
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

function start(targetPaths = []) {
  const status = state.getIndexingStatus();
  if (status === 'scanning' || status === 'uploading') {
    state.addLog('Indexer is already running.');
    return;
  }

  state.addLog(`Starting indexer (${CONCURRENCY} workers, batch size ${BATCH_SIZE})...`);
  activeIndexerPromise = runIndexing(targetPaths)
    .catch(err => {
      state.addLog(`Critical indexer error: ${err.message}`);
      state.updateIndexingStatus('idle');
    })
    .finally(() => {
      activeIndexerPromise = null;
    });
}

function pause() {
  const status = state.getIndexingStatus();
  if (status === 'scanning' || status === 'uploading') {
    state.addLog('Pausing indexer...');
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
  state.addLog('Live Scheduler activated. Delta-sync will run every 15 minutes.');

  if (currentState.indexingStatus !== 'scanning' && currentState.indexingStatus !== 'uploading') {
    start();
  }

  schedulerInterval = setInterval(() => {
    const st = state.loadState();
    if (st.indexingStatus !== 'scanning' && st.indexingStatus !== 'uploading') {
      state.addLog('Live Scheduler: starting periodic delta-sync...');
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
