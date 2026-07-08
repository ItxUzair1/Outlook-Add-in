const path = require('path');
const crypto = require('crypto');
const state = require('./state');
const { parseRobustEmailFile } = require('./robustParser');

const { MeiliSearch } = require('meilisearch');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

function sanitizeSurrogates(str) {
  if (typeof str !== 'string') return str;
  if (str.toWellFormed) return str.toWellFormed();
  return str.replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|([^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]/g, '$1\uFFFD');
}

const RETRY_BATCH_SIZE = 250;
const YIELD_EVERY_N = 25;

function yieldToEventLoop() {
  return new Promise(resolve => setImmediate(resolve));
}

/**
 * Orchestrator loop for retrying indexing of failed emails.
 */
async function runRetryErrors() {
  state.updateIndexingStatus('retrying');
  state.addLog('Starting fallback recovery for unparseable error emails...');

  const errorFiles = state.getUnparseableFiles();
  if (errorFiles.length === 0) {
    state.addLog('No error emails found in ledger to retry. Done.');
    state.updateIndexingStatus('idle');
    return;
  }

  state.addLog(`Found ${errorFiles.length} emails in the error log. Running robust recovery...`);

  state.updateStats({
    totalFilesFound: errorFiles.length,
    filesIndexedThisSession: 0,
    filesSkipped: 0,
    currentFilePath: 'Starting recovery...',
  }, { immediate: true });

  const folders = state.getFolders();
  let indexedCount = 0;
  let batch = [];
  let batchFilePaths = [];

  for (let i = 0; i < errorFiles.length; i++) {
    const filePath = errorFiles[i];

    const status = state.getIndexingStatus();
    if (status === 'paused' || status === 'idle') {
      state.addLog('Recovery job paused/stopped by user.');
      state.updateIndexingStatus('idle');
      return;
    }

    state.updateStats({ currentFilePath: filePath }, { persist: false });

    if (i > 0 && i % YIELD_EVERY_N === 0) {
      await yieldToEventLoop();
    }

    try {
      const folder = folders.find(f => filePath.toLowerCase().startsWith(f.path.toLowerCase())) || {};
      const parsed = await parseRobustEmailFile(filePath);

      batch.push({
        id: crypto.createHash('sha256').update(filePath).digest('hex'),
        subject: sanitizeSurrogates(parsed.subject),
        sender: sanitizeSurrogates(parsed.sender),
        recipients: sanitizeSurrogates(parsed.recipients),
        cc: sanitizeSurrogates(parsed.cc),
        bcc: sanitizeSurrogates(parsed.bcc),
        sentAt: parsed.sentAt,
        body: sanitizeSurrogates(parsed.body),
        hasAttachments: parsed.hasAttachments,
        filePath: parsed.filePath,
        comment: sanitizeSurrogates(parsed.comment),
        indexedRootPath: folder.path || '',
        indexedRootType: folder.type || 'local',
        collectionId: folder.type === 'collection'
          ? (folder.description || folder.collectionId)
          : (folder.collectionId || null),
        isPublic: folder.isPublic !== false,
        allowedUsers: (folder.allowedUsers || []).map(u => u.toLowerCase()),
      });
      batchFilePaths.push(filePath);
    } catch (err) {
      console.error(`[Recovery] Unexpected parsing crash on ${filePath}:`, err.message);
    }

    if (batch.length >= RETRY_BATCH_SIZE) {
      await flushRetryBatch(batch, batchFilePaths);
      indexedCount += batch.length;
      batch = [];
      batchFilePaths = [];
    }
  }

  if (batch.length > 0) {
    await flushRetryBatch(batch, batchFilePaths);
    indexedCount += batch.length;
  }

  state.addLog(`Recovery complete. Successfully indexed ${indexedCount} of ${errorFiles.length} previously failed emails.`);
  state.updateIndexingStatus('idle');
  state.updateStats({ currentFilePath: '', speed: 0 }, { immediate: true });
}

async function flushRetryBatch(documentsBatch, pathsBatch) {
  try {
    state.addLog(`Recovery: Uploading batch of ${documentsBatch.length} recovered emails to Meilisearch...`);
    await emailIndex.addDocuments(documentsBatch, { primaryKey: 'id' });

    for (const filePath of pathsBatch) {
      state.markFileUploaded(filePath);
      state.removeFileUnparseable(filePath);
    }

    const stats = state.getStats();
    const newFilesFailed = Math.max(0, state.getUnparseableFiles().length);
    state.updateStats({
      filesIndexed: stats.filesIndexed + documentsBatch.length,
      filesFailed: newFilesFailed,
      filesIndexedThisSession: stats.filesIndexedThisSession + documentsBatch.length,
    }, { immediate: true });
  } catch (err) {
    state.addLog(`Recovery upload batch failed: ${err.message}`);
    state.addErrorLog('Recovery Batch Upload', err.message);
  }
}

module.exports = {
  runRetryErrors,
};
