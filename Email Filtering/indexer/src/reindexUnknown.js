const fs = require('fs');
const path = require('path');
const { MeiliSearch } = require('meilisearch');
const {
  parseRobustEmailFile,
  needsReindex,
  buildReindexPatch,
} = require('./robustParser');

require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

const PAGE_SIZE = 500;
const BATCH_SIZE = 100;
const YIELD_EVERY = 25;
const SEARCH_PAGE_SIZE = 1000;

function yieldToEventLoop() {
  return new Promise(resolve => setImmediate(resolve));
}

function sanitizeSurrogates(str) {
  if (typeof str !== 'string') return str;
  if (str.toWellFormed) return str.toWellFormed();
  return str.replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|([^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]/g, '$1\uFFFD');
}

function sanitizePatch(patch) {
  const sanitized = { id: patch.id };
  for (const [key, value] of Object.entries(patch)) {
    if (key === 'id') continue;
    sanitized[key] = typeof value === 'string' ? sanitizeSurrogates(value) : value;
  }
  return sanitized;
}

async function fetchAllSearchHits(query) {
  const hitsById = new Map();
  let offset = 0;

  while (true) {
    const res = await emailIndex.search(query, {
      limit: SEARCH_PAGE_SIZE,
      offset,
      attributesToRetrieve: ['id', 'filePath', 'sender', 'recipients', 'body', 'sentAt', 'subject', 'cc', 'hasAttachments'],
    });

    for (const hit of res.hits) {
      hitsById.set(hit.id, hit);
    }

    if (res.hits.length < SEARCH_PAGE_SIZE) break;
    offset += SEARCH_PAGE_SIZE;
    await yieldToEventLoop();
  }

  return hitsById;
}

async function collectProblematicDocuments(log) {
  log('Collecting emails with Unknown Sender or Legacy Email labels...');
  const hitsById = await fetchAllSearchHits('"Unknown Sender"');
  
  const legacyHits = await fetchAllSearchHits('"Legacy Email"');
  for (const [id, doc] of legacyHits) {
    hitsById.set(id, doc);
  }
  log(`Found ${hitsById.size} emails from targeted search.`);

  log('Scanning index for emails with empty To or empty body (Ultra-fast scan)...');
  const SCAN_LIMIT = 5000;
  let offset = 0;
  let scanned = 0;

  while (true) {
    const response = await emailIndex.getDocuments({
      limit: SCAN_LIMIT,
      offset,
      fields: ['id', 'filePath', 'sender', 'recipients', 'body', 'sentAt', 'subject', 'cc', 'hasAttachments'],
    });

    if (response.results.length === 0) break;

    for (const doc of response.results) {
      scanned++;
      if (needsReindex(doc)) {
        hitsById.set(doc.id, doc);
      }
    }

    if (response.results.length < SCAN_LIMIT) break;
    offset += SCAN_LIMIT;

    log(`Scanned ${scanned} indexed emails, ${hitsById.size} need repair so far...`);
    await yieldToEventLoop();
  }

  log(`Scan complete. ${hitsById.size} emails flagged for re-parsing (${scanned} total checked).`);
  return [...hitsById.values()];
}

async function runReindexUnknown({ log = console.log, onProgress = () => {}, shouldStop = () => false }) {
  log('Starting repair for Unknown Sender, empty To, and empty body emails...');

  const documents = await collectProblematicDocuments(log);
  if (documents.length === 0) {
    log('No problematic emails found.');
    return { success: true, count: 0, scanned: 0, skipped: 0 };
  }

  let successCount = 0;
  let skippedCount = 0;
  let updates = [];

  for (let i = 0; i < documents.length; i++) {
    if (shouldStop()) {
      log('Re-index stopped by user.');
      break;
    }

    const doc = documents[i];

    if (i % YIELD_EVERY === 0) await yieldToEventLoop();

    onProgress({
      total: documents.length,
      scanned: i + 1,
      repaired: successCount,
      skipped: skippedCount,
      currentFilePath: doc.filePath,
    });

    if (!doc.filePath || !fs.existsSync(doc.filePath)) {
      skippedCount++;
      continue;
    }

    try {
      const parsed = await parseRobustEmailFile(doc.filePath);
      const patch = buildReindexPatch(doc, parsed);
      if (patch) {
        updates.push(sanitizePatch(patch));
        successCount++;
      } else {
        skippedCount++;
      }
    } catch {
      skippedCount++;
    }

    if (updates.length >= BATCH_SIZE) {
      log(`Sending batch of ${updates.length} updates to Meilisearch...`);
      await emailIndex.updateDocuments(updates);
      updates = [];
      await yieldToEventLoop();
    }
  }

  if (updates.length > 0) {
    await emailIndex.updateDocuments(updates);
  }

  log(`Finished! Repaired ${successCount} of ${documents.length} problematic emails (${skippedCount} unchanged or skipped).`);
  return {
    success: true,
    count: successCount,
    scanned: documents.length,
    skipped: skippedCount,
  };
}

module.exports = { runReindexUnknown, needsReindex, buildReindexPatch };
