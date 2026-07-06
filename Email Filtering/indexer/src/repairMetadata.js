/**
 * Lightweight metadata repair for already-indexed emails.
 * Re-reads only To/Cc/Date/Sender from disk and updates Meilisearch —
 * no full re-scan and no body re-indexing.
 *
 * Designed to stay responsive: parsing runs in worker threads and the
 * main loop yields to the event loop frequently.
 */
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });
const fs = require('fs');
const { MeiliSearch } = require('meilisearch');
const { parseInWorker } = require('./parsePool');

const MEILI_BATCH_SIZE = 100;
const PAGE_SIZE = 500;
const YIELD_EVERY = 20;
const PROGRESS_EVERY = 25;
const LOG_EVERY = 500;

const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

function yieldToEventLoop() {
  return new Promise(resolve => setImmediate(resolve));
}

function needsRepair(doc) {
  const recipients = doc.recipients;
  const hasRecipients = Array.isArray(recipients)
    ? recipients.length > 0
    : !!(recipients && String(recipients).trim());
  const hasDate = doc.sentAt && Number(doc.sentAt) > 0;
  return !hasRecipients || !hasDate;
}

function buildPatch(doc, parsed) {
  const patch = { id: doc.id };
  let changed = false;

  const currentRecipients = Array.isArray(doc.recipients)
    ? doc.recipients.join(', ')
    : (doc.recipients || '');
  if (!currentRecipients.trim() && parsed.recipients) {
    patch.recipients = parsed.recipients;
    changed = true;
  }

  const currentCc = Array.isArray(doc.cc) ? doc.cc.join(', ') : (doc.cc || '');
  if (!currentCc.trim() && parsed.cc) {
    patch.cc = parsed.cc;
    changed = true;
  }

  if ((!doc.sentAt || Number(doc.sentAt) === 0) && parsed.sentAt > 0) {
    patch.sentAt = parsed.sentAt;
    changed = true;
  }

  if ((!doc.sender || doc.sender === 'Legacy Email') && parsed.sender) {
    patch.sender = parsed.sender;
    changed = true;
  }

  return changed ? patch : null;
}

async function repairDocument(doc) {
  if (!doc.filePath) return null;

  try {
    await fs.promises.access(doc.filePath);
  } catch {
    return null;
  }

  try {
    const parsed = await parseInWorker(doc.filePath);
    return buildPatch(doc, parsed);
  } catch (err) {
    return null;
  }
}

async function runRepair(options = {}) {
  const log = options.log || console.log;
  const onProgress = options.onProgress || (() => {});
  const shouldStop = options.shouldStop || (() => false);

  log('Starting metadata repair (To / Cc / Date only)...');
  log('The app will stay responsive — parsing runs in background workers.');
  log(`Meilisearch: ${process.env.MEILI_URL || 'http://localhost:7700'}`);

  const countResponse = await emailIndex.search('', { limit: 1 });
  const totalEstimate = countResponse.estimatedTotalHits || 0;

  let offset = 0;
  let scanned = 0;
  let repaired = 0;
  let skipped = 0;
  let failed = 0;
  let pendingBatch = [];
  let lastLogRepaired = 0;

  onProgress({
    total: totalEstimate,
    scanned,
    repaired,
    skipped,
    currentFilePath: '',
  });

  async function flushBatch() {
    if (pendingBatch.length === 0) return;
    const batch = pendingBatch;
    pendingBatch = [];
    try {
      await emailIndex.updateDocuments(batch);
      repaired += batch.length;
      if (repaired - lastLogRepaired >= LOG_EVERY) {
        log(`Metadata repair: ${repaired} emails updated (${scanned}/${totalEstimate} checked)...`);
        lastLogRepaired = repaired;
      }
    } catch (err) {
      failed += batch.length;
      log(`Metadata repair batch failed: ${err.message}`);
    }
  }

  while (true) {
    if (shouldStop()) {
      log('Metadata repair stopped by user.');
      break;
    }

    const response = await emailIndex.getDocuments({
      limit: PAGE_SIZE,
      offset,
      fields: ['id', 'filePath', 'recipients', 'cc', 'sentAt', 'sender'],
    });

    if (response.results.length === 0) break;

    for (const doc of response.results) {
      if (shouldStop()) break;

      scanned++;

      if (scanned % YIELD_EVERY === 0) {
        await yieldToEventLoop();
      }

      if (scanned % PROGRESS_EVERY === 0) {
        onProgress({
          total: totalEstimate,
          scanned,
          repaired,
          skipped,
          currentFilePath: doc.filePath || '',
        });
      }

      if (!needsRepair(doc)) {
        skipped++;
        continue;
      }

      const patch = await repairDocument(doc);
      if (patch) {
        pendingBatch.push(patch);
        if (pendingBatch.length >= MEILI_BATCH_SIZE) {
          await flushBatch();
          await yieldToEventLoop();
        }
      }
    }

    if (shouldStop()) break;
    if (response.results.length < PAGE_SIZE) break;
    offset += PAGE_SIZE;
    await yieldToEventLoop();
  }

  await flushBatch();

  onProgress({
    total: totalEstimate,
    scanned,
    repaired,
    skipped,
    currentFilePath: '',
  });

  const summary = `Metadata repair complete. Scanned: ${scanned}, Repaired: ${repaired}, Skipped: ${skipped}, Failed: ${failed}`;
  log(summary);

  return { scanned, repaired, skipped, failed, stopped: shouldStop() };
}

if (require.main === module) {
  runRepair().catch(err => {
    console.error('Fatal error during metadata repair:', err);
    process.exit(1);
  });
}

module.exports = { runRepair, needsRepair };
