const { Worker } = require('worker_threads');
const path = require('path');
const os = require('os');
const { parseEmailFile } = require('./parser');

// ── Tuning constants ──────────────────────────────────────────────────────────
// Use up to 4 workers but never more than (cpuCount - 1) so the host stays
// responsive. On a 2-core machine this gives 1 worker; on 8+ cores it gives 4.
const WORKER_COUNT = Math.max(1, Math.min(os.cpus().length - 1, 4));
const PARSE_TIMEOUT_MS = 15 * 1000; // 15 s — generous but not wasteful
// ─────────────────────────────────────────────────────────────────────────────

let jobId = 0;
const pending = new Map(); // jobId → { resolve, reject, timeoutId, workerId }

// Pool state
const workers = []; // array of { worker, failed, busy }

function createWorkerSlot(index) {
  const slot = { worker: null, failed: false, index };
  spawnWorker(slot);
  return slot;
}

function spawnWorker(slot) {
  try {
    const w = new Worker(path.join(__dirname, 'parseWorker.js'));

    w.on('message', ({ id, result, error }) => {
      const job = pending.get(id);
      if (!job) return;
      pending.delete(id);
      if (job.timeoutId) clearTimeout(job.timeoutId);
      slot.busy = false;
      if (error) job.reject(new Error(error));
      else job.resolve(result);
    });

    w.on('error', (err) => {
      console.error(`[ParsePool] Worker ${slot.index} error: ${err.message}`);
      slot.failed = true;
      rejectSlotJobs(slot);
      // Recycle the slot after a short delay
      setTimeout(() => {
        slot.failed = false;
        spawnWorker(slot);
      }, 500);
    });

    w.on('exit', (code) => {
      if (code !== 0) {
        slot.failed = true;
        rejectSlotJobs(slot);
        setTimeout(() => {
          slot.failed = false;
          spawnWorker(slot);
        }, 500);
      }
    });

    slot.worker = w;
    slot.failed = false;
    slot.busy = false;
  } catch (err) {
    console.error(`[ParsePool] Failed to spawn worker ${slot.index}:`, err.message);
    slot.failed = true;
  }
}

function rejectSlotJobs(slot) {
  for (const [id, job] of pending) {
    if (job.workerId === slot.index) {
      pending.delete(id);
      if (job.timeoutId) clearTimeout(job.timeoutId);
      job.reject(new Error(`Parse worker ${slot.index} crashed`));
    }
  }
}

// Initialise pool
for (let i = 0; i < WORKER_COUNT; i++) {
  workers.push(createWorkerSlot(i));
}

// ── Round-robin index ─────────────────────────────────────────────────────────
let rrIndex = 0;

function pickWorker() {
  // Try round-robin first; fall back to any healthy worker
  for (let attempt = 0; attempt < workers.length; attempt++) {
    const slot = workers[rrIndex % workers.length];
    rrIndex++;
    if (!slot.failed && slot.worker) return slot;
  }
  return null; // all workers failed → fall back to main thread
}

function parseWithTimeout(parsePromise, filePath, timeoutMs) {
  return new Promise((resolve, reject) => {
    const timeoutId = setTimeout(() => {
      reject(new Error(`Parse timed out after ${Math.round(timeoutMs / 1000)}s (file skipped): ${path.basename(filePath)}`));
    }, timeoutMs);

    parsePromise
      .then(result => { clearTimeout(timeoutId); resolve(result); })
      .catch(err  => { clearTimeout(timeoutId); reject(err); });
  });
}

function parseInWorker(filePath) {
  const slot = pickWorker();

  if (!slot) {
    // All workers unavailable — parse on main thread with timeout
    return parseWithTimeout(parseEmailFile(filePath), filePath, PARSE_TIMEOUT_MS);
  }

  return new Promise((resolve, reject) => {
    const id = ++jobId;

    const timeoutId = setTimeout(() => {
      if (!pending.has(id)) return;
      pending.delete(id);
      console.error(`[ParsePool] Timeout on worker ${slot.index} for: ${path.basename(filePath)}`);
      // Recycle the timed-out worker
      if (slot.worker) {
        slot.worker.terminate().catch(() => {});
        slot.worker = null;
      }
      slot.busy = false;
      setTimeout(() => spawnWorker(slot), 100);
      reject(new Error(`Parse timed out after ${Math.round(PARSE_TIMEOUT_MS / 1000)}s (file skipped): ${path.basename(filePath)}`));
    }, PARSE_TIMEOUT_MS);

    pending.set(id, { resolve, reject, timeoutId, workerId: slot.index });
    slot.busy = true;
    slot.worker.postMessage({ id, filePath });
  });
}

function terminateWorker() {
  for (const slot of workers) {
    if (slot.worker) {
      slot.worker.terminate().catch(() => {});
      slot.worker = null;
    }
  }
  for (const [, { reject, timeoutId }] of pending) {
    if (timeoutId) clearTimeout(timeoutId);
    reject(new Error('Parse pool shut down'));
  }
  pending.clear();
}

module.exports = {
  parseInWorker,
  terminateWorker,
  PARSE_TIMEOUT_MS,
  WORKER_COUNT
};
