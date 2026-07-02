const { Worker } = require('worker_threads');
const path = require('path');
const os = require('os');

const WORKER_COUNT = Math.max(1, Math.min(os.cpus().length - 1, 4));
const PARSE_TIMEOUT_MS = 15 * 1000;

// In-memory skip set for this process (normalized paths)
const crashSkipSet = new Set();

let jobId = 0;
const pending = new Map(); // jobId → { resolve, reject, timeoutId, workerId, filePath }

const workers = [];

function normalizePath(filePath) {
  return path.normalize(filePath).toLowerCase();
}

function persistUnparseable(filePath) {
  try {
    const state = require('./state');
    state.markFileUnparseable(filePath);
  } catch (err) {
    console.error('[ParsePool] Failed to persist unparseable file:', err.message);
  }
}

function createWorkerSlot(index) {
  const slot = { worker: null, failed: false, busy: false, index, crashCount: 0, handlingCrash: false };
  spawnWorker(slot);
  return slot;
}

function scheduleSlotRestart(slot) {
  if (slot.handlingCrash) return;
  slot.handlingCrash = true;
  slot.failed = true;

  if (slot.worker) {
    slot.worker.terminate().catch(() => {});
    slot.worker = null;
  }

  rejectSlotJobs(slot, true);

  const delay = Math.min(500 * Math.pow(2, slot.crashCount || 0), 10000);
  slot.crashCount = (slot.crashCount || 0) + 1;

  setTimeout(() => {
    slot.handlingCrash = false;
    slot.failed = false;
    spawnWorker(slot);
  }, delay);
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
      slot.crashCount = 0;

      if (error) {
        const normalized = job.filePath ? normalizePath(job.filePath) : null;
        if (normalized) crashSkipSet.add(normalized);
        job.reject(new Error(error));
      } else {
        job.resolve(result);
      }
    });

    w.on('error', (err) => {
      console.error(`[ParsePool] Worker ${slot.index} error: ${err.message}`);
      scheduleSlotRestart(slot);
    });

    w.on('exit', (code) => {
      if (code !== 0) {
        console.error(`[ParsePool] Worker ${slot.index} exited with code ${code}`);
        scheduleSlotRestart(slot);
      } else {
        slot.crashCount = 0;
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

function rejectSlotJobs(slot, hardCrash = false) {
  for (const [id, job] of pending) {
    if (job.workerId !== slot.index) continue;
    pending.delete(id);
    if (job.timeoutId) clearTimeout(job.timeoutId);
    slot.busy = false;

    if (hardCrash && job.filePath) {
      const normalized = normalizePath(job.filePath);
      crashSkipSet.add(normalized);
      persistUnparseable(job.filePath);
      console.warn(`[ParsePool] Worker ${slot.index} hard-crashed on: ${path.basename(job.filePath)} — permanently skipped`);
    }

    job.reject(new Error(`Parse worker ${slot.index} crashed`));
  }
}

for (let i = 0; i < WORKER_COUNT; i++) {
  workers.push(createWorkerSlot(i));
}

let rrIndex = 0;

function pickWorker() {
  for (let attempt = 0; attempt < workers.length; attempt++) {
    const slot = workers[rrIndex % workers.length];
    rrIndex++;
    if (!slot.failed && !slot.handlingCrash && slot.worker) return slot;
  }
  return null;
}

async function parseInWorker(filePath) {
  const normalized = normalizePath(filePath);

  if (crashSkipSet.has(normalized)) {
    throw new Error(`Skipped (caused worker crash previously): ${path.basename(filePath)}`);
  }

  try {
    const state = require('./state');
    if (state.isFileUnparseable(filePath)) {
      throw new Error(`Skipped (unparseable — worker crash/timeout previously): ${path.basename(filePath)}`);
    }
  } catch (err) {
    if (err.message.includes('Skipped (unparseable')) throw err;
  }

  let slot = pickWorker();

  while (!slot) {
    await new Promise(resolve => setTimeout(resolve, 200));
    slot = pickWorker();
  }

  return new Promise((resolve, reject) => {
    const id = ++jobId;

    const timeoutId = setTimeout(() => {
      if (!pending.has(id)) return;
      pending.delete(id);
      console.error(`[ParsePool] Timeout on worker ${slot.index} for: ${path.basename(filePath)}`);
      crashSkipSet.add(normalized);
      persistUnparseable(filePath);

      if (slot.worker) {
        slot.worker.terminate().catch(() => {});
        slot.worker = null;
      }
      slot.busy = false;
      setTimeout(() => spawnWorker(slot), 100);

      reject(new Error(`Parse timed out after ${Math.round(PARSE_TIMEOUT_MS / 1000)}s (file skipped): ${path.basename(filePath)}`));
    }, PARSE_TIMEOUT_MS);

    pending.set(id, { resolve, reject, timeoutId, workerId: slot.index, filePath });
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
