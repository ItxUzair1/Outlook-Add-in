import fs from "fs/promises";
import path from "path";

const locks = {};

async function acquireLock(filePath) {
  if (!locks[filePath]) {
    locks[filePath] = Promise.resolve();
  }
  
  let release;
  const nextLock = new Promise((resolve) => {
    release = resolve;
  });
  
  const currentLock = locks[filePath];
  locks[filePath] = nextLock;
  
  await currentLock;
  return release;
}

async function ensureJsonFile(filePath, seed) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  try {
    await fs.access(filePath);
  } catch {
    await fs.writeFile(filePath, JSON.stringify(seed, null, 2), "utf-8");
  }
}

export async function readJson(filePath, seed) {
  const release = await acquireLock(filePath);
  try {
    await ensureJsonFile(filePath, seed);
    const raw = await fs.readFile(filePath, "utf-8");
    if (!raw.trim()) {
      return seed;
    }
    return JSON.parse(raw);
  } catch (err) {
    if (err instanceof SyntaxError) {
      console.warn(`[jsonStore] Corrupted JSON in ${filePath}, resetting to seed:`, err.message);
      return seed;
    }
    throw err;
  } finally {
    release();
  }
}

export async function writeJson(filePath, data, { compact = false } = {}) {
  const release = await acquireLock(filePath);
  try {
    await ensureJsonFile(filePath, data);
    const json = compact ? JSON.stringify(data) : JSON.stringify(data, null, 2);
    await fs.writeFile(filePath, json, "utf-8");
  } finally {
    release();
  }
}
