import fs from "fs/promises";
import path from "path";

async function ensureJsonFile(filePath, seed) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  try {
    await fs.access(filePath);
  } catch {
    await fs.writeFile(filePath, JSON.stringify(seed, null, 2), "utf-8");
  }
}

export async function readJson(filePath, seed) {
  await ensureJsonFile(filePath, seed);
  const raw = await fs.readFile(filePath, "utf-8");
  return JSON.parse(raw);
}

export async function writeJson(filePath, data, { compact = false } = {}) {
  await ensureJsonFile(filePath, data);
  const json = compact ? JSON.stringify(data) : JSON.stringify(data, null, 2);
  await fs.writeFile(filePath, json, "utf-8");
}
