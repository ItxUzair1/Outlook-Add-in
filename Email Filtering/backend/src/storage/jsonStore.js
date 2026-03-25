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

export async function writeJson(filePath, data) {
  await ensureJsonFile(filePath, data);
  await fs.writeFile(filePath, JSON.stringify(data, null, 2), "utf-8");
}
