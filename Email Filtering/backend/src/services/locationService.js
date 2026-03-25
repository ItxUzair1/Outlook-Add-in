import { v4 as uuidv4 } from "uuid";
import fs from "fs/promises";
import { exec } from "child_process";
import { promisify } from "util";
import { getLocations, saveLocations } from "../storage/repositories.js";

const execAsync = promisify(exec);

export async function exploreLocation(path) {
  // Use 'start' on Windows to open folder asynchronously and avoid false-positive exit code errors
  const command = process.platform === "win32" ? `start "" "${path}"` : `open "${path}"`;
  await execAsync(command);
}

export async function listLocations() {
  return getLocations();
}

export async function listSuggestedLocations(limit = 10) {
  const data = await getLocations();
  return data
    .filter((x) => x.isSuggested || Boolean(x.lastUsedAt))
    .sort((a, b) => {
      // Prioritize explicit isSuggested, then lastUsedAt
      if (a.isSuggested && !b.isSuggested) return -1;
      if (!a.isSuggested && b.isSuggested) return 1;
      return new Date(b.lastUsedAt || 0) - new Date(a.lastUsedAt || 0);
    })
    .slice(0, limit);
}

export async function removeSuggestion(id) {
  const data = await getLocations();
  const idx = data.findIndex((x) => x.id === id);
  if (idx < 0) return null;

  data[idx] = {
    ...data[idx],
    isSuggested: false,
    lastUsedAt: null, // Also clear lastUsedAt to stop it from being suggested by usage
    updatedAt: new Date().toISOString(),
  };

  await saveLocations(data);
  return data[idx];
}

export async function toggleSuggestion(id) {
  const data = await getLocations();
  const idx = data.findIndex((x) => x.id === id);
  if (idx < 0) return null;

  data[idx] = {
    ...data[idx],
    isSuggested: !data[idx].isSuggested,
    updatedAt: new Date().toISOString(),
  };

  await saveLocations(data);
  return data[idx];
}

export async function createLocation(payload) {
  const now = new Date().toISOString();
  const data = await getLocations();

  const item = {
    id: uuidv4(),
    type: payload.type || "network",
    path: payload.path,
    description: payload.description || "",
    collection: payload.collection || "Projects",
    isDefault: Boolean(payload.isDefault),
    isSuggested: payload.isSuggested || false,
    createdAt: now,
    updatedAt: now,
    lastUsedAt: null,
  };

  if (item.isDefault) {
    data.forEach((x) => {
      x.isDefault = false;
    });
  }

  data.push(item);
  await saveLocations(data);
  return item;
}

export async function updateLocation(id, payload) {
  const data = await getLocations();
  const idx = data.findIndex((x) => x.id === id);
  if (idx < 0) {
    return null;
  }

  if (payload.isDefault) {
    data.forEach((x) => {
      x.isDefault = false;
    });
  }

  data[idx] = {
    ...data[idx],
    ...payload,
    updatedAt: new Date().toISOString(),
  };

  await saveLocations(data);
  return data[idx];
}

export async function removeLocation(id) {
  const data = await getLocations();
  const filtered = data.filter((x) => x.id !== id);
  const removed = filtered.length !== data.length;

  if (removed) {
    await saveLocations(filtered);
  }

  return removed;
}

export async function markUsedByPaths(targetPaths) {
  const data = await getLocations();
  const now = new Date().toISOString();
  let changed = false;

  const updated = data.map((x) => {
    if (targetPaths.includes(x.path)) {
      changed = true;
      return { ...x, lastUsedAt: now, updatedAt: now };
    }

    return x;
  });

  if (changed) {
    await saveLocations(updated);
  }
}

async function isConnected(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

export async function checkConnectivity() {
  const data = await getLocations();
  const results = {};
  for (const item of data) {
    results[item.id] = await isConnected(item.path);
  }
  return results;
}
