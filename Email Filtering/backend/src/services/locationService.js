import { v4 as uuidv4 } from "uuid";
import fs from "fs/promises";
import path from "path";
import { exec } from "child_process";
import { promisify } from "util";
import { 
  getLocations, 
  saveLocations, 
  getSenderFavouritesStore,
  saveSenderFavouritesStore,
  getSenderHistoryStore,
  saveSenderHistoryStore
} from "../storage/repositories.js";
import { readJson } from "../storage/jsonStore.js";
import { loadCollectionFile, saveCollectionFile } from "./collectionService.js";
import { config } from "../config/index.js";
import os from "os";
import { Meilisearch } from "meilisearch";

const execAsync = promisify(exec);

const prefsPath = path.join(config.dataDir, "preferences.json");

// Initialize Meilisearch client
const meiliClient = new Meilisearch({
  host: process.env.MEILI_URL || 'http://127.0.0.1:7700',
  apiKey: process.env.MEILI_MASTER_KEY
});
const emailIndex = meiliClient.index('emails');

async function resolveCollectionLocation(id) {
  if (!id || !id.startsWith("col_")) return null;

  const prefs = await readJson(prefsPath, {});
  const loadedCollections = prefs.loadedCollections || [];

  for (const filePath of loadedCollections) {
    const colName = path.basename(filePath.replace(/\\/g, "/"), ".mmcollection");
    const prefix = `col_${colName}_`;
    if (id.startsWith(prefix)) {
      const targetOriginalId = id.substring(prefix.length);
      const colLocs = await loadCollectionFile(filePath);
      const idx = colLocs.findIndex((loc, index) => {
        const originalId = loc.id || index;
        return String(originalId) === String(targetOriginalId);
      });
      if (idx !== -1) {
        return { filePath, colLocs, index: idx };
      }
    }
  }
  return null;
}

export async function exploreLocation(targetPath) {
  if (process.platform === "win32") {
    const timestamp = Date.now();
    const vbsPath = path.join(os.tmpdir(), `koyoexplore_${timestamp}.vbs`);
    const helperPath = path.join(os.tmpdir(), `koyoexpfocus_${timestamp}.vbs`);

    // If no path is selected, open default Explorer (This PC / Quick Access)
    // The window title is usually "File Explorer" or "This PC" or "Quick access".
    // We can try to activate "File Explorer" first, then fallback.
    const folderName = targetPath ? (path.basename(targetPath) || targetPath) : "File Explorer";

    // Focus-helper: waits for Explorer window to appear, then forces it to foreground
    const focusHelperScript = [
      'WScript.Sleep 800',
      'Set ws = CreateObject("WScript.Shell")',
      'ws.SendKeys "%"',
      'WScript.Sleep 100',
      `ws.AppActivate "${folderName.replace(/"/g, '""')}"`,
      // Fallback for default explorer if "File Explorer" isn't the exact title
      ...(!targetPath ? [
          'WScript.Sleep 100',
          'ws.AppActivate "This PC"',
          'WScript.Sleep 100',
          'ws.AppActivate "Quick access"',
          'WScript.Sleep 100',
          'ws.AppActivate "Home"'
      ] : [])
    ].join("\r\n");

    // Main script: launches focus helper, then opens the folder
    const mainScript = [
      'Set fso = CreateObject("Scripting.FileSystemObject")',
      'Set WshShell = CreateObject("WScript.Shell")',
      '',
      `WshShell.Run "wscript ""${helperPath.replace(/\\/g, "\\\\").replace(/"/g, '""')}""", 0, False`,
      '',
      'WScript.Sleep 100',
      'WshShell.SendKeys "%"',
      'Set shell = CreateObject("Shell.Application")',
      targetPath 
        ? `shell.Open "${targetPath.replace(/"/g, '""')}"` 
        : `WshShell.Run "explorer.exe", 1, False`,
      '',
      'On Error Resume Next',
      `fso.DeleteFile "${helperPath.replace(/\\/g, "\\\\")}", True`,
      'On Error Goto 0',
    ].join("\r\n");

    try {
      await fs.writeFile(helperPath, focusHelperScript);
      await fs.writeFile(vbsPath, mainScript);
      await execAsync(`cscript //nologo "${vbsPath}"`);
      try { await fs.unlink(vbsPath); } catch (e) {}
      try { await fs.unlink(helperPath); } catch (e) {}
    } catch (err) {
      console.warn("Explore location warning:", err.message);
    }
  } else {
    try {
      if (!targetPath) {
        await execAsync(`open .`);
      } else {
        await execAsync(`open "${targetPath}"`);
      }
    } catch (err) {
      console.warn("Explore location warning:", err.message);
    }
  }
}

export async function getSenderHistoryStats(sender) {
  if (!sender || !sender.trim()) {
    return {};
  }
  try {
    const cleanSender = sender.trim().toLowerCase();
    const store = await getSenderHistoryStore();
    if (!store[cleanSender]) return {};
    
    // Convert array format to expected dictionary format
    const folderStats = {};
    for (const item of store[cleanSender]) {
      const dir = item.path.replace(/\\/g, "/").toLowerCase();
      folderStats[dir] = { 
         count: item.usageCount || 1, 
         lastUsed: new Date(item.lastUsedAt || 0).getTime() 
      };
    }
    return folderStats;
  } catch (err) {
    console.warn("[locationService] Failed to get sender history stats:", err.message);
    return {};
  }
}

export async function getGeneralHistoryStats() {
  try {
    const searchResponse = await emailIndex.search('', {
      limit: 2000,
      sort: ['sentAt:desc'],
      attributesToRetrieve: ['filePath', 'filedAt', 'sentAt']
    });

    const folderStats = {};
    for (const item of searchResponse.hits) {
      if (!item.filePath) continue;

      const dir = path.dirname(item.filePath).replace(/\\/g, "/").toLowerCase();
      if (!folderStats[dir]) {
        folderStats[dir] = { count: 0, lastUsed: 0 };
      }
      folderStats[dir].count += 1;

      const useTime = new Date(item.filedAt || item.sentAt || 0).getTime();
      if (useTime > folderStats[dir].lastUsed) {
        folderStats[dir].lastUsed = useTime;
      }
    }
    return folderStats;
  } catch (err) {
    console.warn("[locationService] Failed to get general history stats:", err.message);
    return {};
  }
}

export async function listLocations(sender) {
  let locations = await getLocations();

  // ── Passive self-heal: remove any "Discovered" entries whose path is now covered
  // by a loaded .mmcollection file. This fixes the overnight regression where Discover
  // runs before the collection is loaded and writes stale Discovered entries to the DB.
  // We do this quietly on every GET /api/locations to keep the database clean.
  try {
    const prefs = await readJson(prefsPath, {});
    const loadedCollections = prefs.loadedCollections || [];
    if (loadedCollections.length > 0) {
      const collectionCoveredPaths = new Set();
      for (const filePath of loadedCollections) {
        try {
          const colLocs = await loadCollectionFile(filePath);
          if (Array.isArray(colLocs)) {
            for (const loc of colLocs) {
              const p = loc.folder || loc.path;
              if (p) collectionCoveredPaths.add(p.toLowerCase().replace(/\\/g, "/"));
            }
          }
        } catch (err) { /* ignore unreadable collection files */ }
      }
      if (collectionCoveredPaths.size > 0) {
        const stale = locations.filter(loc =>
          String(loc.collection || "").toLowerCase() === "discovered" &&
          collectionCoveredPaths.has((loc.path || "").toLowerCase().replace(/\\/g, "/"))
        );
        if (stale.length > 0) {
          const staleIds = new Set(stale.map(l => l.id));
          locations = locations.filter(l => !staleIds.has(l.id));
          await saveLocations(locations);
          console.log(`[locationService] listLocations: auto-removed ${stale.length} stale Discovered entries`);
        }
      }
    }
  } catch (err) {
    console.warn("[locationService] listLocations: stale Discovered cleanup failed (non-fatal):", err.message);
  }

  try {
    const [folderStats, generalStats, favourites] = await Promise.all([
      getSenderHistoryStats(sender),
      getGeneralHistoryStats(),
      getSenderFavourites(sender)
    ]);

    const normalizePath = (p) => {
      if (!p) return "";
      return p.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
    };

    const normalizedFavourites = (favourites || []).map(p => normalizePath(p));

    return [...locations].sort((a, b) => {
      // 1. Keep unused folders at the bottom
      if (a.isUnused && !b.isUnused) return 1;
      if (!a.isUnused && b.isUnused) return -1;
      if (a.isUnused && b.isUnused) return 0;

      // 2. Prioritize starred/favourites for this sender (or general if sender is empty)
      const normA = normalizePath(a.path);
      const normB = normalizePath(b.path);

      const isFavA = sender ? normalizedFavourites.includes(normA) : !!a.isSuggested;
      const isFavB = sender ? normalizedFavourites.includes(normB) : !!b.isSuggested;

      if (isFavA && !isFavB) return -1;
      if (!isFavA && isFavB) return 1;
      if (isFavA && isFavB) {
        // Both starred: sort by sender frequency, then recency, then general usage, then alphabetical description
        const statA = folderStats[normA];
        const statB = folderStats[normB];
        if (statA && !statB) return -1;
        if (!statA && statB) return 1;
        if (statA && statB) {
          if (statA.count !== statB.count) return statB.count - statA.count;
          if (statA.lastUsed !== statB.lastUsed) return statB.lastUsed - statA.lastUsed;
        }
        const genA = generalStats[normA];
        const genB = generalStats[normB];
        if (genA && !genB) return -1;
        if (!genA && genB) return 1;
        if (genA && genB) {
          if (genA.count !== genB.count) return genB.count - genA.count;
          if (genA.lastUsed !== genB.lastUsed) return genB.lastUsed - genA.lastUsed;
        }
        return String(a.description || "").localeCompare(String(b.description || ""));
      }

      // 3. Prioritize matching sender history details
      const statA = folderStats[normA];
      const statB = folderStats[normB];

      if (statA && !statB) return -1;
      if (!statA && statB) return 1;
      if (statA && statB) {
        if (statA.count !== statB.count) {
          return statB.count - statA.count;
        }
        if (statA.lastUsed !== statB.lastUsed) {
          return statB.lastUsed - statA.lastUsed;
        }
      }

      // 4. Prioritize general history details
      const genA = generalStats[normA];
      const genB = generalStats[normB];

      if (genA && !genB) return -1;
      if (!genA && genB) return 1;
      if (genA && genB) {
        if (genA.count !== genB.count) {
          return genB.count - genA.count;
        }
        if (genA.lastUsed !== genB.lastUsed) {
          return genB.lastUsed - genA.lastUsed;
        }
      }

      // 5. General lastUsedAt from locations.json (backup)
      const timeA = a.lastUsedAt ? new Date(a.lastUsedAt).getTime() : 0;
      const timeB = b.lastUsedAt ? new Date(b.lastUsedAt).getTime() : 0;
      if (timeA !== timeB) {
        return timeB - timeA;
      }

      // 6. Alphabetical description
      return String(a.description || "").localeCompare(String(b.description || ""));
    });
  } catch (err) {
    console.warn("[locationService] Failed to sort locations:", err.message);
    // Simple default sort fallback
    return [...locations].sort((a, b) => {
      if (a.isUnused && !b.isUnused) return 1;
      if (!a.isUnused && b.isUnused) return -1;
      const timeA = a.lastUsedAt ? new Date(a.lastUsedAt).getTime() : 0;
      const timeB = b.lastUsedAt ? new Date(b.lastUsedAt).getTime() : 0;
      return timeB - timeA;
    });
  }
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

export async function getSenderFavourites(sender) {
  if (!sender || !sender.trim()) {
    return [];
  }
  const store = await getSenderFavouritesStore();
  return store[sender.trim().toLowerCase()] || [];
}

export async function removeSuggestion(id, sender) {
  if (sender && sender.trim()) {
    const cleanSender = sender.trim().toLowerCase();
    let folderPath = null;
    let locationData = null;

    if (id && id.startsWith("col_")) {
      const resolved = await resolveCollectionLocation(id);
      if (resolved) {
        const { filePath, colLocs, index } = resolved;
        const loc = colLocs[index];
        folderPath = loc.folder || loc.path;
        locationData = {
          ...loc,
          id,
          path: folderPath,
          collection: path.basename(filePath.replace(/\\/g, "/"), ".mmcollection")
        };
      }
    } else {
      const data = await getLocations();
      const idx = data.findIndex((x) => x.id === id);
      if (idx >= 0) {
        folderPath = data[idx].path;
        locationData = data[idx];
      }
    }

    if (!folderPath || !locationData) {
      return null;
    }

    const store = await getSenderFavouritesStore();
    if (store[cleanSender]) {
      const normPath = folderPath.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
      store[cleanSender] = store[cleanSender].filter(p => p.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim() !== normPath);
      await saveSenderFavouritesStore(store);
    }
    return locationData;
  }

  // Fallback to legacy behaviour if no sender is provided
  if (id && id.startsWith("col_")) {
    const resolved = await resolveCollectionLocation(id);
    if (!resolved) return null;
    const { filePath, colLocs, index } = resolved;
    colLocs[index].isSuggested = false;
    await saveCollectionFile(filePath, colLocs);
    return {
      ...colLocs[index],
      id,
      path: colLocs[index].folder || colLocs[index].path,
      collection: path.basename(filePath.replace(/\\/g, "/"), ".mmcollection")
    };
  }

  const data = await getLocations();
  const idx = data.findIndex((x) => x.id === id);
  if (idx < 0) return null;

  data[idx] = {
    ...data[idx],
    isSuggested: false,
    lastUsedAt: null,
    updatedAt: new Date().toISOString(),
  };

  await saveLocations(data);
  return data[idx];
}

export async function toggleSuggestion(id, sender) {
  if (sender && sender.trim()) {
    const cleanSender = sender.trim().toLowerCase();
    let folderPath = null;
    let locationData = null;

    if (id && id.startsWith("col_")) {
      const resolved = await resolveCollectionLocation(id);
      if (resolved) {
        const { filePath, colLocs, index } = resolved;
        const loc = colLocs[index];
        folderPath = loc.folder || loc.path;
        locationData = {
          ...loc,
          id,
          path: folderPath,
          collection: path.basename(filePath.replace(/\\/g, "/"), ".mmcollection")
        };
      }
    } else {
      const data = await getLocations();
      const idx = data.findIndex((x) => x.id === id);
      if (idx >= 0) {
        folderPath = data[idx].path;
        locationData = data[idx];
      }
    }

    if (!folderPath || !locationData) {
      return null;
    }

    const store = await getSenderFavouritesStore();
    if (!store[cleanSender]) {
      store[cleanSender] = [];
    }

    const normPath = folderPath.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
    const index = store[cleanSender].findIndex(p => p.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim() === normPath);

    if (index >= 0) {
      // Remove it
      store[cleanSender].splice(index, 1);
    } else {
      // Add it
      store[cleanSender].push(folderPath);
    }

    await saveSenderFavouritesStore(store);
    return locationData;
  }

  // Fallback to legacy behaviour if no sender is provided
  if (id && id.startsWith("col_")) {
    const resolved = await resolveCollectionLocation(id);
    if (!resolved) return null;
    const { filePath, colLocs, index } = resolved;
    colLocs[index].isSuggested = !colLocs[index].isSuggested;
    await saveCollectionFile(filePath, colLocs);
    return {
      ...colLocs[index],
      id,
      path: colLocs[index].folder || colLocs[index].path,
      collection: path.basename(filePath.replace(/\\/g, "/"), ".mmcollection")
    };
  }

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

export async function markUnused(id) {
  if (id && id.startsWith("col_")) {
    const resolved = await resolveCollectionLocation(id);
    if (!resolved) return null;
    const { filePath, colLocs, index } = resolved;
    colLocs[index].isUnused = !colLocs[index].isUnused;
    await saveCollectionFile(filePath, colLocs);
    return {
      ...colLocs[index],
      id,
      path: colLocs[index].folder || colLocs[index].path,
      collection: path.basename(filePath.replace(/\\/g, "/"), ".mmcollection")
    };
  }

  const data = await getLocations();
  const idx = data.findIndex((x) => x.id === id);
  if (idx < 0) return null;

  const isCurrentlyUnused = !!data[idx].isUnused;
  
  data[idx] = {
    ...data[idx],
    isUnused: !isCurrentlyUnused,
    lastUsedAt: null,
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
    collection: payload.collection || "Portfolio",
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
  if (id && id.startsWith("col_")) {
    const resolved = await resolveCollectionLocation(id);
    if (!resolved) return null;
    const { filePath, colLocs, index } = resolved;
    const updatedLoc = {
      ...colLocs[index],
      ...payload
    };
    if (payload.path) {
      updatedLoc.folder = payload.path;
    }
    colLocs[index] = updatedLoc;
    await saveCollectionFile(filePath, colLocs);
    return {
      ...updatedLoc,
      id,
      path: updatedLoc.folder || updatedLoc.path,
      collection: path.basename(filePath.replace(/\\/g, "/"), ".mmcollection")
    };
  }

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
  if (id && id.startsWith("col_")) {
    const resolved = await resolveCollectionLocation(id);
    if (!resolved) return false;
    const { filePath, colLocs, index } = resolved;
    colLocs.splice(index, 1);
    await saveCollectionFile(filePath, colLocs);
    return true;
  }

  const data = await getLocations();
  const filtered = data.filter((x) => x.id !== id);
  const removed = filtered.length !== data.length;

  if (removed) {
    await saveLocations(filtered);
  }

  return removed;
}

export async function markUsedByPaths(targetPaths, sender = null) {
  const data = await getLocations();
  const now = new Date().toISOString();
  let changed = false;

  const normalize = (p) => {
    if (!p) return "";
    return p.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
  };

  const normalizedTargets = (targetPaths || []).map(normalize);

  const updated = data.map((x) => {
    if (normalizedTargets.includes(normalize(x.path))) {
      changed = true;
      return { ...x, lastUsedAt: now, updatedAt: now };
    }

    return x;
  });

  if (changed) {
    await saveLocations(updated);
  }

  // Update sender history locally
  if (sender && sender.trim() && targetPaths && targetPaths.length > 0) {
    const cleanSender = sender.trim().toLowerCase();
    const historyStore = await getSenderHistoryStore();
    if (!historyStore[cleanSender]) {
      historyStore[cleanSender] = [];
    }

    let historyChanged = false;
    for (const targetPath of targetPaths) {
      const normPath = normalize(targetPath);
      const existing = historyStore[cleanSender].find(x => normalize(x.path) === normPath);
      if (existing) {
        existing.usageCount = (existing.usageCount || 0) + 1;
        existing.lastUsedAt = now;
        historyChanged = true;
      } else {
        historyStore[cleanSender].push({
          path: targetPath,
          usageCount: 1,
          lastUsedAt: now
        });
        historyChanged = true;
      }
    }

    if (historyChanged) {
      await saveSenderHistoryStore(historyStore);
    }
  }
}


async function mapWithConcurrency(array, fn, limit) {
  const results = [];
  const executing = new Set();
  for (let i = 0; i < array.length; i++) {
    const item = array[i];
    const index = i;
    const p = Promise.resolve().then(() => fn(item, index)).then((res) => {
      results[index] = res;
      executing.delete(p);
    });
    executing.add(p);
    if (executing.size >= limit) {
      await Promise.race(executing);
    }
  }
  await Promise.all(executing);
  return results;
}

async function isConnected(filePath) {
  try {
    // Race fs.access against a 3-second timeout so unreachable UNC paths
    // do not block the entire connectivity check.
    await Promise.race([
      fs.access(filePath),
      new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), 3000))
    ]);
    return true;
  } catch {
    return false;
  }
}

export async function checkConnectivity() {
  const data = await getLocations();
  // Limit concurrency to 4 to prevent threadpool starvation
  const entries = await mapWithConcurrency(
    data,
    async (item) => [item.id, await isConnected(item.path)],
    4
  );
  return Object.fromEntries(entries);
}

export async function checkPathsConnectivity(paths) {
  // Limit concurrency to 4 to prevent threadpool starvation
  const entries = await mapWithConcurrency(
    paths,
    async (p) => [p.id, await isConnected(p.path)],
    4
  );
  return Object.fromEntries(entries);
}

/**
 * Scans the search index for unique filing directories and auto-creates
 * location entries for any directories not already registered.
 * @returns {{ addedCount: number, totalScanned: number }}
 */
export async function discoverLocations() {
  const existingLocations = await getLocations();
  const existingPaths = new Set(existingLocations.map(loc => (loc.path || "").toLowerCase().replace(/\\/g, "/")));

  // ── Build a set of all paths that are already covered by a loaded collection file ──
  // We must NEVER create a "Discovered" entry for a folder that lives in a .mmcollection
  // file — otherwise those entries shadow the real collection entries and cause the
  // "Personal → Discovered" regression that appears after leaving the PC on overnight.
  const collectionCoveredPaths = new Set();
  try {
    const prefs = await readJson(prefsPath, {});
    const loadedCollections = prefs.loadedCollections || [];
    for (const filePath of loadedCollections) {
      try {
        const colLocs = await loadCollectionFile(filePath);
        if (Array.isArray(colLocs)) {
          for (const loc of colLocs) {
            const p = loc.folder || loc.path;
            if (p) {
              collectionCoveredPaths.add(p.toLowerCase().replace(/\\/g, "/"));
            }
          }
        }
      } catch (err) {
        // Ignore unreadable collection files — treat as uncovered
      }
    }
  } catch (err) {
    console.warn("[locationService] discoverLocations: failed to read preferences for collection paths:", err.message);
  }

  // ── Remove any stale "Discovered" entries that are now covered by a collection file ──
  // This self-heals the database for PCs that were left on overnight and had Discovered
  // entries written before the collection file finished loading.
  const staleDiscovered = existingLocations.filter(loc => {
    if (String(loc.collection || "").toLowerCase() !== "discovered") return false;
    const normPath = (loc.path || "").toLowerCase().replace(/\\/g, "/");
    return collectionCoveredPaths.has(normPath);
  });
  if (staleDiscovered.length > 0) {
    const staleIds = new Set(staleDiscovered.map(l => l.id));
    const cleaned = existingLocations.filter(l => !staleIds.has(l.id));
    await saveLocations(cleaned);
    // Re-sync existingPaths after cleanup
    existingPaths.clear();
    cleaned.forEach(loc => existingPaths.add((loc.path || "").toLowerCase().replace(/\\/g, "/")));
    console.log(`[locationService] discoverLocations: removed ${staleDiscovered.length} stale Discovered entries covered by collection files`);
  }

  const discoveredDirs = new Set();
  let totalScanned = 0;

  try {
    const searchResponse = await emailIndex.search('', {
      limit: 5000,
      attributesToRetrieve: ['filePath']
    });

    totalScanned = searchResponse.hits.length;

    for (const item of searchResponse.hits) {
      if (!item.filePath) continue;
      const dir = path.dirname(item.filePath);
      if (dir) {
        discoveredDirs.add(dir);
      }
    }
  } catch (err) {
    console.warn("[locationService] Failed to fetch paths for discovery:", err.message);
  }

  // Filter out directories that already exist as locations OR are covered by a collection file
  const newDirs = [];
  for (const dir of discoveredDirs) {
    const normalized = dir.toLowerCase().replace(/\\/g, "/");
    if (!existingPaths.has(normalized) && !collectionCoveredPaths.has(normalized)) {
      newDirs.push(dir);
    }
  }

  // Create location entries for each new directory
  const now = new Date().toISOString();
  const newLocations = newDirs.map(dir => ({
    id: uuidv4(),
    type: dir.startsWith("\\\\") ? "network" : "local",
    path: dir,
    description: path.basename(dir) || dir,
    collection: "Discovered",
    isDefault: false,
    isSuggested: false,
    createdAt: now,
    updatedAt: now,
    lastUsedAt: null,
  }));

  if (newLocations.length > 0) {
    const currentLocations = await getLocations();
    const allLocations = [...currentLocations, ...newLocations];
    await saveLocations(allLocations);
  }

  return { addedCount: newLocations.length, totalScanned };
}
