import { v4 as uuidv4 } from "uuid";
import fs from "fs/promises";
import path from "path";
import { exec } from "child_process";
import { promisify } from "util";
import { getLocations, saveLocations, getSearchIndex } from "../storage/repositories.js";

import os from "os";

const execAsync = promisify(exec);

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

export async function listLocations(sender) {
  const locations = await getLocations();

  const defaultSort = (a, b) => {
    // Keep unused folders at the bottom
    if (a.isUnused && !b.isUnused) return 1;
    if (!a.isUnused && b.isUnused) return -1;
    const timeA = a.lastUsedAt ? new Date(a.lastUsedAt).getTime() : 0;
    const timeB = b.lastUsedAt ? new Date(b.lastUsedAt).getTime() : 0;
    return timeB - timeA;
  };

  if (!sender || !sender.trim()) {
    return [...locations].sort(defaultSort);
  }

  try {
    const index = await getSearchIndex();
    const cleanSender = sender.trim().toLowerCase();

    // Filter index for entries where the sender matches
    const senderFilings = index.filter(item => 
      item.sender && item.sender.trim().toLowerCase() === cleanSender
    );

    if (senderFilings.length === 0) {
      return [...locations].sort(defaultSort);
    }

    // Group files by normalized parent directory, count frequency and track latest use
    const folderStats = {};
    for (const item of senderFilings) {
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

    const normalizePath = (p) => {
      if (!p) return "";
      return p.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
    };

    return [...locations].sort((a, b) => {
      // Keep unused folders at the bottom
      if (a.isUnused && !b.isUnused) return 1;
      if (!a.isUnused && b.isUnused) return -1;

      const normA = normalizePath(a.path);
      const normB = normalizePath(b.path);

      const statA = folderStats[normA];
      const statB = folderStats[normB];

      if (statA && !statB) return -1;
      if (!statA && statB) return 1;

      if (statA && statB) {
        // Both matched - sort by count descending, then by lastUsed descending
        if (statA.count !== statB.count) {
          return statB.count - statA.count;
        }
        return statB.lastUsed - statA.lastUsed;
      }

      // Neither matched - sort by general lastUsedAt descending
      const timeA = a.lastUsedAt ? new Date(a.lastUsedAt).getTime() : 0;
      const timeB = b.lastUsedAt ? new Date(b.lastUsedAt).getTime() : 0;
      return timeB - timeA;
    });
  } catch (err) {
    console.warn("[locationService] Failed to sort locations by sender:", err.message);
    return [...locations].sort(defaultSort);
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

export async function markUnused(id) {
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
  const index = await getSearchIndex();
  const existingLocations = await getLocations();
  const existingPaths = new Set(existingLocations.map(loc => (loc.path || "").toLowerCase().replace(/\\/g, "/")));

  // Collect unique parent directories from the search index
  const discoveredDirs = new Set();
  for (const item of index) {
    if (!item.filePath) continue;
    const dir = path.dirname(item.filePath);
    if (dir) {
      discoveredDirs.add(dir);
    }
  }

  // Filter out directories that already exist as locations
  const newDirs = [];
  for (const dir of discoveredDirs) {
    const normalized = dir.toLowerCase().replace(/\\/g, "/");
    if (!existingPaths.has(normalized)) {
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
    const allLocations = [...existingLocations, ...newLocations];
    await saveLocations(allLocations);
  }

  return { addedCount: newLocations.length, totalScanned: index.length };
}
