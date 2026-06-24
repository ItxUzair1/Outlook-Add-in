import { Router } from "express";
import { getSearchIndex, saveSearchIndex, getLocations } from "../../storage/repositories.js";
import { exec } from "child_process";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { config } from "../../config/index.js";
import { readJson } from "../../storage/jsonStore.js";
import { loadCollectionFile } from "../../services/collectionService.js";
import MsgReaderPkg from "@kenjiuno/msgreader";

const MsgReader = MsgReaderPkg.default || MsgReaderPkg;

const router = Router();

// ── Search & Crawler Helper Functions ────────────────────────────────────────

// Helper to decode RFC 2047 header values
function decodeRFC2047(str) {
  if (!str) return "";
  return str.replace(/=\?([^?]+)\?([QB])\?([^?]*)\?=/gi, (match, charset, encoding, text) => {
    if (encoding.toUpperCase() === "B") {
      try {
        return Buffer.from(text, "base64").toString(charset.toLowerCase() === "utf-8" ? "utf8" : "binary");
      } catch (err) {
        return text;
      }
    } else if (encoding.toUpperCase() === "Q") {
      const decoded = text.replace(/_/g, " ").replace(/=([0-9A-F]{2})/gi, (m, hex) => {
        return String.fromCharCode(parseInt(hex, 16));
      });
      try {
        return Buffer.from(decoded, "binary").toString(charset.toLowerCase() === "utf-8" ? "utf8" : "binary");
      } catch (err) {
        return decoded;
      }
    }
    return match;
  });
}

// Helper to parse EML headers
async function parseEmlHeader(filePath) {
  let fileHandle;
  try {
    fileHandle = await fs.open(filePath, "r");
    const buffer = Buffer.alloc(16384);
    const { bytesRead } = await fileHandle.read(buffer, 0, 16384, 0);
    const content = buffer.toString("utf8", 0, bytesRead);
    
    const headerEndIndex = content.search(/\r?\n\r?\n/);
    const headerText = headerEndIndex !== -1 ? content.slice(0, headerEndIndex) : content;
    const unfoldedText = headerText.replace(/\r?\n[ \t]+/g, " ");

    const lines = unfoldedText.split(/\r?\n/);
    const headers = {};
    for (const line of lines) {
      const colonIndex = line.indexOf(":");
      if (colonIndex !== -1) {
        const key = line.slice(0, colonIndex).trim().toLowerCase();
        const value = line.slice(colonIndex + 1).trim();
        headers[key] = value;
      }
    }

    const subject = decodeRFC2047(headers.subject || "");
    const sender = decodeRFC2047(headers.from || "");
    const toStr = decodeRFC2047(headers.to || "");
    const ccStr = decodeRFC2047(headers.cc || "");
    const dateStr = headers.date || "";

    // Split address list on commas that are OUTSIDE angle brackets
    // so that display names like "Smith, John <john@firm.com>" are not broken
    const splitAddresses = (str) => {
      if (!str) return [];
      const addrs = [];
      let depth = 0;
      let current = "";
      for (const ch of str) {
        if (ch === "<") depth++;
        else if (ch === ">") depth--;
        if (ch === "," && depth === 0) {
          if (current.trim()) addrs.push(current.trim());
          current = "";
        } else {
          current += ch;
        }
      }
      if (current.trim()) addrs.push(current.trim());
      return addrs;
    };

    const recipients = splitAddresses(toStr);
    const cc = splitAddresses(ccStr);

    // Keep the full "Display Name <email>" form so the From column shows the person's name
    const cleanSender = sender.trim() || "Unknown Sender";

    return {
      subject: subject || path.basename(filePath, path.extname(filePath)),
      sender: cleanSender || "Unknown Sender",
      recipients,
      cc,
      sentAt: dateStr ? new Date(dateStr).toISOString() : null
    };
  } catch (err) {
    console.error(`Failed to parse EML headers for ${filePath}:`, err.message);
    return null;
  } finally {
    if (fileHandle) {
      await fileHandle.close();
    }
  }
}

// Helper to parse MSG file details using MsgReader with filename/stats fallback
async function parseMsgFile(filePath) {
  try {
    const fileBuffer = await fs.readFile(filePath);
    const reader = new MsgReader(fileBuffer);
    const info = reader.getFileData();

    // Subject
    const subject = info.subject || "";

    // Sender
    let sender = "";
    if (info.senderEmail) {
      sender = info.senderName ? `${info.senderName} <${info.senderEmail}>` : info.senderEmail;
    } else {
      sender = info.senderName || "";
    }

    // Recipients and CC
    const recipients = [];
    const cc = [];
    if (Array.isArray(info.recipients)) {
      for (const rec of info.recipients) {
        const addr = rec.emailAddress || rec.smtpAddress || "";
        const name = rec.name && rec.name !== addr ? rec.name : "";
        const full = addr ? (name ? `${name} <${addr}>` : addr) : rec.name || "";
        if (full) {
          if (rec.recipType === "to" || rec.recipientType === "to") {
            recipients.push(full);
          } else if (rec.recipType === "cc" || rec.recipientType === "cc") {
            cc.push(full);
          } else {
            recipients.push(full);
          }
        }
      }
    }

    // Sent Date
    let sentAt = null;
    if (info.clientSubmitTime) {
      try {
        sentAt = new Date(info.clientSubmitTime).toISOString();
      } catch (e) {}
    }
    if (!sentAt && info.messageDeliveryTime) {
      try {
        sentAt = new Date(info.messageDeliveryTime).toISOString();
      } catch (e) {}
    }
    if (!sentAt) {
      try {
        const stat = await fs.stat(filePath);
        sentAt = stat.mtime.toISOString();
      } catch (e) {}
    }

    return {
      subject: subject || path.basename(filePath, path.extname(filePath)),
      sender: sender || "Unknown Sender",
      recipients,
      cc: cc,
      sentAt
    };
  } catch (err) {
    console.warn(`[searchRoutes] Failed to parse MSG file ${filePath} with MsgReader:`, err.message);
    // Safe fallback: parse from file name and filesystem stats
    try {
      const stat = await fs.stat(filePath);
      const baseName = path.basename(filePath, path.extname(filePath));
      
      let subject = baseName;
      let sentAt = stat.mtime.toISOString();
      
      const datePrefixMatch = baseName.match(/^(\d{8})_(\d{6})_(.*)$/);
      if (datePrefixMatch) {
        const [_, yyyymmdd, hhmmss, rest] = datePrefixMatch;
        subject = rest;
        try {
          const year = yyyymmdd.slice(0, 4);
          const month = yyyymmdd.slice(4, 6);
          const day = yyyymmdd.slice(6, 8);
          const hour = hhmmss.slice(0, 2);
          const min = hhmmss.slice(2, 4);
          const sec = hhmmss.slice(4, 6);
          sentAt = new Date(`${year}-${month}-${day}T${hour}:${min}:${sec}.000Z`).toISOString();
        } catch (e) {}
      } else {
        const datePrefixMatch2 = baseName.match(/^(\d{8})_(.*)$/);
        if (datePrefixMatch2) {
          const [_, yyyymmdd, rest] = datePrefixMatch2;
          subject = rest;
          try {
            const year = yyyymmdd.slice(0, 4);
            const month = yyyymmdd.slice(4, 6);
            const day = yyyymmdd.slice(6, 8);
            sentAt = new Date(`${year}-${month}-${day}T00:00:00.000Z`).toISOString();
          } catch (e) {}
        }
      }
      
      return {
        subject: subject || baseName,
        sender: "Legacy Email",
        recipients: [],
        cc: [],
        sentAt
      };
    } catch (fallbackErr) {
      return null;
    }
  }
}

// Unified parser
async function parseEmailFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".eml") {
    const parsed = await parseEmlHeader(filePath);
    if (parsed) return parsed;
  }
  return parseMsgFile(filePath);
}

// Fast directory scanner
async function scanDirectory(dirPath, maxDepth = 5, currentDepth = 0) {
  const files = [];
  try {
    const entries = await fs.readdir(dirPath, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dirPath, entry.name);
      if (entry.isFile()) {
        const ext = path.extname(entry.name).toLowerCase();
        if (ext === ".eml" || ext === ".msg") {
          files.push(fullPath);
        }
      } else if (entry.isDirectory() && currentDepth < maxDepth) {
        if (entry.name.startsWith(".") || entry.name.toLowerCase() === "node_modules") {
          continue;
        }
        const subFiles = await scanDirectory(fullPath, maxDepth, currentDepth + 1);
        files.push(...subFiles);
      }
    }
  } catch (err) {
    // Ignore error
  }
  return files;
}


/**
 * Returns the list of directories to scan based on the given searchScope.
 * Centralises scope→directory resolution so both search and sync use the same logic.
 */
async function getScopedDirectories(searchScope) {
  const dirs = [];
  const resolvedScope = searchScope || "locations_i_use";

  if (resolvedScope === "personal_only" || resolvedScope === "locations_i_use" || resolvedScope === "all_locations") {
    const locations = await getLocations();
    dirs.push(...locations.map(loc => loc.path).filter(Boolean));
  }

  if (resolvedScope === "locations_i_use" || resolvedScope === "all_locations") {
    // Also include collection paths
    try {
      const prefsPath = path.join(config.dataDir, "preferences.json");
      const prefs = await readJson(prefsPath, {});
      if (prefs.loadedCollections && Array.isArray(prefs.loadedCollections)) {
        for (const filePath of prefs.loadedCollections) {
          try {
            const colLocs = await loadCollectionFile(filePath);
            if (Array.isArray(colLocs)) {
              for (const loc of colLocs) {
                const p = loc.folder || loc.path;
                if (p) dirs.push(p);
              }
            }
          } catch (err) {}
        }
      }
    } catch (err) {}
  } else if (resolvedScope.startsWith("collection:")) {
    const colPath = resolvedScope.replace("collection:", "");
    try {
      const colLocs = await loadCollectionFile(colPath);
      if (Array.isArray(colLocs)) {
        for (const loc of colLocs) {
          const p = loc.folder || loc.path;
          if (p) dirs.push(p);
        }
      }
    } catch (err) {}
  }

  return [...new Set(dirs.filter(Boolean))];
}


import { MeiliSearch } from 'meilisearch';

const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

async function getIndexerState() {
  try {
    const statePath = path.join(process.cwd(), '..', 'indexer', 'data', 'indexer_state.json');
    const content = await fs.readFile(statePath, 'utf8');
    return JSON.parse(content);
  } catch (err) {
    return { folders: [] };
  }
}

router.get("/api/active-collections", async (req, res) => {
  try {
    const state = await getIndexerState();
    const collectionIds = [...new Set(state.folders.map(f => f.collectionId).filter(Boolean))];
    res.json({ collections: collectionIds });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.get("/", async (req, res, next) => {
  try {
    const { keywords = "", searchScope } = req.query;
    
    let meiliFilters = [];
    
    const resolvedScope = searchScope || "locations_i_use";
    
    if (resolvedScope === "personal_only") {
      meiliFilters.push('indexedRootType = "local"');
    } else if (resolvedScope.startsWith("collection:")) {
      const colPath = resolvedScope.replace("collection:", "").replace(/\\/g, '\\\\');
      meiliFilters.push(`collectionId = "${colPath}"`);
    } else if (resolvedScope === "locations_i_use") {
      const state = await getIndexerState();
      const rootPaths = [...new Set(state.folders.map(f => f.path))];
      if (rootPaths.length > 0) {
        const inClause = rootPaths.map(p => `"${p.replace(/\\/g, '\\\\')}"`).join(', ');
        meiliFilters.push(`indexedRootPath IN [${inClause}]`);
      } else {
        return res.json({ results: [], estimatedTotalHits: 0 });
      }
    }
    
    const searchParams = { limit: 100 };
    if (meiliFilters.length > 0) {
      searchParams.filter = meiliFilters;
    }
    
    const searchResponse = await emailIndex.search(keywords, searchParams);
    res.json({ count: searchResponse.hits.length, results: searchResponse.hits, estimatedTotalHits: searchResponse.estimatedTotalHits });
  } catch (err) {
    console.error("Meilisearch search error:", err);
    res.status(500).json({ error: "Search failed" });
  }
});


/**
 * GET /api/search/browse-folder
 * Opens a folder picker using a native executable.
 * Accepts optional ?startPath=C:\path query param
 */
router.get("/browse-folder", async (req, res, next) => {
  const possiblePaths = [
    path.join(path.dirname(process.execPath), "koyobrowse.exe"),
    path.join(process.cwd(), "koyobrowse.exe"),
    path.join(process.cwd(), "bin", "koyobrowse.exe"),
    path.join(process.cwd(), "src", "bin", "koyobrowse.exe"),
  ];

  let exePath = null;
  for (const p of possiblePaths) {
    try {
      await fs.access(p);
      exePath = p;
      break;
    } catch (err) {}
  }

  if (!exePath) {
    return res.status(500).json({ error: "Folder picker utility (koyobrowse.exe) not found" });
  }

  let cmd = `"${exePath}" "Select Destination Folder"`;
  if (req.query.startPath) {
    cmd += ` "${req.query.startPath}"`;
  }

  exec(
    cmd,
    { timeout: 120000 },
    (error, stdout, stderr) => {
      if (error && error.killed) {
        return res.status(500).json({ error: "Folder picker timed out" });
      }
      if (error && !stdout.trim()) {
        return res.status(500).json({ error: "Failed to open folder picker", details: stderr || error.message });
      }
      const selectedPath = stdout.trim();
      res.json({ path: selectedPath });
    }
  );
});

/**
 * GET /api/search/browse-file
 * Opens a file picker using a native executable.
 */
router.get("/browse-file", async (req, res, next) => {
  const possiblePaths = [
    path.join(path.dirname(process.execPath), "koyofile.exe"),
    path.join(process.cwd(), "koyofile.exe"),
    path.join(process.cwd(), "bin", "koyofile.exe"),
    path.join(process.cwd(), "src", "bin", "koyofile.exe"),
  ];

  let exePath = null;
  for (const p of possiblePaths) {
    try {
      await fs.access(p);
      exePath = p;
      break;
    } catch (err) {}
  }

  if (!exePath) {
    return res.status(500).json({ error: "File picker utility (koyofile.exe) not found" });
  }

  exec(
    `"${exePath}"`,
    { timeout: 120000 },
    (error, stdout, stderr) => {
      if (error && error.killed) {
        return res.status(500).json({ error: "File picker timed out" });
      }
      if (error && !stdout.trim()) {
        return res.status(500).json({ error: "Failed to open file picker", details: stderr || error.message });
      }
      const selectedPath = stdout.trim();
      res.json({ path: selectedPath });
    }
  );
});

/**
 * POST /api/search/open
 * Opens the file in its default OS application (Outlook).
 */
router.post("/open", async (req, res, next) => {
  try {
    const { filePath } = req.body;
    if (!filePath) return res.status(400).json({ error: "filePath is required" });

    // Verify file exists first
    try {
      await fs.access(filePath);
    } catch (err) {
      console.warn(`[searchRoutes] Open attempt failed: File not found at ${filePath}`);
      return res.status(404).json({ error: "File not found at original location", code: "ENOENT" });
    }

    // Use 'start' command on Windows to launch default app
    // Surround with double quotes to handle spaces in paths
    exec(`start "" "${filePath}"`, (error) => {
      if (error) {
          console.error(`[searchRoutes] Failed to open file: ${error.message}`);
          return res.status(500).json({ error: `Could not open file: ${error.message}` });
      }
      res.json({ status: "success" });
    });
  } catch (e) {
    next(e);
  }
});

/**
 * GET /api/search/open-local
 * Opens the file in its default OS application via a GET request (useful for email hyperlinks).
 * Example: https://localhost:4000/api/search/open-local?path=C:/foo/bar.pdf
 */
router.get("/open-local", async (req, res, next) => {
  try {
    const filePath = req.query.path;
    if (!filePath) return res.status(400).send("File path is required.");

    // Verify file exists
    try {
      await fs.access(filePath);
    } catch (err) {
      console.warn(`[searchRoutes] Open attempt failed: File not found at ${filePath}`);
      return res.status(404).send(`File not found at: ${filePath}`);
    }

    // Use 'start' command on Windows
    exec(`start "" "${filePath}"`, (error) => {
      if (error) {
          console.error(`[searchRoutes] Failed to open file: ${error.message}`);
          return res.status(500).send(`Could not open file: ${error.message}`);
      }
      
      // Send a self-closing HTML page with a fallback message
      res.send(`
        <html>
          <head>
            <title>Opening File...</title>
            <style>
              body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; text-align: center; margin-top: 50px; color: #333; }
              .container { max-width: 400px; margin: 0 auto; padding: 20px; border: 1px solid #ccc; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
              h2 { color: #0078D4; }
            </style>
          </head>
          <body>
            <div class="container">
              <h2>File Opened Successfully</h2>
              <p>Your file has been opened in Windows.</p>
              <p style="color: #666; font-size: 14px;">You can now safely close this tab.</p>
            </div>
            <script>
              // Attempt to close the window automatically
              try {
                window.open('', '_self', '');
                window.close();
              } catch (e) {}
              
              setTimeout(function() {
                 window.close();
              }, 500);
            </script>
          </body>
        </html>
      `);
    });
  } catch (e) {
    next(e);
  }
});

/**
 * POST /api/search/open-folder
 * Opens the containing folder of the search result.
 */
router.post("/open-folder", async (req, res, next) => {
  try {
    const { filePath } = req.body;
    if (!filePath) return res.status(400).json({ error: "filePath is required" });

    // Extract directory from file path
    const dirPath = path.dirname(filePath);

    // Verify directory exists first
    try {
      await fs.access(dirPath);
    } catch (err) {
      console.warn(`[searchRoutes] Open Folder attempt failed: Directory not found at ${dirPath}`);
      return res.status(404).json({ error: "Folder not found at original location", code: "ENOENT" });
    }

    // Use 'start' command to launch Explorer at that directory
    exec(`start "" "${dirPath}"`, (error) => {
      if (error) {
          console.error(`[searchRoutes] Failed to open folder: ${error.message}`);
          return res.status(500).json({ error: `Could not open folder: ${error.message}` });
      }
      res.json({ status: "success" });
    });
  } catch (e) {
    next(e);
  }
});

/**
 * DELETE /api/search/:id
 * Deletes the physical file and removes its entry from the index.
 */
router.delete("/:id", async (req, res, next) => {
  try {
    const { id } = req.params;
    const index = await getSearchIndex();
    const itemIdx = index.findIndex(x => x.id === id);

    if (itemIdx === -1) {
        return res.status(404).json({ error: "Item not found in index" });
    }

    const item = index[itemIdx];

    // Delete the file from the filesystem
    try {
      if (item.filePath) {
          await fs.unlink(item.filePath);
      }
    } catch (err) {
      console.warn(`[searchRoutes] Failed to delete physical file: ${err.message}`);
      // Continue even if physical file is missing, to keep index clean
    }

    // Update the search index
    const updatedIndex = index.filter(x => x.id !== id);
    await saveSearchIndex(updatedIndex);

    res.json({ status: "deleted" });
  } catch (e) {
    next(e);
  }
});

/**
 * POST /api/search/move
 * Moves the physical file to a new destination directory and updates the search index.
 */
router.post("/move", async (req, res, next) => {
  try {
    const { id, destinationDir } = req.body;
    if (!id || !destinationDir) {
      return res.status(400).json({ error: "id and destinationDir are required" });
    }

    const index = await getSearchIndex();
    const itemIdx = index.findIndex(x => x.id === id);

    if (itemIdx === -1) {
      return res.status(404).json({ error: "Item not found in index" });
    }

    const item = index[itemIdx];
    if (!item.filePath) {
      return res.status(400).json({ error: "Item does not have a physical file path" });
    }

    // Verify source exists
    try {
      await fs.access(item.filePath);
    } catch {
      return res.status(404).json({ error: "Original file not found on disk" });
    }

    // Verify destination directory exists
    try {
      const destStat = await fs.stat(destinationDir);
      if (!destStat.isDirectory()) {
         return res.status(400).json({ error: "Destination must be a directory" });
      }
    } catch {
      return res.status(400).json({ error: "Destination directory does not exist or is inaccessible" });
    }

    // Move file
    const fileName = path.basename(item.filePath);
    const newFilePath = path.join(destinationDir, fileName);

    // Prevent overwriting
    try {
      await fs.access(newFilePath);
      return res.status(400).json({ error: "A file with that name already exists at the destination" });
    } catch {
      // Good, file doesn't exist
    }

    try {
      await fs.rename(item.filePath, newFilePath);
    } catch (renameErr) {
      if (renameErr.code === 'EXDEV') {
        // Cross-device link not permitted, use copy + unlink fallback
        await fs.copyFile(item.filePath, newFilePath);
        await fs.unlink(item.filePath);
      } else {
        throw renameErr;
      }
    }

    // Update index
    item.filePath = newFilePath;
    item.parentFolder = destinationDir; // Update parent folder attribute if tracked
    await saveSearchIndex(index);

    res.json({ status: "moved", newPath: newFilePath });
  } catch (e) {
    next(e);
  }
});

/**
 * POST /api/search/sync
 * Scans the entire index and removes entries where the physical file is missing.
 */
router.post("/sync", async (req, res, next) => {
  try {
    const index = await getSearchIndex();
    const { filePaths, searchScope } = req.body || {};

    let newFilesToScan = [];
    let prunedIndex = [...index];
    let removedCount = 0;

    const legacySenderValues = new Set(["Legacy Email", "Legacy Email File (Unindexed)", "Unknown Sender", ""]);

    if (Array.isArray(filePaths) && filePaths.length > 0) {
      // Focused Sync: Index only the specified files (e.g. legacy search results)
      // We only consider it "already indexed and rich" if it's in the index and does not need repair.
      const indexedRichPaths = new Set(
        index
          .filter(item => item.filePath && item.sender && !legacySenderValues.has(item.sender))
          .map(item => (item.filePath || "").toLowerCase().replace(/\\/g, "/"))
      );
      newFilesToScan = filePaths.filter(fp => fp && !indexedRichPaths.has(fp.toLowerCase().replace(/\\/g, "/")));
    } else {
      // Global Sync: Scan directories and prune missing files
      // Use getScopedDirectories to respect the dropdown scope
      const uniqueDirs = await getScopedDirectories(searchScope || "locations_i_use");

      // 2. Scan directories for files on disk
      let filesOnDisk = [];
      if (uniqueDirs.length > 0) {
        const scanPromises = uniqueDirs.map(d => scanDirectory(d, 5));
        const scanResults = await Promise.all(scanPromises);
        filesOnDisk = scanResults.flat();
      }

      // 3. Prune entries in the index that no longer exist on disk (parallel with batch concurrency)
      prunedIndex = [];
      const accessBatchSize = 200;
      for (let i = 0; i < index.length; i += accessBatchSize) {
        const batch = index.slice(i, i + accessBatchSize);
        const results = await Promise.all(batch.map(async (item) => {
          if (!item.filePath) return null;
          try {
            await fs.access(item.filePath);
            return item;
          } catch (err) {
            return null;
          }
        }));
        prunedIndex.push(...results.filter(Boolean));
      }
      removedCount = index.length - prunedIndex.length;

      // 3b. Find already-indexed legacy entries that are missing sender/recipient data
      //     so they can be re-parsed by the new MsgReader parser to populate From/To.
      const toRepair = prunedIndex.filter(item =>
        item.filePath &&
        (!item.sender || legacySenderValues.has(item.sender)) &&
        !item.msgReaderAttempted
      );

      // Identify brand new files on disk (files on disk that are not in the pruned index)
      const indexedPaths = new Set(prunedIndex.map(item => (item.filePath || "").toLowerCase().replace(/\\/g, "/")));
      const brandNewFiles = filesOnDisk.filter(fp => !indexedPaths.has(fp.toLowerCase().replace(/\\/g, "/")));

      // Repair paths that we want to scan (we will keep them in prunedIndex for now, but also scan/parse them)
      const repairFilePaths = toRepair.map(i => i.filePath).filter(Boolean);
      newFilesToScan = [...brandNewFiles, ...repairFilePaths];
    }

    // 5. Cap legacy file parsing to avoid timeouts
    const MAX_LEGACY_INDEX_PER_RUN = 2000;
    const filesToParse = newFilesToScan.slice(0, MAX_LEGACY_INDEX_PER_RUN);

    // Remove only the files we are about to parse/re-parse from prunedIndex, so we don't have duplicates
    // and so we don't lose unrepaired files from the index while they wait for future batches.
    const parsedPathsSet = new Set(filesToParse.map(fp => fp.toLowerCase().replace(/\\/g, "/")));
    prunedIndex = prunedIndex.filter(item =>
      !item.filePath || !parsedPathsSet.has(item.filePath.toLowerCase().replace(/\\/g, "/"))
    );

    // 6. Concurrently parse new files (batch size of 25)
    const newRows = [];
    const BATCH_SIZE = 50;
    const filedAt = new Date().toLocaleString("en-US", { timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone });
    
    for (let i = 0; i < filesToParse.length; i += BATCH_SIZE) {
      const batch = filesToParse.slice(i, i + BATCH_SIZE);
      const batchResults = await Promise.all(batch.map(async (fp) => {
        try {
          const parsed = await parseEmailFile(fp);
          if (parsed) {
            const stat = await fs.stat(fp);
            return {
              id: `indexed-${parsed.internetMessageId || parsed.subject}-${fp}-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
              internetMessageId: parsed.internetMessageId || null,
              subject: parsed.subject || path.basename(fp, path.extname(fp)),
              sender: parsed.sender || "Legacy Email",
              recipients: parsed.recipients || [],
              cc: parsed.cc || [],
              sentAt: parsed.sentAt || stat.mtime.toISOString(),
              filedAt: filedAt,
              hasAttachments: false,
              filePath: fp,
              comment: "Legacy email found via folder sync",
              body: "",
              isLegacyIndexed: true,
              msgReaderAttempted: true
            };
          }
        } catch (err) {
          console.warn(`[searchRoutes] Sync: Failed to parse legacy file ${fp}:`, err.message);
        }
        return null;
      }));
      newRows.push(...batchResults.filter(Boolean));
    }

    // 7. De-duplicate final index by filePath (ensuring unique file paths in database)
    const seen = new Set();
    const updatedIndex = [];
    const combinedIndex = [...prunedIndex, ...newRows]; // Prioritize existing rich database entries first

    for (const item of combinedIndex) {
      if (!item.filePath) continue;
      const key = item.filePath.toLowerCase().replace(/\\/g, "/");
      if (seen.has(key)) continue;
      seen.add(key);
      updatedIndex.push(item);
    }

    // 8. Save the search index
    await saveSearchIndex(updatedIndex);

    res.json({
      status: "synced",
      removedCount: removedCount,
      addedCount: newRows.length,
      totalCount: updatedIndex.length,
      hasMore: newFilesToScan.length > MAX_LEGACY_INDEX_PER_RUN
    });
  } catch (e) {
    next(e);
  }
});

export default router;
