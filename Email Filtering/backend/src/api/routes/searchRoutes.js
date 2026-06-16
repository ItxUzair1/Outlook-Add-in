import { Router } from "express";
import { getSearchIndex, saveSearchIndex, getLocations } from "../../storage/repositories.js";
import { exec } from "child_process";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { config } from "../../config/index.js";
import { readJson } from "../../storage/jsonStore.js";
import { loadCollectionFile } from "../../services/collectionService.js";

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

    const parseAddresses = (str) => {
      if (!str) return [];
      return str.split(",").map(addr => {
        const match = addr.match(/<([^>]+)>/);
        return (match ? match[1] : addr).trim();
      }).filter(Boolean);
    };

    const recipients = parseAddresses(toStr);
    const cc = parseAddresses(ccStr);
    
    const senderMatch = sender.match(/<([^>]+)>/);
    const cleanSender = senderMatch ? senderMatch[1] : sender.trim();

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

// Helper to parse MSG file details using file name and stats
async function parseMsgFile(filePath) {
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
  } catch (err) {
    return null;
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
async function scanDirectory(dirPath, maxDepth = 2, currentDepth = 0) {
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
 * GET /api/search?dateRange=&from=&to=&cc=&subject=&body=&hasAttachments=&location=&keywords=&resultKind=&searchScope=
 * resultKind: all (default) | files — files = index row whose filePath is not .eml/.msg (e.g. saved attachments).
 * searchScope: locations_i_use (default) | all_locations — restricts results to user's configured locations or searches all.
 * Searches the filed email index with optional filters.
 */
router.get("/", async (req, res, next) => {
  try {
    const index = await getSearchIndex();

    const {
      dateRange,   // e.g. "1m", "3m", "6m", "1y", "all"
      from,
      to,
      cc,
      subject,
      keywords,    // matches subject, sender, recipients, path, body; + comment if including=true
      location,    // filed location path keyword
      hasAttachments, // "true" / "false"
      body,        // search within indexed body
      resultKind,  // "all" | "files"
      searchScope, // "locations_i_use" | "all_locations"
    } = req.query;

    let results = [...index];

    // ── Search scope filter (locations I use vs all) ────────────────────────
    if (!searchScope || searchScope === "locations_i_use") {
      // Include user's configured locations
      const locations = await getLocations();
      const locationPaths = locations.map(loc => (loc.path || "").toLowerCase().replace(/\\/g, "/"));

      // Also read loaded collections from preferences and load their location paths
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
                  if (p) {
                    locationPaths.push(p.toLowerCase().replace(/\\/g, "/"));
                  }
                }
              }
            } catch (err) {
              console.warn(`[searchRoutes] Failed to read collection file ${filePath} for search scope:`, err.message);
            }
          }
        }
      } catch (err) {
        console.warn("[searchRoutes] Failed to read preferences for loaded collections:", err.message);
      }

      const hasExplicitLocationQuery = location && location.trim();

      if (locationPaths.length > 0 && !hasExplicitLocationQuery) {
        results = results.filter(r => {
          const fp = (r.filePath || "").toLowerCase().replace(/\\/g, "/");
          return locationPaths.some(lp => lp && fp.startsWith(lp));
        });
      }
    }
    // When searchScope === "all_locations", no location-based filtering is applied

    // ── Date range filter ────────────────────────────────────────────────────
    if (dateRange && dateRange !== "all") {
      const now = new Date();
      const cutoff = new Date(now);
      switch (dateRange) {
        case "1m":  cutoff.setMonth(now.getMonth() - 1); break;
        case "3m":  cutoff.setMonth(now.getMonth() - 3); break;
        case "6m":  cutoff.setMonth(now.getMonth() - 6); break;
        case "1y":  cutoff.setFullYear(now.getFullYear() - 1); break;
      }
      results = results.filter(r => r.sentAt && new Date(r.sentAt) >= cutoff);
    }

    // ── From filter ─────────────────────────────────────────────────────────
    if (from && from.trim()) {
      const q = from.trim().toLowerCase();
      results = results.filter(r =>
        (r.sender || "").toLowerCase().includes(q)
      );
    }

    // ── To filter ───────────────────────────────────────────────────────────
    if (to && to.trim()) {
      const q = to.trim().toLowerCase();
      results = results.filter(r =>
        Array.isArray(r.recipients)
          ? r.recipients.some(addr => addr.toLowerCase().includes(q))
          : (r.recipients || "").toLowerCase().includes(q)
      );
    }

    // ── CC filter ───────────────────────────────────────────────────────────
    if (cc && cc.trim()) {
      const q = cc.trim().toLowerCase();
      results = results.filter(r =>
        Array.isArray(r.cc)
          ? r.cc.some(addr => addr.toLowerCase().includes(q))
          : (r.cc || "").toLowerCase().includes(q)
      );
    }

    // ── Subject filter ───────────────────────────────────────────────────────
    if (subject && subject.trim()) {
      const q = subject.trim().toLowerCase();
      results = results.filter(r =>
        (r.subject || "").toLowerCase().includes(q)
      );
    }

    // ── Location / filed path filter ─────────────────────────────────────────
    if (location && location.trim()) {
      const q = location.trim().toLowerCase().replace(/\\/g, "/");
      results = results.filter(r =>
        (r.filePath || "").toLowerCase().replace(/\\/g, "/").includes(q)
      );
    }

    // ── Attachments filter ───────────────────────────────────────────────────
    if (hasAttachments === "true") {
      results = results.filter(r => r.hasAttachments === true);
    } else if (hasAttachments === "false") {
      results = results.filter(r => !r.hasAttachments);
    }

    // ── Result kind: email message vs other filed file ───────────────────────
    if (resultKind === "files") {
      results = results.filter((r) => {
        const fp = (r.filePath || "").toLowerCase();
        return fp && !fp.endsWith(".eml") && !fp.endsWith(".msg");
      });
    }

    // ── Body filter ──────────────────────────────────────────────────────────
    if (body && body.trim()) {
      const q = body.trim().toLowerCase();
      results = results.filter(r =>
        (r.body || "").toLowerCase().includes(q)
      );
    }

    // ── General keywords (subject, sender, recipients, path, body; + comment if including) ───
    if (keywords && keywords.trim()) {
      const q = keywords.trim().toLowerCase();
      const includingValue = req.query.including === "true";

      // Parse query to extract terms (words or quoted phrases)
      const termRegex = /"([^"]+)"|(\S+)/g;
      const terms = [];
      let match;
      while ((match = termRegex.exec(q)) !== null) {
        const term = match[1] || match[2];
        if (term && term.trim()) {
          terms.push(term.trim());
        }
      }

      if (terms.length > 0) {
        results = results.filter(r => {
          const recipients = Array.isArray(r.recipients) ? r.recipients.join(" ") : (r.recipients || "");
          const cc = Array.isArray(r.cc) ? r.cc.join(" ") : (r.cc || "");
          
          const searchableFields = [
            r.subject || "",
            r.sender || "",
            recipients,
            cc,
            r.filePath || "",
            r.body || "",
            includingValue ? (r.comment || "") : ""
          ].map(val => val.toLowerCase());

          // Every search term must be found in at least one of the searchable fields
          return terms.every(term =>
            searchableFields.some(field => field.includes(term))
          );
        });
      }
    }

    // ── Dynamic scan of locations for unindexed files ────────────────────────
    if (resultKind !== "files") {
      try {
        // Collect folders to scan (active locations + collections)
        const locations = await getLocations();
        const scanDirs = locations.map(loc => loc.path).filter(Boolean);

        // If the location search query is a full absolute path, scan it directly
        const isAbsolutePath = location && (
          path.isAbsolute(location.trim()) ||
          location.trim().startsWith("\\\\") ||
          /^[a-zA-Z]:/.test(location.trim())
        );
        if (isAbsolutePath) {
          scanDirs.push(location.trim());
        }
        
        // Also read collections from preferences
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
                    if (p) scanDirs.push(p);
                  }
                }
              } catch (err) {}
            }
          }
        } catch (err) {}

        let uniqueDirs = [...new Set(scanDirs.filter(Boolean))];

        // ── Paul's Request: Focus dynamic scan directories based on Location/Job query ──
        if (location && location.trim() && !isAbsolutePath) {
          const locQuery = location.trim().toLowerCase().replace(/\\/g, "/");
          const matchingDirs = uniqueDirs.filter(d => 
            d.toLowerCase().replace(/\\/g, "/").includes(locQuery)
          );
          if (matchingDirs.length > 0) {
            uniqueDirs = matchingDirs;
          }
        }

        if (uniqueDirs.length > 0) {
          const scanPromises = uniqueDirs.map(d => scanDirectory(d, 2));
          const scanResults = await Promise.all(scanPromises);
          const allFilePaths = [...new Set(scanResults.flat().map(p => path.resolve(p)))];

          // Filter out already indexed files
          const indexedPaths = new Set(index.map(r => (r.filePath || "").toLowerCase().replace(/\\/g, "/")));
          const unindexedPaths = allFilePaths.filter(fp => !indexedPaths.has(fp.toLowerCase().replace(/\\/g, "/")));

          if (unindexedPaths.length > 0) {
            const subjectFilter = subject && subject.trim() ? subject.trim().toLowerCase() : null;
            const locationFilter = location && location.trim() ? location.trim().toLowerCase() : null;
            const bodyFilter = body && body.trim() ? body.trim().toLowerCase() : null;
            const fromFilter = from && from.trim() ? from.trim().toLowerCase() : null;
            const toFilter = to && to.trim() ? to.trim().toLowerCase() : null;
            const ccFilter = cc && cc.trim() ? cc.trim().toLowerCase() : null;

            // Parse keywords
            let keywordTerms = [];
            if (keywords && keywords.trim()) {
              const q = keywords.trim().toLowerCase();
              const termRegex = /"([^"]+)"|(\S+)/g;
              let match;
              while ((match = termRegex.exec(q)) !== null) {
                const term = match[1] || match[2];
                if (term && term.trim()) keywordTerms.push(term.trim());
              }
            }

            const hasAnyFilter = !!(keywordTerms.length > 0 || subjectFilter || locationFilter || bodyFilter || fromFilter || toFilter || ccFilter);
            const maxFilesToProcess = hasAnyFilter ? 500 : 100;
            
            let processedCount = 0;
            const matchedUnindexed = [];

            for (const fp of unindexedPaths) {
              const sub = path.basename(fp, path.extname(fp));
              const normPath = fp.toLowerCase().replace(/\\/g, "/");

              let matches = true;

              if (keywordTerms.length > 0) {
                matches = keywordTerms.every(term => 
                  sub.toLowerCase().includes(term) || normPath.includes(term)
                );
              }

              if (matches && subjectFilter && !sub.toLowerCase().includes(subjectFilter)) matches = false;
              if (matches && locationFilter) {
                const normLocFilter = locationFilter.replace(/\\/g, "/");
                if (!normPath.includes(normLocFilter)) matches = false;
              }
              if (matches && bodyFilter) matches = false;
              if (matches && fromFilter && !sub.toLowerCase().includes(fromFilter)) matches = false;
              if (matches && toFilter && !sub.toLowerCase().includes(toFilter)) matches = false;
              if (matches && ccFilter && !sub.toLowerCase().includes(ccFilter)) matches = false;

              if (matches) {
                matchedUnindexed.push({ filePath: fp, subject: sub });
                processedCount++;
                if (processedCount >= maxFilesToProcess) break;
              }
            }

            const unindexedResults = [];
            for (const item of matchedUnindexed) {
              try {
                const stat = await fs.stat(item.filePath);

                if (dateRange && dateRange !== "all") {
                  const now = new Date();
                  const cutoff = new Date(now);
                  switch (dateRange) {
                    case "1m":  cutoff.setMonth(now.getMonth() - 1); break;
                    case "3m":  cutoff.setMonth(now.getMonth() - 3); break;
                    case "6m":  cutoff.setMonth(now.getMonth() - 6); break;
                    case "1y":  cutoff.setFullYear(now.getFullYear() - 1); break;
                  }
                  if (stat.mtime < cutoff) continue;
                }

                unindexedResults.push({
                  id: `unindexed-${item.filePath}-${stat.mtimeMs}`,
                  internetMessageId: null,
                  subject: item.subject,
                  sender: "Legacy Email File (Unindexed)",
                  recipients: [],
                  cc: [],
                  sentAt: stat.mtime.toISOString(),
                  filedAt: stat.mtime.toISOString(),
                  hasAttachments: false,
                  filePath: item.filePath,
                  comment: "Legacy email found via folder search",
                  body: "",
                  isUnindexed: true
                });
              } catch (statErr) {}
            }

            results = [...results, ...unindexedResults];
          }
        }
      } catch (scanErr) {
        console.warn("[searchRoutes] Failed to perform dynamic files scan:", scanErr.message);
      }
    }

    // Sort by sentAt descending
    results.sort((a, b) => new Date(b.sentAt || b.filedAt || 0) - new Date(a.sentAt || a.filedAt || 0));

    // De-duplicate final results by unique filePath
    const seenPaths = new Set();
    const finalResults = [];
    for (const item of results) {
      if (!item.filePath) {
        finalResults.push(item);
        continue;
      }
      const normPath = item.filePath.toLowerCase().replace(/\\/g, "/");
      if (seenPaths.has(normPath)) {
        continue;
      }
      seenPaths.add(normPath);
      finalResults.push(item);
    }

    res.json({ count: finalResults.length, results: finalResults });
  } catch (e) {
    next(e);
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

    // 1. Gather all active locations and shared collection paths
    const locations = await getLocations();
    const scanDirs = locations.map(loc => loc.path).filter(Boolean);
    
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
                if (p) scanDirs.push(p);
              }
            }
          } catch (err) {}
        }
      }
    } catch (err) {}

    const uniqueDirs = [...new Set(scanDirs.filter(Boolean))];

    // 2. Scan directories for files on disk
    let filesOnDisk = [];
    if (uniqueDirs.length > 0) {
      const scanPromises = uniqueDirs.map(d => scanDirectory(d, 2));
      const scanResults = await Promise.all(scanPromises);
      filesOnDisk = scanResults.flat();
    }

    // 3. Prune entries in the index that no longer exist on disk
    const prunedIndex = [];
    const missingPaths = [];
    for (const item of index) {
      if (!item.filePath) continue;
      try {
        await fs.access(item.filePath);
        prunedIndex.push(item);
      } catch (err) {
        missingPaths.push(item.filePath);
      }
    }

    // 4. Identify files on disk that are NOT in the pruned index
    const indexedPaths = new Set(prunedIndex.map(item => (item.filePath || "").toLowerCase().replace(/\\/g, "/")));
    const newFilesToScan = filesOnDisk.filter(fp => !indexedPaths.has(fp.toLowerCase().replace(/\\/g, "/")));

    // 5. Cap legacy file parsing to avoid timeouts
    const MAX_LEGACY_INDEX_PER_RUN = 500;
    const filesToParse = newFilesToScan.slice(0, MAX_LEGACY_INDEX_PER_RUN);

    // 6. Concurrently parse new files (batch size of 25)
    const newRows = [];
    const BATCH_SIZE = 25;
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
              isLegacyIndexed: true
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
      removedCount: index.length - prunedIndex.length,
      addedCount: newRows.length,
      totalCount: updatedIndex.length,
      hasMore: newFilesToScan.length > MAX_LEGACY_INDEX_PER_RUN
    });
  } catch (e) {
    next(e);
  }
});

export default router;
