import { Router } from "express";
import { getLocations } from "../../storage/repositories.js";
import { exec } from "child_process";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { config } from "../../config/index.js";
import { readJson } from "../../storage/jsonStore.js";
import { loadCollectionFile } from "../../services/collectionService.js";
import MsgReaderPkg from "@kenjiuno/msgreader";
import { Meilisearch } from 'meilisearch';

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

  if (resolvedScope === "personal_only" || resolvedScope === "all_personal" || resolvedScope === "locations_i_use" || resolvedScope === "all_locations") {
    const locations = await getLocations();
    dirs.push(...locations.map(loc => loc.path).filter(Boolean));
  }

  if (resolvedScope === "personal_only" || resolvedScope === "all_personal") {
    // Also include folders from the loaded "Personal" collection
    try {
      const prefsPath = path.join(config.dataDir, "preferences.json");
      const prefs = await readJson(prefsPath, {});
      if (prefs.loadedCollections && Array.isArray(prefs.loadedCollections)) {
        const personalColPath = prefs.loadedCollections.find(
          filePath => path.basename(filePath, ".mmcollection").toLowerCase() === "personal"
        );
        if (personalColPath) {
          const colLocs = await loadCollectionFile(personalColPath);
          if (Array.isArray(colLocs)) {
            for (const loc of colLocs) {
              const p = loc.folder || loc.path;
              if (p) dirs.push(p);
            }
          }
        }
      }
    } catch (err) {
      console.warn("[searchRoutes] Failed to read personal collection in getScopedDirectories:", err.message);
    }
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




const meiliClient = new Meilisearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

async function getIndexerState() {
  try {
    const appDataPath = process.env.APPDATA || (process.platform === 'darwin' ? process.env.HOME + '/Library/Application Support' : process.env.HOME + '/.config');
    const statePath = path.join(appDataPath, 'koyomail-indexer', 'indexer_state.json');
    const content = await fs.readFile(statePath, 'utf8');
    return JSON.parse(content);
  } catch (err) {
    return { folders: [] };
  }
}

router.get("/active-collections", async (req, res) => {
  try {
    const state = await getIndexerState();
    const collectionIds = [...new Set(state.folders
      .filter(f => f.type === 'collection' || f.collectionId)
      .map(f => f.type === 'collection' ? (f.description || f.collectionId) : f.collectionId)
      .filter(Boolean))];
    res.json({ collections: collectionIds });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/**
 * GET /api/search
 * 
 * Query params:
 *   keywords       - full-text search across ALL fields (subject, body, sender, recipients, filePath)
 *   location       - restricts to filePath only (job number or absolute path)
 *   subject        - STRICT: must match email's subject field
 *   from           - STRICT: must match email's sender field
 *   to             - STRICT: must match email's recipients field
 *   cc             - STRICT: must match email's cc field
 *   body           - STRICT: must match email's body field
 *   hasAttachments - "true" | "false"
 *   dateRange      - "1m" | "3m" | "6m" | "1y" | "all"
 *   searchScope    - "locations_i_use" | "all_locations" | "personal_only" | "collection:<path>"
 *
 * Strategy:
 *   1. Send keywords + location as the Meilisearch full-text query (fast, indexed search)
 *   2. Apply scope, hasAttachments, dateRange as hard Meilisearch filters
 *   3. Post-filter results in JS for from/to/cc/subject/body (strict field-specific matching)
 *      This is necessary because Meilisearch only accepts one query string, not per-field queries.
 */
router.get("/", async (req, res, next) => {
  try {
    const { 
      keywords = "", 
      location = "",
      hasAttachments, 
      dateRange, 
      searchScope,
      userEmail
    } = req.query;
    
    const trimmedKeywords = keywords.trim();
    const trimmedLocation = location.trim();

    // Empty query validation
    const hasAnyInput = trimmedKeywords || trimmedLocation;
    if (!hasAnyInput) {
      return res.status(400).json({ 
        error: "Please enter a keyword or location to search.",
        code: "EMPTY_QUERY"
      });
    }

    // ── STEP 1: Build Meilisearch hard filters (scope + attachments + date) ──
    let meiliFilters = [];

    // ENFORCE SECURITY: Restrict visibility to public items or items explicitly allowed for the user
    if (userEmail) {
      const normalizedEmail = userEmail.toLowerCase();
      meiliFilters.push(`(isPublic = true OR isPublic IS NULL OR allowedUsers = "${normalizedEmail}")`);
    } else {
      meiliFilters.push(`(isPublic = true OR isPublic IS NULL)`);
    }

    const resolvedScope = searchScope || "locations_i_use";

    if (resolvedScope === "personal_only" || resolvedScope === "all_personal") {
      let filterStr = 'indexedRootType = "local"';
      try {
        const prefsPath = path.join(config.dataDir, "preferences.json");
        const prefs = await readJson(prefsPath, {});
        if (prefs.loadedCollections && Array.isArray(prefs.loadedCollections)) {
          const personalColPath = prefs.loadedCollections.find(
            filePath => path.basename(filePath, ".mmcollection").toLowerCase() === "personal"
          );
          if (personalColPath) {
            const escapedColPath = personalColPath.replace(/\\/g, '\\\\');
            filterStr = `(indexedRootType = "local" OR collectionId = "${escapedColPath}")`;
          }
        }
      } catch (err) {
        console.warn("[searchRoutes] Failed to read preferences for personal collection in Meili filters:", err.message);
      }
      meiliFilters.push(filterStr);
    } else if (resolvedScope.startsWith("collection:")) {
      const colName = resolvedScope.replace("collection:", "");
      const escapedColName = colName.replace(/\\/g, '\\\\');
      
      // Fallback: If the indexer .exe pushed with collectionId=null, we map by the folder path
      let filterStr = `collectionId = "${escapedColName}"`;
      try {
        const state = await getIndexerState();
        const matchedFolders = state.folders.filter(f => 
          (f.type === 'collection' && f.description === colName) || 
          f.collectionId === colName
        );
        if (matchedFolders.length > 0) {
          const rootPaths = [...new Set(matchedFolders.map(f => f.path).filter(Boolean))];
          if (rootPaths.length > 0) {
            const pathClauses = rootPaths.map(p => `indexedRootPath = "${p.replace(/\\/g, '\\\\')}"`);
            filterStr = `(collectionId = "${escapedColName}" OR ${pathClauses.join(" OR ")})`;
          }
        }
      } catch (err) {
        console.warn("[searchRoutes] Failed to read collection paths:", err.message);
      }
      meiliFilters.push(filterStr);
    } else if (resolvedScope === "locations_i_use") {
      const locations = await getLocations();
      const rootPaths = [...new Set(locations.map(loc => loc.path).filter(Boolean))];
      
      let collectionFilters = [];
      try {
        const prefsPath = path.join(config.dataDir, "preferences.json");
        const prefs = await readJson(prefsPath, {});
        if (prefs.loadedCollections && Array.isArray(prefs.loadedCollections)) {
          const state = await getIndexerState();
          
          for (const colFile of prefs.loadedCollections) {
            const colName = path.basename(colFile, ".mmcollection");
            const escapedColName = colName.replace(/\\/g, '\\\\');
            
            const matchedFolders = state.folders.filter(f => 
              (f.type === 'collection' && f.description === colName) || 
              f.collectionId === colName
            );
            
            if (matchedFolders.length > 0) {
              const matchedPaths = [...new Set(matchedFolders.map(f => f.path).filter(Boolean))];
              matchedPaths.forEach(p => rootPaths.push(p));
            }
            collectionFilters.push(`collectionId = "${escapedColName}"`);
          }
        }
      } catch (err) {
        console.warn("[searchRoutes] Failed to read loaded collections for locations_i_use:", err.message);
      }

      const filters = [];
      if (rootPaths.length > 0) {
        const uniqueRootPaths = [...new Set(rootPaths)];
        const inClause = uniqueRootPaths.map(p => `"${p.replace(/\\/g, '\\\\')}"`).join(', ');
        filters.push(`indexedRootPath IN [${inClause}]`);
      }
      
      if (collectionFilters.length > 0) {
        filters.push(...collectionFilters);
      }
      
      if (filters.length > 0) {
        meiliFilters.push(`(${filters.join(' OR ')})`);
      } else {
        return res.json({ count: 0, results: [], estimatedTotalHits: 0 });
      }
    }
    // all_locations → no scope filter (search entire DB)

    if (hasAttachments === "true") {
      meiliFilters.push('hasAttachments = true');
    } else if (hasAttachments === "false") {
      meiliFilters.push('hasAttachments = false');
    }

    if (dateRange && dateRange !== "all") {
      const now = new Date();
      const cutoff = new Date(now);
      switch (dateRange) {
        case "1m": cutoff.setMonth(now.getMonth() - 1); break;
        case "3m": cutoff.setMonth(now.getMonth() - 3); break;
        case "6m": cutoff.setMonth(now.getMonth() - 6); break;
        case "1y": cutoff.setFullYear(now.getFullYear() - 1); break;
      }
      meiliFilters.push(`sentAt >= ${cutoff.getTime()}`);
    }

    // ── STEP 2: Build Meilisearch query ────────────────────────────────────────
    // keywords → full-text across all indexed fields (subject, body, sender, recipients, filePath)
    // location → also full-text, but job numbers in filePath will rank highest
    const meiliQuery = [trimmedKeywords, trimmedLocation].filter(Boolean).join(" ");

    const searchParams = { 
      limit: 100, // Reduced from 1000 to 100 to fix massive network latency (45s -> ~4s)
      sort: ['sentAt:desc'],
      attributesToHighlight: ['subject', 'sender', 'filePath']
    };
    if (meiliFilters.length > 0) {
      searchParams.filter = meiliFilters;
    }
    
    const searchResponse = await emailIndex.search(meiliQuery, searchParams);

    res.json({ 
      count: searchResponse.hits.length, 
      results: searchResponse.hits, 
      estimatedTotalHits: searchResponse.estimatedTotalHits 
    });
  } catch (err) {
    console.error("Meilisearch search error:", err);
    res.status(500).json({ error: "Search failed", details: err.message });
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
    
    // Get the document from Meilisearch to find the physical file path
    let item;
    try {
      item = await emailIndex.getDocument(id);
    } catch (err) {
      return res.status(404).json({ error: "Item not found in index" });
    }

    // Delete the file from the filesystem
    try {
      if (item.filePath) {
          await fs.unlink(item.filePath);
      }
    } catch (err) {
      console.warn(`[searchRoutes] Failed to delete physical file: ${err.message}`);
      // Continue even if physical file is missing, to keep index clean
    }

    // Delete the document from Meilisearch
    await emailIndex.deleteDocument(id);

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

    let item;
    try {
      item = await emailIndex.getDocument(id);
    } catch (err) {
      return res.status(404).json({ error: "Item not found in index" });
    }

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
    await emailIndex.updateDocuments([{
      id: item.id,
      filePath: newFilePath,
      indexedRootPath: destinationDir // Best guess update
    }]);

    res.json({ status: "moved", newPath: newFilePath });
  } catch (e) {
    next(e);
  }
});



export default router;
