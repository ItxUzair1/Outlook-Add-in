import { Router } from "express";
import { getLocations } from "../../storage/repositories.js";
import { exec } from "child_process";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { config } from "../../config/index.js";
import { readJson } from "../../storage/jsonStore.js";
import { getCollectionNameFromPath, loadCollectionFile } from "../../services/collectionService.js";
import { exploreLocation } from "../../services/locationService.js";
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
          filePath => getCollectionNameFromPath(filePath).toLowerCase() === "personal"
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

// Default keyword search skips body for speed; pass includeBody=true to search body too.
const KEYWORD_SEARCH_FIELDS = [
  'subject',
  'sender',
  'recipients',
  'cc',
  'bcc',
  'filePath'
];
const KEYWORD_SEARCH_FIELDS_WITH_BODY = [...KEYWORD_SEARCH_FIELDS, 'body'];

// Metadata-only fields returned in search list (body loaded via /preview).
const SEARCH_LIST_ATTRIBUTES = [
  'id',
  'subject',
  'sender',
  'recipients',
  'cc',
  'bcc',
  'sentAt',
  'filedAt',
  'filePath',
  'hasAttachments',
  'collectionId',
  'indexedRootPath',
  'indexedRootType',
  'isPublic',
];

const SEARCH_PAGE_SIZE = 50;
const SEARCH_MAX_PAGE_SIZE = 100;

function canUserViewDocument(doc, userEmail) {
  if (!doc) return false;
  if (doc.isPublic === true || doc.isPublic == null) return true;
  if (!userEmail) return false;
  const normalizedEmail = userEmail.toLowerCase();
  const allowed = doc.allowedUsers;
  if (Array.isArray(allowed)) {
    return allowed.some((u) => String(u).toLowerCase() === normalizedEmail);
  }
  return String(allowed || '').toLowerCase() === normalizedEmail;
}

function escapeMeiliFilterString(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function normalizePathForCompare(p) {
  return String(p).replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase();
}

/** Backslash and forward-slash variants for Meilisearch exact-match filters. */
function pathFilterVariants(rawPath) {
  const trimmed = String(rawPath).replace(/[/\\]+$/, "");
  if (!trimmed) return [];
  const backslash = trimmed;
  const forward = trimmed.replace(/\\/g, "/");
  return forward === backslash ? [backslash] : [backslash, forward];
}

function longestCommonPathPrefix(paths) {
  if (!paths.length) return "";
  const normalized = paths.map((p) => normalizePathForCompare(p));
  let prefix = normalized[0];
  for (let i = 1; i < normalized.length; i++) {
    while (prefix && !normalized[i].startsWith(prefix)) {
      const cut = prefix.lastIndexOf("/");
      prefix = cut >= 0 ? prefix.slice(0, cut) : "";
    }
    if (!prefix) return "";
  }
  if (!prefix) return "";
  // Preserve backslashes when the first path used them (typical on Windows/UNC).
  return paths[0].includes("\\") ? prefix.replace(/\//g, "\\") : prefix;
}

/**
 * Collapse many collection subfolders to a small set of scan roots for indexedRootPath IN [...].
 * Railway Meilisearch does not support STARTS WITH — we filter on indexedRootPath and rely on
 * the full-text query against filePath for subfolder/project name matching.
 */
function collapsePathsForScopeFilter(paths) {
  const cleaned = [...new Set(
    paths.map((p) => String(p).replace(/[/\\]+$/, "")).filter(Boolean)
  )];
  if (cleaned.length <= 1) return cleaned;

  // Drop strict child paths when a parent path is already in the set.
  const withoutChildren = cleaned.filter((p) => {
    const pNorm = normalizePathForCompare(p);
    return !cleaned.some((other) => {
      if (other === p) return false;
      const oNorm = normalizePathForCompare(other);
      return pNorm.startsWith(`${oNorm}/`);
    });
  });

  const roots = withoutChildren.length > 0 ? withoutChildren : cleaned;
  if (roots.length <= 12) return roots;

  const common = longestCommonPathPrefix(roots);
  if (common && common.length > 10) return [common];

  // Mixed UNC + drive-letter mappings — collapse each prefix group separately.
  const groups = new Map();
  for (const p of roots) {
    const parts = normalizePathForCompare(p).split("/").filter(Boolean);
    const key = parts.slice(0, Math.min(4, parts.length)).join("/");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(p);
  }

  const groupRoots = [];
  for (const groupPaths of groups.values()) {
    const groupCommon = longestCommonPathPrefix(groupPaths);
    if (groupCommon && groupCommon.length > 3) {
      groupRoots.push(groupCommon);
    } else {
      groupRoots.push(...groupPaths);
    }
  }

  return [...new Set(groupRoots)];
}

/** Build indexedRootPath IN filter using only operators supported by Meilisearch v1.10 and earlier. */
function buildRootPathScopeFilter(rootPaths) {
  const collapsed = collapsePathsForScopeFilter(rootPaths);
  const inValues = new Set();

  for (const p of collapsed) {
    for (const variant of pathFilterVariants(p)) {
      inValues.add(`"${escapeMeiliFilterString(variant)}"`);
    }
  }

  if (inValues.size === 0) return null;
  if (inValues.size === 1) {
    return `(indexedRootPath = ${[...inValues][0]})`;
  }
  return `(indexedRootPath IN [${[...inValues].join(", ")}])`;
}

async function getLoadedCollectionFiles() {
  try {
    const prefsPath = path.join(config.dataDir, "preferences.json");
    const prefs = await readJson(prefsPath, {});
    return Array.isArray(prefs.loadedCollections) ? prefs.loadedCollections : [];
  } catch {
    return [];
  }
}

/** Resolve all root folder paths associated with a collection name. */
async function getCollectionRootPaths(colName) {
  const rootPaths = new Set();

  try {
    const locations = await getLocations();
    for (const loc of locations) {
      if (loc.collection && loc.collection.toLowerCase() === colName.toLowerCase() && loc.path) {
        rootPaths.add(loc.path);
      }
    }
  } catch (err) {
    console.warn("[searchRoutes] Failed to read locations for collection paths:", err.message);
  }

  const loadedCollections = await getLoadedCollectionFiles();
  for (const filePath of loadedCollections) {
    const name = getCollectionNameFromPath(filePath);
    if (name.toLowerCase() !== colName.toLowerCase()) continue;
    try {
      const colLocs = await loadCollectionFile(filePath);
      if (Array.isArray(colLocs)) {
        for (const loc of colLocs) {
          const p = loc.folder || loc.path;
          if (p) rootPaths.add(p);
        }
      }
    } catch (err) {
      console.warn("[searchRoutes] Failed to read collection file for scope:", err.message);
    }
  }

  try {
    const indexerState = await getIndexerState();
    const matchedFolders = indexerState.folders.filter(
      (f) => (f.type === "collection" && f.description === colName) || f.collectionId === colName
    );
    for (const folder of matchedFolders) {
      if (folder.path) rootPaths.add(folder.path);
    }
  } catch (err) {
    console.warn("[searchRoutes] Failed to read indexer state for collection paths:", err.message);
  }

  return [...rootPaths];
}

/** Root paths for "Locations I Use" — locations.json, collection files, and indexer folders. */
async function getLocationsIUseRootPaths() {
  const rootPaths = new Set();

  try {
    const locations = await getLocations();
    for (const loc of locations) {
      if (loc.path) rootPaths.add(loc.path);
    }
  } catch (err) {
    console.warn("[searchRoutes] Failed to read locations for locations_i_use:", err.message);
  }

  const loadedCollections = await getLoadedCollectionFiles();
  for (const filePath of loadedCollections) {
    try {
      const colLocs = await loadCollectionFile(filePath);
      if (Array.isArray(colLocs)) {
        for (const loc of colLocs) {
          const p = loc.folder || loc.path;
          if (p) rootPaths.add(p);
        }
      }
    } catch (err) {
      console.warn("[searchRoutes] Failed to read collection paths for locations_i_use:", err.message);
    }
  }

  try {
    const indexerState = await getIndexerState();
    for (const colFile of loadedCollections) {
      const colName = getCollectionNameFromPath(colFile);
      const matchedFolders = indexerState.folders.filter(
        (f) => (f.type === "collection" && f.description === colName) || f.collectionId === colName
      );
      for (const folder of matchedFolders) {
        if (folder.path) rootPaths.add(folder.path);
      }
    }
  } catch (err) {
    console.warn("[searchRoutes] Failed to read indexer folders for locations_i_use:", err.message);
  }

  return [...rootPaths];
}

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
 *   searchScope    - "locations_i_use" | "all_locations" | "personal_only" | "collection:<path>"
 *   includeBody    - "true" to include email body in keyword search (slower)
 *   offset         - pagination offset (default 0)
 *   limit          - page size (default 50, max 100)
 *
 * Strategy:
 *   1. Send keywords + location as the Meilisearch full-text query (fast, indexed search)
 *   2. Apply scope, hasAttachments as hard Meilisearch filters
 *   3. Post-filter results in JS for from/to/cc/subject/body (strict field-specific matching)
 *      This is necessary because Meilisearch only accepts one query string, not per-field queries.
 */
router.get("/", async (req, res, next) => {
  try {
    const { 
      keywords = "", 
      location = "",
      hasAttachments, 
      timeSpan,
      userEmail,
      offset: offsetParam,
      limit: limitParam,
    } = req.query;
    
    const parsedOffset = Math.max(0, parseInt(offsetParam, 10) || 0);
    const parsedLimit = Math.min(
      SEARCH_MAX_PAGE_SIZE,
      Math.max(1, parseInt(limitParam, 10) || SEARCH_PAGE_SIZE)
    );
    
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

    if (timeSpan && timeSpan !== "all_time") {
      const now = Date.now();
      const periods = {
        "past_week": 7 * 24 * 60 * 60 * 1000,
        "past_month": 30 * 24 * 60 * 60 * 1000,
        "past_3_months": 90 * 24 * 60 * 60 * 1000,
        "past_6_months": 180 * 24 * 60 * 60 * 1000,
        "past_year": 365 * 24 * 60 * 60 * 1000
      };
      if (periods[timeSpan]) {
        const threshold = now - periods[timeSpan];
        meiliFilters.push(`sentAt >= ${threshold}`);
      }
    }

    if (hasAttachments === "true") {
      meiliFilters.push('hasAttachments = true');
    } else if (hasAttachments === "false") {
      meiliFilters.push('hasAttachments = false');
    }

    // ── STEP 2: Build Meilisearch query ────────────────────────────────────────
    const meiliQueryParts = [];
    if (trimmedKeywords) meiliQueryParts.push(trimmedKeywords);
    // Remove quotes around trimmedLocation so Meilisearch can do prefix matching on partial project names
    if (trimmedLocation) meiliQueryParts.push(trimmedLocation);
    
    const meiliQuery = meiliQueryParts.join(" ");
    const searchParams = { 
      limit: parsedLimit,
      offset: parsedOffset,
      matchingStrategy: 'all',
      attributesToRetrieve: SEARCH_LIST_ATTRIBUTES,
      attributesToHighlight: ['subject', 'sender', 'filePath'],
      attributesToSearchOn: KEYWORD_SEARCH_FIELDS_WITH_BODY, // Always include body
    };
    if (meiliFilters.length > 0) {
      searchParams.filter = meiliFilters;
    }
    
    const searchResponse = await emailIndex.search(meiliQuery, searchParams);

    const pageHits = searchResponse.hits;
    const estimatedTotalHits = searchResponse.estimatedTotalHits ?? pageHits.length;
    const loadedThrough = parsedOffset + pageHits.length;

    res.json({ 
      count: pageHits.length,
      results: pageHits,
      estimatedTotalHits,
      offset: parsedOffset,
      limit: parsedLimit,
      hasMore: loadedThrough < estimatedTotalHits,
      loadedCount: loadedThrough,
    });
  } catch (err) {
    console.error("Meilisearch search error:", err);
    res.status(500).json({ error: "Search failed", details: err.message });
  }
});


/**
 * GET /api/search/preview?id=...
 * Returns full email body (and metadata) for the preview pane — not included in list search.
 */
router.get("/preview", async (req, res, next) => {
  try {
    const { id, userEmail } = req.query;
    if (!id) {
      return res.status(400).json({ error: "id is required" });
    }

    let doc;
    try {
      doc = await emailIndex.getDocument(String(id));
    } catch (err) {
      return res.status(404).json({ error: "Item not found in index" });
    }

    if (!canUserViewDocument(doc, userEmail)) {
      return res.status(403).json({ error: "You do not have permission to view this item" });
    }

    res.json({
      id: doc.id,
      subject: doc.subject,
      sender: doc.sender,
      recipients: doc.recipients,
      cc: doc.cc,
      sentAt: doc.sentAt,
      hasAttachments: doc.hasAttachments,
      filePath: doc.filePath,
      body: (doc.body || "")
        .replace(/<style[\s\S]*?<\/style>/gi, "")
        .replace(/<script[\s\S]*?<\/script>/gi, "")
        .replace(/<!--[\s\S]*?-->/g, "")
        .replace(/<[^>]*>?/gm, "")
        .replace(/&nbsp;/gi, " ")
        .trim(),
    });
  } catch (err) {
    console.error("Meilisearch preview error:", err);
    res.status(500).json({ error: "Preview failed", details: err.message });
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
 * POST /api/search/copy
 * Copies the file to the native Windows clipboard.
 */
router.post("/copy", async (req, res, next) => {
  try {
    const { filePath, filePaths } = req.body;
    const pathsToCopy = filePaths || (filePath ? [filePath] : []);
    
    if (pathsToCopy.length === 0) return res.status(400).json({ error: "filePath or filePaths is required" });

    // Verify files exist first
    const validPaths = [];
    for (const p of pathsToCopy) {
      try {
        await fs.access(p);
        validPaths.push(p);
      } catch (err) {
        console.warn(`[searchRoutes] Copy attempt failed: File not found at ${p}`);
      }
    }
    
    if (validPaths.length === 0) {
      return res.status(404).json({ error: "No files found at original locations", code: "ENOENT" });
    }

    // Use PowerShell's Set-Clipboard cmdlet
    // Use -LiteralPath with comma separated paths
    const formattedPaths = validPaths.map(p => `'${p.replace(/'/g, "''")}'`).join(", ");
    const psCmd = `powershell.exe -NoProfile -Command "Set-Clipboard -LiteralPath ${formattedPaths}"`;
    
    exec(psCmd, (error) => {
      if (error) {
          console.error(`[searchRoutes] Failed to copy file to clipboard: ${error.message}`);
          return res.status(500).json({ error: `Could not copy file: ${error.message}` });
      }
      res.json({ status: "success", copiedCount: validPaths.length });
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

    // Open the folder in the foreground
    await exploreLocation(dirPath);
    res.json({ status: "success" });
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
