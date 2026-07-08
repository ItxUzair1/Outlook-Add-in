const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const MsgReaderPkg = require('@kenjiuno/msgreader');
const MsgReader = MsgReaderPkg.default || MsgReaderPkg;
const state = require('./state');

const { MeiliSearch } = require('meilisearch');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

// Fixes lone surrogate crashes in Meilisearch JSON payloads
function sanitizeSurrogates(str) {
  if (typeof str !== 'string') return str;
  if (str.toWellFormed) return str.toWellFormed();
  return str.replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|([^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]/g, "$1\uFFFD");
}

const RETRY_BATCH_SIZE = 250;
const YIELD_EVERY_N = 25; // Yield to event loop frequently to keep UI perfectly responsive
const FILE_SIZE_THRESHOLD = 50 * 1024 * 1024; // 50 MB threshold

function yieldToEventLoop() {
  return new Promise(resolve => setImmediate(resolve));
}

function decodeQuotedPrintable(str) {
  if (!str) return "";
  return str
    .replace(/=\r?\n/g, "")
    .replace(/=([0-9A-F]{2})/gi, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
}

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

/**
 * Robust fallback parser that bypasses heavy memory/CPU parsing operations.
 */
async function parseErrorFileFallback(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const stat = await fs.promises.stat(filePath);
  const baseName = path.basename(filePath, ext);

  // Defaults
  let subject = baseName;
  let sender = "Unknown Sender";
  let recipients = "";
  let cc = "";
  let sentAt = stat.mtime.getTime();
  let body = "";
  let hasAttachments = false;

  try {
    if (ext === '.eml') {
      // EML Fallback
      let content = "";
      
      if (stat.size > FILE_SIZE_THRESHOLD) {
        // If file is > 5MB, only read the first 128KB to protect memory
        const fileHandle = await fs.promises.open(filePath, 'r');
        const buffer = Buffer.alloc(131072);
        const { bytesRead } = await fileHandle.read(buffer, 0, 131072, 0);
        await fileHandle.close();
        content = buffer.toString('utf8', 0, bytesRead);
      } else {
        content = await fs.promises.readFile(filePath, 'utf8');
      }

      const headerEndIndex = content.search(/\r?\n\r?\n/);
      const headerText = headerEndIndex !== -1 ? content.slice(0, headerEndIndex) : content;
      const bodyText = headerEndIndex !== -1 ? content.slice(headerEndIndex + 4, headerEndIndex + 15000) : "";

      const unfoldedText = headerText.replace(/\r?\n[ \t]+/g, ' ');
      const headers = {};
      const lines = unfoldedText.split(/\r?\n/);
      for (const line of lines) {
        const colonIndex = line.indexOf(':');
        if (colonIndex !== -1) {
          const key = line.slice(0, colonIndex).trim().toLowerCase();
          const value = line.slice(colonIndex + 1).trim();
          headers[key] = value;
        }
      }

      subject = decodeRFC2047(headers.subject || '') || subject;
      sender = decodeRFC2047(headers.from || '') || sender;
      recipients = decodeRFC2047(headers.to || '');
      cc = decodeRFC2047(headers.cc || '');
      
      if (headers.date) {
        const parsedDate = new Date(headers.date);
        if (!isNaN(parsedDate.getTime())) {
          sentAt = parsedDate.getTime();
        }
      }

      hasAttachments = /Content-Disposition:\s*attachment/i.test(content) || 
                       /Content-Type:\s*[^;\s]+;\s*name=/i.test(content);

      // Extract a clean body preview from plain text/HTML blocks
      body = bodyText
        .replace(/<style[\s\S]*?<\/style>/gi, '')
        .replace(/<script[\s\S]*?<\/script>/gi, '')
        .replace(/<[^>]*>?/gm, '')
        .replace(/&nbsp;/gi, ' ')
        .substring(0, 5000)
        .trim();
        
    } else if (ext === '.msg') {
      // MSG Fallback
      if (stat.size > FILE_SIZE_THRESHOLD) {
        // If MSG is larger than 5MB, bypass MsgReader parsing completely to prevent memory exhaust
        throw new Error("File exceeds 5MB memory limit, falling back to name metadata");
      }

      const fileBuffer = await fs.promises.readFile(filePath);
      const reader = new MsgReader(fileBuffer);
      const info = reader.getFileData();

      subject = info.subject || subject;
      if (info.senderEmail) {
        sender = info.senderName ? `${info.senderName} <${info.senderEmail}>` : info.senderEmail;
      } else {
        sender = info.senderName || sender;
      }

      const toList = [];
      const ccList = [];
      if (Array.isArray(info.recipients)) {
        for (const rec of info.recipients) {
          const addr = rec.emailAddress || rec.smtpAddress || "";
          const name = rec.name && rec.name !== addr ? rec.name : "";
          const full = addr ? (name ? `${name} <${addr}>` : addr) : rec.name || "";
          if (full) {
            if (rec.recipType === 'cc' || rec.recipientType === 'cc') {
              ccList.push(full);
            } else {
              toList.push(full);
            }
          }
        }
      }
      recipients = toList.join(', ');
      cc = ccList.join(', ');

      if (info.clientSubmitTime) {
        const d = new Date(info.clientSubmitTime);
        if (!isNaN(d.getTime())) sentAt = d.getTime();
      } else if (info.messageDeliveryTime) {
        const d = new Date(info.messageDeliveryTime);
        if (!isNaN(d.getTime())) sentAt = d.getTime();
      }

      hasAttachments = info.attachments && info.attachments.length > 0;
      body = (info.body || "").substring(0, 5000).trim();
    }
  } catch (err) {
    // If anything fails or throws, we fallback to filename and date stats
    console.warn(`[Fallback Parser] Active fallback for ${filePath}: ${err.message}`);
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
        const parsedDate = new Date(`${year}-${month}-${day}T${hour}:${min}:${sec}.000Z`);
        if (!isNaN(parsedDate.getTime())) sentAt = parsedDate.getTime();
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
          const parsedDate = new Date(`${year}-${month}-${day}T00:00:00.000Z`);
          if (!isNaN(parsedDate.getTime())) sentAt = parsedDate.getTime();
        } catch (e) {}
      }
    }
  }

  return {
    subject: String(subject || "").trim(),
    sender: String(sender || "").trim(),
    recipients: String(recipients || "").trim(),
    cc: String(cc || "").trim(),
    bcc: "",
    sentAt,
    body: String(body || "").substring(0, 5000),
    hasAttachments,
    filePath,
    comment: 'Parsed by robust fallback parser'
  };
}

/**
 * Orchestrator loop for retrying indexing of failed emails.
 */
async function runRetryErrors() {
  state.updateIndexingStatus('retrying');
  state.addLog("Starting fallback recovery for unparseable error emails...");

  const errorFiles = state.getUnparseableFiles();
  if (errorFiles.length === 0) {
    state.addLog("No error emails found in ledger to retry. Done.");
    state.updateIndexingStatus('idle');
    return;
  }

  state.addLog(`Found ${errorFiles.length} emails in the error log. Running fallback recovery...`);
  
  // Set initial metrics
  state.updateStats({
    totalFilesFound: errorFiles.length,
    filesIndexedThisSession: 0,
    filesSkipped: 0,
    currentFilePath: "Starting recovery..."
  }, { immediate: true });

  const folders = state.getFolders();
  let indexedCount = 0;
  let batch = [];
  let batchFilePaths = [];

  // Asynchronous loop
  for (let i = 0; i < errorFiles.length; i++) {
    const filePath = errorFiles[i];

    // Check if job got paused/cancelled
    const status = state.getIndexingStatus();
    if (status === 'paused' || status === 'idle') {
      state.addLog("Recovery job paused/stopped by user.");
      state.updateIndexingStatus('idle');
      return;
    }

    state.updateStats({ currentFilePath: filePath }, { persist: false });

    // Yield control back to event loop to keep UI smooth and responsive
    if (i > 0 && i % YIELD_EVERY_N === 0) {
      await yieldToEventLoop();
    }

    try {
      // Find matching folder config to apply allowedUsers / isPublic correctly
      const folder = folders.find(f => filePath.toLowerCase().startsWith(f.path.toLowerCase())) || {};
      
      const parsed = await parseErrorFileFallback(filePath);
      
      batch.push({
        id: crypto.createHash('sha256').update(filePath).digest('hex'),
        subject: sanitizeSurrogates(parsed.subject),
        sender: sanitizeSurrogates(parsed.sender),
        recipients: sanitizeSurrogates(parsed.recipients),
        cc: sanitizeSurrogates(parsed.cc),
        bcc: sanitizeSurrogates(parsed.bcc),
        sentAt: parsed.sentAt,
        body: sanitizeSurrogates(parsed.body),
        hasAttachments: parsed.hasAttachments,
        filePath: parsed.filePath,
        comment: sanitizeSurrogates(parsed.comment),
        indexedRootPath: folder.path || "",
        indexedRootType: folder.type || 'local',
        collectionId: folder.type === 'collection'
          ? (folder.description || folder.collectionId)
          : (folder.collectionId || null),
        isPublic: folder.isPublic !== false,
        allowedUsers: (folder.allowedUsers || []).map(u => u.toLowerCase())
      });
      batchFilePaths.push(filePath);
    } catch (err) {
      console.error(`[Recovery] Unexpected parsing crash on ${filePath}:`, err.message);
    }

    // Flush batch when size matches
    if (batch.length >= RETRY_BATCH_SIZE) {
      await flushRetryBatch(batch, batchFilePaths);
      indexedCount += batch.length;
      batch = [];
      batchFilePaths = [];
    }
  }

  // Final flush
  if (batch.length > 0) {
    await flushRetryBatch(batch, batchFilePaths);
    indexedCount += batch.length;
  }

  state.addLog(`Recovery complete. Successfully indexed ${indexedCount} of ${errorFiles.length} previously failed emails.`);
  state.updateIndexingStatus('idle');
  state.updateStats({ currentFilePath: "", speed: 0 }, { immediate: true });
}

async function flushRetryBatch(documentsBatch, pathsBatch) {
  try {
    state.addLog(`Recovery: Uploading batch of ${documentsBatch.length} recovered emails to Meilisearch...`);
    await emailIndex.addDocuments(documentsBatch, { primaryKey: 'id' });
    
    // Clear them from ledger and set as uploaded
    for (const filePath of pathsBatch) {
      state.markFileUploaded(filePath);
      state.removeFileUnparseable(filePath);
    }

    // Update statistics
    const stats = state.getStats();
    const newFilesFailed = Math.max(0, state.getUnparseableFiles().length);
    state.updateStats({
      filesIndexed: stats.filesIndexed + documentsBatch.length,
      filesFailed: newFilesFailed,
      filesIndexedThisSession: stats.filesIndexedThisSession + documentsBatch.length
    }, { immediate: true });

  } catch (err) {
    state.addLog(`Recovery upload batch failed: ${err.message}`);
    state.addErrorLog('Recovery Batch Upload', err.message);
  }
}

module.exports = {
  runRetryErrors
};
