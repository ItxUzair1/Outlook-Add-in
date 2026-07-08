const fs = require('fs');
const path = require('path');
const MsgReaderPkg = require('@kenjiuno/msgreader');
const MsgReader = MsgReaderPkg.default || MsgReaderPkg;
const { MeiliSearch } = require('meilisearch');

require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

function yieldToEventLoop() {
  return new Promise(resolve => setImmediate(resolve));
}

function getMetadata(filePath) {
  try {
    const fileBuffer = fs.readFileSync(filePath);
    const reader = new MsgReader(fileBuffer);
    const info = reader.getFileData();

    if (info.error) return null;

    let sender = "";
    if (info.senderEmail) {
      sender = info.senderName ? `${info.senderName} <${info.senderEmail}>` : info.senderEmail;
    } else {
      sender = info.senderName || "";
    }

    const toList = [];
    const ccList = [];
    if (Array.isArray(info.recipients)) {
      for (const rec of info.recipients) {
        const addr = rec.emailAddress || rec.smtpAddress || "";
        const name = rec.name && rec.name !== addr ? rec.name : "";
        const full = addr ? (name ? `${name} <${addr}>` : addr) : rec.name || "";
        if (full) {
          if (rec.recipType === 'cc' || rec.recipientType === 'cc' || rec.recipientType === 2) {
            ccList.push(full);
          } else {
            toList.push(full);
          }
        }
      }
    }
    
    return {
      subject: info.subject || "",
      sender: sender || "Unknown Sender",
      recipients: toList.join(', '),
      cc: ccList.join(', '),
      body: (info.body || "").substring(0, 5000).trim(),
    };
  } catch (err) {
    return null;
  }
}

async function runReindexUnknown({ log = console.log, onProgress = () => {} }) {
  log("Fetching Unknown Sender emails from Meilisearch...");
  const res = await emailIndex.search('"Unknown Sender"', {
    limit: 1500,
    attributesToRetrieve: ['id', 'filePath'],
  });

  const hits = res.hits;
  log(`Found ${hits.length} emails to process.`);
  
  if (hits.length === 0) return { success: true, count: 0 };

  let successCount = 0;
  let updates = [];

  for (let i = 0; i < hits.length; i++) {
    const doc = hits[i];
    
    // Yield occasionally to keep the server responsive
    if (i % 25 === 0) await yieldToEventLoop();

    onProgress({
      total: hits.length,
      scanned: i + 1,
      repaired: successCount,
      skipped: i - successCount,
      currentFilePath: doc.filePath,
    });

    if (!fs.existsSync(doc.filePath)) {
      continue;
    }

    const metadata = getMetadata(doc.filePath);
    if (metadata && metadata.sender !== "Unknown Sender" && metadata.sender.trim() !== "") {
      updates.push({
        id: doc.id,
        subject: metadata.subject,
        sender: metadata.sender,
        recipients: metadata.recipients,
        cc: metadata.cc,
        body: metadata.body,
      });
      successCount++;
    }
    
    if (updates.length >= 100) {
      log(`Sending batch of 100 updates to Meilisearch...`);
      await emailIndex.updateDocuments(updates);
      updates = [];
    }
  }

  if (updates.length > 0) {
    await emailIndex.updateDocuments(updates);
  }

  log(`Finished! Successfully re-extracted metadata for ${successCount} emails.`);
  return { success: true, count: successCount };
}

module.exports = { runReindexUnknown };
