const fs = require('fs');
const path = require('path');
const { simpleParser } = require('mailparser');
const MsgReaderPkg = require('@kenjiuno/msgreader');
const MsgReader = MsgReaderPkg.default || MsgReaderPkg;

/**
 * Coerces parsed email fields to plain strings for Meilisearch.
 * msgreader/mailparser sometimes return body as object, Buffer, or array.
 */
function toSearchableText(value, maxLen = 50000) {
  if (value == null || value === '') return '';
  if (typeof value === 'string') return value.length > maxLen ? value.substring(0, maxLen) : value;
  if (Buffer.isBuffer(value)) {
    const text = value.toString('utf8');
    return text.length > maxLen ? text.substring(0, maxLen) : text;
  }
  if (Array.isArray(value)) {
    return toSearchableText(value.map(v => toSearchableText(v, maxLen)).filter(Boolean).join(' '), maxLen);
  }
  if (typeof value === 'object') {
    if (typeof value.text === 'string') return toSearchableText(value.text, maxLen);
    if (typeof value.content === 'string') return toSearchableText(value.content, maxLen);
    if (typeof value.value === 'string') return toSearchableText(value.value, maxLen);
    try {
      return toSearchableText(JSON.stringify(value), maxLen);
    } catch {
      return '';
    }
  }
  const text = String(value);
  return text.length > maxLen ? text.substring(0, maxLen) : text;
}

function toAddressText(value) {
  if (value == null) return '';
  if (typeof value === 'string') return value;
  if (typeof value?.text === 'string') return value.text;
  if (Array.isArray(value)) return value.map(toAddressText).filter(Boolean).join(', ');
  return String(value);
}

function formatMsgRecipient(rec) {
  const addr = rec.emailAddress || rec.smtpAddress || rec.email || '';
  const name = rec.name && rec.name !== addr ? rec.name : '';
  return addr ? (name ? `${name} <${addr}>` : addr) : (rec.name || '');
}

function splitRecipientsByType(recipients) {
  const toList = [];
  const ccList = [];
  const bccList = [];

  if (!Array.isArray(recipients)) {
    return { toList, ccList, bccList };
  }

  for (const rec of recipients) {
    const full = formatMsgRecipient(rec);
    if (!full) continue;

    const type = rec.recipType || rec.recipientType;
    if (type === 'cc') ccList.push(full);
    else if (type === 'bcc') bccList.push(full);
    else toList.push(full);
  }

  return { toList, ccList, bccList };
}

function parseTransportHeaders(headers) {
  if (!headers || typeof headers !== 'string') return null;

  const result = { to: '', cc: '', date: '' };
  const unfolded = headers.replace(/\r?\n[ \t]+/g, ' ');
  const lines = unfolded.split(/\r?\n/);

  for (const line of lines) {
    const colonIndex = line.indexOf(':');
    if (colonIndex === -1) continue;
    const key = line.slice(0, colonIndex).trim().toLowerCase();
    const value = line.slice(colonIndex + 1).trim();
    if (key === 'to' && !result.to) result.to = value;
    else if (key === 'cc' && !result.cc) result.cc = value;
    else if (key === 'date' && !result.date) result.date = value;
  }

  return result;
}

function extractMsgTimestamp(info) {
  const candidates = [info.clientSubmitTime, info.messageDeliveryTime, info.date];
  for (const value of candidates) {
    if (!value) continue;
    const dateObj = new Date(value);
    if (!isNaN(dateObj.getTime())) return dateObj.getTime();
  }
  return 0;
}

/**
 * Extracts data from a .msg or .eml file.
 * Returns an object with the required search fields.
 * @param {string} filePath Absolute path to the email file
 * @returns {Promise<Object>} Extracted email data
 */
async function parseEmailFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.eml') {
    return parseEml(filePath);
  } else if (ext === '.msg') {
    return parseMsg(filePath);
  } else {
    throw new Error(`Unsupported file extension: ${ext}`);
  }
}

async function parseEml(filePath) {
  const fileStream = fs.createReadStream(filePath);
  const parsed = await simpleParser(fileStream);

  return {
    subject: toSearchableText(parsed.subject, 1000),
    sender: toAddressText(parsed.from),
    recipients: toAddressText(parsed.to),
    cc: toAddressText(parsed.cc),
    bcc: toAddressText(parsed.bcc),
    sentAt: parsed.date ? parsed.date.getTime() : 0,
    body: toSearchableText(parsed.text || parsed.html),
    hasAttachments: parsed.attachments && parsed.attachments.length > 0,
    filePath: filePath,
    comment: ''
  };
}

async function parseMsg(filePath) {
  const buffer = await fs.promises.readFile(filePath);
  
  // Check if it's actually a JSON file disguised as .msg
  const textPreview = buffer.toString('utf8', 0, 500).trim();
  if (textPreview.startsWith('{')) {
    try {
      const parsed = JSON.parse(buffer.toString('utf8'));
      if (parsed.internetMessageId || parsed.subject || parsed.sentAt) {
        return {
          subject: toSearchableText(parsed.subject, 1000),
          sender: toSearchableText(parsed.sender),
          recipients: toSearchableText((parsed.to || []).join(', ')),
          cc: toSearchableText((parsed.cc || []).join(', ')),
          bcc: '',
          sentAt: parsed.sentAt ? new Date(parsed.sentAt).getTime() : 0,
          body: toSearchableText(parsed.bodyPreview || parsed.body),
          hasAttachments: !!parsed.hasAttachments,
          filePath: filePath,
          comment: ''
        };
      }
    } catch (e) {
      // Not a valid JSON file, fallback to MsgReader
    }
  }

  try {
    const reader = new MsgReader(buffer);
    const parsed = reader.getFileData();

    if (parsed.error) {
      if (parsed.error.includes('Unsupported file type')) {
        try {
          // Some email clients save standard MIME EML files with a .msg extension.
          return await parseEml(filePath);
        } catch (fallbackErr) {
          throw new Error(`Unsupported file format. File is neither a valid OLE MSG nor a standard EML.`);
        }
      }
      throw new Error(`Error parsing MSG file: ${parsed.error}`);
    }

    let sender = '';
    if (parsed.senderEmail) {
      sender = parsed.senderName
        ? `${parsed.senderName} <${parsed.senderEmail}>`
        : parsed.senderEmail;
    } else {
      sender = parsed.senderName || '';
    }

    let { toList, ccList, bccList } = splitRecipientsByType(parsed.recipients);
    let timestamp = extractMsgTimestamp(parsed);

    // Fallback: parse To/Cc/Date from embedded transport headers when recipient
    // sub-objects are missing (common in forwarded/sent-item MSG files).
    if ((toList.length === 0 || !timestamp) && parsed.headers) {
      const headerData = parseTransportHeaders(parsed.headers);
      if (headerData) {
        if (toList.length === 0 && headerData.to) toList = [headerData.to];
        if (ccList.length === 0 && headerData.cc) ccList = [headerData.cc];
        if (!timestamp && headerData.date) {
          const headerDate = new Date(headerData.date);
          if (!isNaN(headerDate.getTime())) timestamp = headerDate.getTime();
        }
      }
    }

    return {
      subject: toSearchableText(parsed.subject, 1000),
      sender: toSearchableText(sender),
      recipients: toSearchableText(toList.join(', ')),
      cc: toSearchableText(ccList.join(', ')),
      bcc: toSearchableText(bccList.join(', ')),
      sentAt: timestamp,
      body: toSearchableText(parsed.body),
      hasAttachments: parsed.attachments && parsed.attachments.length > 0,
      filePath: filePath,
      comment: ''
    };
  } catch (err) {
    console.warn(`[Indexer] Failed to parse MSG file ${filePath} with MsgReader:`, err.message);
    // Safe fallback: parse from file name and filesystem stats
    try {
      const stat = fs.statSync(filePath);
      const baseName = path.basename(filePath, path.extname(filePath));
      
      let subject = baseName;
      let sentAt = stat.mtime.getTime();
      
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
          if (!isNaN(parsedDate.getTime())) {
            sentAt = parsedDate.getTime();
          }
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
            if (!isNaN(parsedDate.getTime())) {
              sentAt = parsedDate.getTime();
            }
          } catch (e) {}
        }
      }
      
      return {
        subject: subject || baseName,
        sender: "Legacy Email",
        recipients: "",
        cc: "",
        bcc: "",
        sentAt: sentAt,
        body: "",
        hasAttachments: false,
        filePath: filePath,
        comment: "Parsing failed, using filesystem fallback"
      };
    } catch (fallbackErr) {
      throw new Error(`Failed to parse MSG file: ${err.message}`);
    }
  }
}

module.exports = {
  parseEmailFile,
  toSearchableText
};
