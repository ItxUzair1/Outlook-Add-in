/**
 * Robust email parser for recovery/re-index paths.
 * Handles Unknown Sender, empty To/body, HTML-only bodies, transport headers,
 * and large files without loading entire attachments into memory.
 */
const fs = require('fs');
const path = require('path');
const { simpleParser } = require('mailparser');
const MsgReaderPkg = require('@kenjiuno/msgreader');
const MsgReader = MsgReaderPkg.default || MsgReaderPkg;
const { decompressRTF } = require('@kenjiuno/decompressrtf');

const FILE_SIZE_THRESHOLD = 50 * 1024 * 1024;
const HEADER_READ_SIZE = 131072;
const DEFAULT_BODY_MAX = 150000;
const DEFAULT_SUBJECT_MAX = 1000;

const BAD_SENDERS = new Set(['unknown sender', 'legacy email', '']);

function decodeRFC2047(str) {
  if (!str) return '';
  return str.replace(/=\?([^?]+)\?([QB])\?([^?]*)\?=/gi, (match, charset, encoding, text) => {
    if (encoding.toUpperCase() === 'B') {
      try {
        return Buffer.from(text, 'base64').toString(charset.toLowerCase() === 'utf-8' ? 'utf8' : 'binary');
      } catch {
        return text;
      }
    }
    if (encoding.toUpperCase() === 'Q') {
      const decoded = text.replace(/_/g, ' ').replace(/=([0-9A-F]{2})/gi, (m, hex) =>
        String.fromCharCode(parseInt(hex, 16))
      );
      try {
        return Buffer.from(decoded, 'binary').toString(charset.toLowerCase() === 'utf-8' ? 'utf8' : 'binary');
      } catch {
        return decoded;
      }
    }
    return match;
  });
}

function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]*>?/gm, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function truncateText(str, maxLen) {
  if (!str) return '';
  const text = String(str);
  return text.length > maxLen ? text.substring(0, maxLen) : text;
}

function parseHeaderBlock(headerText) {
  const headers = {};
  if (!headerText) return headers;

  const unfolded = headerText.replace(/\r?\n[ \t]+/g, ' ');
  for (const line of unfolded.split(/\r?\n/)) {
    const colonIndex = line.indexOf(':');
    if (colonIndex === -1) continue;
    const key = line.slice(0, colonIndex).trim().toLowerCase();
    const value = line.slice(colonIndex + 1).trim();
    if (!headers[key]) headers[key] = value;
  }
  return headers;
}

function parseTransportHeaders(headers) {
  const parsed = parseHeaderBlock(headers);
  return {
    from: decodeRFC2047(parsed.from || ''),
    to: decodeRFC2047(parsed.to || ''),
    cc: decodeRFC2047(parsed.cc || ''),
    bcc: decodeRFC2047(parsed.bcc || ''),
    subject: decodeRFC2047(parsed.subject || ''),
    date: parsed.date || '',
  };
}

function isCcRecipient(rec) {
  const type = rec.recipType || rec.recipientType;
  return type === 'cc' || type === 2 || type === '2';
}

function isBccRecipient(rec) {
  const type = rec.recipType || rec.recipientType;
  return type === 'bcc' || type === 3 || type === '3';
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
    if (isCcRecipient(rec)) ccList.push(full);
    else if (isBccRecipient(rec)) bccList.push(full);
    else toList.push(full);
  }

  return { toList, ccList, bccList };
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

function extractBodyFromMimeContent(content, maxLen) {
  if (!content) return '';

  const headerEndIndex = content.search(/\r?\n\r?\n/);
  if (headerEndIndex === -1) return '';

  const headerText = content.slice(0, headerEndIndex);
  const bodyContent = content.slice(headerEndIndex + 2, headerEndIndex + 2 + maxLen);
  const headers = parseHeaderBlock(headerText);
  const contentType = (headers['content-type'] || '').toLowerCase();

  if (contentType.includes('multipart/')) {
    const boundaryMatch = contentType.match(/boundary="?([^";\s]+)"?/i);
    if (boundaryMatch) {
      const boundary = boundaryMatch[1];
      const parts = bodyContent.split(new RegExp(`--${boundary.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`, 'i'));

      let plainText = '';
      let htmlText = '';

      for (const part of parts) {
        if (!part || part.trim() === '--' || part.trim() === '') continue;
        const partHeaderEnd = part.search(/\r?\n\r?\n/);
        if (partHeaderEnd === -1) continue;

        const partHeaders = parseHeaderBlock(part.slice(0, partHeaderEnd));
        const partBody = part.slice(partHeaderEnd + 2);
        const partType = (partHeaders['content-type'] || '').toLowerCase();

        if (partType.includes('text/plain') && !plainText) {
          plainText = partBody.trim();
        } else if (partType.includes('text/html') && !htmlText) {
          htmlText = partBody.trim();
        }
      }

      if (plainText) return truncateText(plainText, maxLen);
      if (htmlText) return truncateText(stripHtml(htmlText), maxLen);
    }
  }

  if (contentType.includes('text/html')) {
    return truncateText(stripHtml(bodyContent), maxLen);
  }

  return truncateText(bodyContent.replace(/=\r?\n/g, '').trim(), maxLen);
}

function extractMsgBody(info, maxLen) {
  const plain = (info.body || '').trim();
  if (plain) return truncateText(plain, maxLen);

  const html = (info.bodyHtml || info.html || '').trim();
  if (html) return truncateText(stripHtml(html), maxLen);

  if (info.compressedRtf && info.compressedRtf.length > 16) {
    try {
      const decompressed = decompressRTF(info.compressedRtf);
      if (decompressed) {
        const rtfStr = Buffer.from(decompressed).toString('utf8');
        // A naive but highly effective way to strip RTF tags for search indexing:
        // This removes \commands, {} braces, and leaves just the readable text blocks.
        const plainRtf = rtfStr
          .replace(/\\[a-z]+\d*\s?/ig, ' ')
          .replace(/[{}]/g, '')
          .replace(/\s{2,}/g, ' ')
          .trim();
        if (plainRtf) return truncateText(plainRtf, maxLen);
      }
    } catch {
      // Ignore RTF decompression failures
    }
  }

  if (info.headers) {
    const fromHeaders = extractBodyFromMimeContent(info.headers, maxLen);
    if (fromHeaders) return fromHeaders;
  }

  return '';
}

function extractMsgSender(info, headerData) {
  if (info.senderEmail) {
    return info.senderName
      ? `${info.senderName} <${info.senderEmail}>`
      : info.senderEmail;
  }
  if (info.senderName) return info.senderName;
  if (headerData?.from) return headerData.from;
  return '';
}

function resolveSender(sender, recipients) {
  const trimmed = String(sender || '').trim();
  if (trimmed && !BAD_SENDERS.has(trimmed.toLowerCase())) return trimmed;
  if (recipients && String(recipients).trim()) return 'Self / Sent Item';
  return trimmed;
}

function parseFilenameMetadata(filePath, stat) {
  const ext = path.extname(filePath);
  const baseName = path.basename(filePath, ext);
  let subject = baseName;
  let sentAt = stat ? stat.mtime.getTime() : 0;

  const datePrefixMatch = baseName.match(/^(\d{8})_(\d{6})_(.*)$/);
  if (datePrefixMatch) {
    const [, yyyymmdd, hhmmss, rest] = datePrefixMatch;
    subject = rest;
    try {
      const parsedDate = new Date(
        `${yyyymmdd.slice(0, 4)}-${yyyymmdd.slice(4, 6)}-${yyyymmdd.slice(6, 8)}T` +
        `${hhmmss.slice(0, 2)}:${hhmmss.slice(2, 4)}:${hhmmss.slice(4, 6)}.000Z`
      );
      if (!isNaN(parsedDate.getTime())) sentAt = parsedDate.getTime();
    } catch {}
  } else {
    const datePrefixMatch2 = baseName.match(/^(\d{8})_(.*)$/);
    if (datePrefixMatch2) {
      const [, yyyymmdd, rest] = datePrefixMatch2;
      subject = rest;
      try {
        const parsedDate = new Date(
          `${yyyymmdd.slice(0, 4)}-${yyyymmdd.slice(4, 6)}-${yyyymmdd.slice(6, 8)}T00:00:00.000Z`
        );
        if (!isNaN(parsedDate.getTime())) sentAt = parsedDate.getTime();
      } catch {}
    }
  }

  return { subject, sentAt };
}

async function readFilePrefix(filePath, maxBytes) {
  const fileHandle = await fs.promises.open(filePath, 'r');
  try {
    const buffer = Buffer.alloc(maxBytes);
    const { bytesRead } = await fileHandle.read(buffer, 0, maxBytes, 0);
    return buffer.toString('utf8', 0, bytesRead);
  } finally {
    await fileHandle.close();
  }
}

async function parseEmlRobust(filePath, stat, options) {
  const maxBodyLen = options.bodyMaxLen || DEFAULT_BODY_MAX;
  const baseName = path.basename(filePath, path.extname(filePath));
  let subject = baseName;
  let sender = '';
  let recipients = '';
  let cc = '';
  let bcc = '';
  let sentAt = stat.mtime.getTime();
  let body = '';
  let hasAttachments = false;

  if (stat.size <= FILE_SIZE_THRESHOLD) {
    try {
      const parsed = await simpleParser(fs.createReadStream(filePath));
      subject = truncateText(parsed.subject || subject, DEFAULT_SUBJECT_MAX);
      sender = parsed.from?.text || String(parsed.from || '');
      recipients = parsed.to?.text || (Array.isArray(parsed.to?.value)
        ? parsed.to.value.map(v => v.address ? (v.name ? `${v.name} <${v.address}>` : v.address) : '').filter(Boolean).join(', ')
        : String(parsed.to || ''));
      cc = parsed.cc?.text || String(parsed.cc || '');
      bcc = parsed.bcc?.text || String(parsed.bcc || '');
      if (parsed.date && !isNaN(parsed.date.getTime())) sentAt = parsed.date.getTime();
      body = truncateText(parsed.text || stripHtml(parsed.html || ''), maxBodyLen);
      hasAttachments = !!(parsed.attachments && parsed.attachments.length > 0);
      sender = resolveSender(sender, recipients);

      return {
        subject: subject.trim(),
        sender: sender.trim(),
        recipients: recipients.trim(),
        cc: cc.trim(),
        bcc: bcc.trim(),
        sentAt,
        body,
        hasAttachments,
        filePath,
        comment: 'Parsed by robust EML parser',
      };
    } catch {
      // Fall through to lightweight header/body extraction
    }
  }

  const content = stat.size > FILE_SIZE_THRESHOLD
    ? await readFilePrefix(filePath, HEADER_READ_SIZE)
    : await fs.promises.readFile(filePath, 'utf8');

  const headerEndIndex = content.search(/\r?\n\r?\n/);
  const headerText = headerEndIndex !== -1 ? content.slice(0, headerEndIndex) : content;
  const headers = parseHeaderBlock(headerText);

  subject = truncateText(decodeRFC2047(headers.subject || '') || subject, DEFAULT_SUBJECT_MAX);
  sender = decodeRFC2047(headers.from || '');
  recipients = decodeRFC2047(headers.to || '');
  cc = decodeRFC2047(headers.cc || '');
  bcc = decodeRFC2047(headers.bcc || '');

  if (headers.date) {
    const parsedDate = new Date(headers.date);
    if (!isNaN(parsedDate.getTime())) sentAt = parsedDate.getTime();
  }

  hasAttachments = /Content-Disposition:\s*attachment/i.test(content) ||
    /Content-Type:\s*[^;\s]+;\s*name=/i.test(content);

  body = extractBodyFromMimeContent(content, maxBodyLen);
  sender = resolveSender(sender, recipients);

  return {
    subject: subject.trim(),
    sender: sender.trim(),
    recipients: recipients.trim(),
    cc: cc.trim(),
    bcc: bcc.trim(),
    sentAt,
    body,
    hasAttachments,
    filePath,
    comment: 'Parsed by robust EML fallback',
  };
}

async function parseMsgRobust(filePath, stat, options) {
  const maxBodyLen = options.bodyMaxLen || DEFAULT_BODY_MAX;
  const filenameMeta = parseFilenameMetadata(filePath, stat);
  let subject = filenameMeta.subject;
  let sender = '';
  let recipients = '';
  let cc = '';
  let bcc = '';
  let sentAt = filenameMeta.sentAt;
  let body = '';
  let hasAttachments = false;

  if (stat.size > FILE_SIZE_THRESHOLD) {
    return {
      subject,
      sender: resolveSender('', recipients),
      recipients,
      cc,
      bcc,
      sentAt,
      body,
      hasAttachments,
      filePath,
      comment: 'Large MSG — filename metadata only',
    };
  }

  const buffer = await fs.promises.readFile(filePath);

  const textPreview = buffer.toString('utf8', 0, 500).trim();
  if (textPreview.startsWith('{')) {
    try {
      const parsed = JSON.parse(buffer.toString('utf8'));
      if (parsed.internetMessageId || parsed.subject || parsed.sentAt) {
        recipients = Array.isArray(parsed.to) ? parsed.to.join(', ') : String(parsed.to || '');
        sender = resolveSender(parsed.sender || '', recipients);
        return {
          subject: truncateText(parsed.subject || subject, DEFAULT_SUBJECT_MAX),
          sender,
          recipients,
          cc: Array.isArray(parsed.cc) ? parsed.cc.join(', ') : String(parsed.cc || ''),
          bcc: '',
          sentAt: parsed.sentAt ? new Date(parsed.sentAt).getTime() : sentAt,
          body: truncateText(parsed.bodyPreview || parsed.body || '', maxBodyLen),
          hasAttachments: !!parsed.hasAttachments,
          filePath,
          comment: 'Parsed JSON email record',
        };
      }
    } catch {
      // Not JSON — continue with MsgReader
    }
  }

  try {
    const reader = new MsgReader(buffer);
    const info = reader.getFileData();

    if (info.error) {
      if (info.error.includes('Unsupported file type')) {
        return parseEmlRobust(filePath, stat, options);
      }
      throw new Error(info.error);
    }

    const headerData = info.headers ? parseTransportHeaders(info.headers) : null;

    subject = truncateText(info.subject || headerData?.subject || subject, DEFAULT_SUBJECT_MAX);
    sender = extractMsgSender(info, headerData);

    let { toList, ccList, bccList } = splitRecipientsByType(info.recipients);
    if (toList.length === 0 && headerData?.to) toList = [headerData.to];
    if (ccList.length === 0 && headerData?.cc) ccList = [headerData.cc];
    if (bccList.length === 0 && headerData?.bcc) bccList = [headerData.bcc];

    recipients = toList.join(', ');
    cc = ccList.join(', ');
    bcc = bccList.join(', ');

    sentAt = extractMsgTimestamp(info);
    if (!sentAt && headerData?.date) {
      const headerDate = new Date(headerData.date);
      if (!isNaN(headerDate.getTime())) sentAt = headerDate.getTime();
    }

    hasAttachments = !!(info.attachments && info.attachments.length > 0);
    body = extractMsgBody(info, maxBodyLen);
    sender = resolveSender(sender, recipients);

    return {
      subject: subject.trim(),
      sender: sender.trim(),
      recipients: recipients.trim(),
      cc: cc.trim(),
      bcc: bcc.trim(),
      sentAt,
      body,
      hasAttachments,
      filePath,
      comment: 'Parsed by robust MSG parser',
    };
  } catch (err) {
    return {
      subject,
      sender: resolveSender('', recipients),
      recipients,
      cc,
      bcc,
      sentAt,
      body,
      hasAttachments,
      filePath,
      comment: `MSG parse fallback: ${err.message}`,
    };
  }
}

/**
 * Parse an email file with maximum metadata/body recovery.
 * @param {string} filePath
 * @param {{ bodyMaxLen?: number }} [options]
 */
async function parseRobustEmailFile(filePath, options = {}) {
  const ext = path.extname(filePath).toLowerCase();
  const stat = await fs.promises.stat(filePath);

  if (ext === '.eml') {
    return parseEmlRobust(filePath, stat, options);
  }
  if (ext === '.msg') {
    return parseMsgRobust(filePath, stat, options);
  }

  throw new Error(`Unsupported file extension: ${ext}`);
}

function isBadSender(sender) {
  return BAD_SENDERS.has(String(sender || '').trim().toLowerCase());
}

function isEmptyField(value) {
  return !value || !String(value).trim();
}

/**
 * Returns true when an indexed document should be re-parsed.
 */
function needsReindex(doc) {
  return isBadSender(doc.sender) ||
    isEmptyField(doc.recipients) ||
    isEmptyField(doc.body);
}

/**
 * Build a Meilisearch patch with only improved fields.
 */
function buildReindexPatch(doc, parsed) {
  const patch = { id: doc.id };
  let changed = false;

  if (isBadSender(doc.sender) && parsed.sender && !isBadSender(parsed.sender)) {
    patch.sender = parsed.sender;
    changed = true;
  }

  if (isEmptyField(doc.recipients) && parsed.recipients) {
    patch.recipients = parsed.recipients;
    changed = true;
  }

  if (isEmptyField(doc.cc) && parsed.cc) {
    patch.cc = parsed.cc;
    changed = true;
  }

  if (isEmptyField(doc.body) && parsed.body) {
    patch.body = parsed.body;
    changed = true;
  }

  if ((!doc.sentAt || Number(doc.sentAt) === 0) && parsed.sentAt > 0) {
    patch.sentAt = parsed.sentAt;
    changed = true;
  }

  const currentSubject = String(doc.subject || '').trim();
  const parsedSubject = String(parsed.subject || '').trim();
  if (parsedSubject && (isEmptyField(currentSubject) || currentSubject === path.basename(doc.filePath || '', path.extname(doc.filePath || '')))) {
    if (parsedSubject !== currentSubject) {
      patch.subject = parsedSubject;
      changed = true;
    }
  }

  if (parsed.hasAttachments && !doc.hasAttachments) {
    patch.hasAttachments = true;
    changed = true;
  }

  return changed ? patch : null;
}

module.exports = {
  parseRobustEmailFile,
  needsReindex,
  buildReindexPatch,
  isBadSender,
  isEmptyField,
  decodeRFC2047,
  stripHtml,
  extractMsgBody,
  extractMsgSender,
  parseTransportHeaders,
  splitRecipientsByType,
  FILE_SIZE_THRESHOLD,
  DEFAULT_BODY_MAX,
};
