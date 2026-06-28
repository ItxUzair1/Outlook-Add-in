const fs = require('fs');
const path = require('path');
const { simpleParser } = require('mailparser');
const MsgReader = require('msgreader').default; // Using the standard package

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
    subject: parsed.subject || '',
    sender: parsed.from?.text || '',
    recipients: parsed.to?.text || '',
    cc: parsed.cc?.text || '',
    bcc: parsed.bcc?.text || '',
    sentAt: parsed.date ? parsed.date.getTime() : 0, // Store as timestamp for sorting
    body: parsed.text || parsed.html || '', // Prefer plain text
    hasAttachments: parsed.attachments && parsed.attachments.length > 0,
    filePath: filePath,
    comment: '' // Comments will be filled if there's a sidecar file later, or left empty
  };
}

async function parseMsg(filePath) {
  const buffer = fs.readFileSync(filePath);
  
  // Check if it's actually a JSON file disguised as .msg
  const textPreview = buffer.toString('utf8', 0, 500).trim();
  if (textPreview.startsWith('{')) {
    try {
      const parsed = JSON.parse(buffer.toString('utf8'));
      if (parsed.internetMessageId || parsed.subject || parsed.sentAt) {
        return {
          subject: parsed.subject || '',
          sender: parsed.sender || '',
          recipients: (parsed.to || []).join(', '),
          cc: (parsed.cc || []).join(', '),
          bcc: '',
          sentAt: parsed.sentAt ? new Date(parsed.sentAt).getTime() : 0,
          body: parsed.bodyPreview || parsed.body || '',
          hasAttachments: !!parsed.hasAttachments,
          filePath: filePath,
          comment: ''
        };
      }
    } catch (e) {
      // Not a valid JSON file, fallback to MsgReader
    }
  }

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

  // Parse dates correctly
  let timestamp = 0;
  if (parsed.date) {
    // msgreader sometimes returns date strings that need parsing
    const dateObj = new Date(parsed.date);
    if (!isNaN(dateObj.getTime())) {
      timestamp = dateObj.getTime();
    }
  }

  return {
    subject: parsed.subject || '',
    sender: parsed.senderName || parsed.senderEmail || '',
    recipients: parsed.recipients?.filter(r => r.recipType === 'to').map(r => r.name || r.email).join(', ') || '',
    cc: parsed.recipients?.filter(r => r.recipType === 'cc').map(r => r.name || r.email).join(', ') || '',
    bcc: parsed.recipients?.filter(r => r.recipType === 'bcc').map(r => r.name || r.email).join(', ') || '',
    sentAt: timestamp,
    body: parsed.body || '', // msgreader provides plain text body
    hasAttachments: parsed.attachments && parsed.attachments.length > 0,
    filePath: filePath,
    comment: ''
  };
}

module.exports = {
  parseEmailFile
};
