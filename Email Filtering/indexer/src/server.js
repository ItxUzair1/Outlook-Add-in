const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const { XMLParser } = require('fast-xml-parser');

const state = require('./state');
const uploader = require('./uploader');
require('dotenv').config();

const { MeiliSearch } = require('meilisearch');
const meiliClient = new MeiliSearch({
  host: process.env.MEILI_URL || 'http://localhost:7700',
  apiKey: process.env.MEILI_MASTER_KEY,
});
const emailIndex = meiliClient.index('emails');

const app = express();
const PORT = process.env.PORT || 4001;

// Middleware
app.use(cors());
app.use(express.json());

// Set up multer memory storage for file uploads
const upload = multer({ storage: multer.memoryStorage() });

// Helper to resolve native executable path
function getExecutablePath(exeName) {
  // Try relative to workspace structure
  const paths = [
    path.join(__dirname, '..', '..', 'backend', 'bin', exeName), // relative to indexer/src/
    path.join(process.cwd(), '..', 'backend', 'bin', exeName),
    path.join(__dirname, '..', 'bin', exeName),
    path.join(process.cwd(), 'bin', exeName),
  ];

  for (const p of paths) {
    if (fs.existsSync(p)) {
      return p;
    }
  }
  return null;
}

/**
 * Helper to parse mmcollection XML content into location objects
 */
function parseCollectionXml(xmlContent) {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_"
  });
  const result = parser.parse(xmlContent);

  const locations = [];
  if (result?.mailmanager?.locations?.store) {
    const stores = Array.isArray(result.mailmanager.locations.store) 
      ? result.mailmanager.locations.store 
      : [result.mailmanager.locations.store];

    const getStr = (val) => {
      if (val === null || val === undefined) return "";
      if (typeof val === "object") return ""; // fast-xml-parser empty XML tag behavior
      return String(val);
    };

    for (const store of stores) {
      locations.push({
        id: getStr(store["@_id"]),
        type: getStr(store.type) || "msg",
        description: getStr(store.description),
        folder: getStr(store.folder),
        isSuggested: store["@_isSuggested"] === "true" || store["@_isSuggested"] === true,
        isUnused: store["@_isUnused"] === "true" || store["@_isUnused"] === true,
      });
    }
  }
  return locations;
}

// --- Endpoints ---

// 1. Get State
app.get('/api/state', (req, res) => {
  try {
    const s = state.loadState();
    res.json(s);
  } catch (err) {
    res.status(500).json({ error: 'Failed to load indexer state', details: err.message });
  }
});

// 2. Add Location
app.post('/api/state/folders', (req, res) => {
  const { path: folderPath, type, description } = req.body;
  if (!folderPath) {
    return res.status(400).json({ error: 'path is required' });
  }

  // Connectivity and type check
  try {
    if (!fs.existsSync(folderPath)) {
      return res.status(400).json({ error: `Path does not exist or is inaccessible: ${folderPath}` });
    }
    const stat = fs.statSync(folderPath);
    if (!stat.isDirectory()) {
      return res.status(400).json({ error: `Path must be a directory, not a file: ${folderPath}` });
    }
  } catch (err) {
    return res.status(400).json({ error: `Cannot access path (${err.message})` });
  }
  
  try {
    const added = state.addFolder(folderPath, type, description);
    res.json({ success: true, added, state: state.loadState() });
  } catch (err) {
    res.status(500).json({ error: 'Failed to add folder', details: err.message });
  }
});

// 3. Remove Location
app.delete('/api/state/folders', (req, res) => {
  const { path: folderPath } = req.body;
  if (!folderPath) {
    return res.status(400).json({ error: 'path is required in request body' });
  }
  
  try {
    const removed = state.removeFolder(folderPath);
    res.json({ success: true, removed, state: state.loadState() });
  } catch (err) {
    res.status(500).json({ error: 'Failed to remove folder', details: err.message });
  }
});

// 4. Indexer Controls
app.post('/api/indexer/start', (req, res) => {
  try {
    uploader.start();
    res.json({ success: true, status: 'started' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to start indexer', details: err.message });
  }
});

app.post('/api/indexer/pause', (req, res) => {
  try {
    uploader.pause();
    res.json({ success: true, status: 'paused' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to pause indexer', details: err.message });
  }
});

app.post('/api/indexer/reset', (req, res) => {
  try {
    uploader.reset();
    res.json({ success: true, status: 'reset' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to reset progress', details: err.message });
  }
});

app.post('/api/scheduler/start', (req, res) => {
  try {
    uploader.startScheduler();
    res.json({ success: true, status: 'scheduler_active' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to start scheduler', details: err.message });
  }
});

app.post('/api/scheduler/stop', (req, res) => {
  try {
    uploader.stopScheduler();
    res.json({ success: true, status: 'scheduler_inactive' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to stop scheduler', details: err.message });
  }
});

// 5. Native Browsers

// Open Windows Folder Picker Dialog
app.get('/api/browse-folder', (req, res) => {
  const exePath = getExecutablePath('koyobrowse.exe');
  if (!exePath) {
    state.addLog('koyobrowse.exe folder picker utility not found');
    return res.status(500).json({ error: 'koyobrowse.exe utility not found' });
  }
  
  let cmd = `"${exePath}" "Select Directory to Index"`;
  if (req.query.startPath) {
    cmd += ` "${req.query.startPath}"`;
  }
  
  exec(cmd, { timeout: 120000 }, (error, stdout, stderr) => {
    if (error && error.killed) {
      return res.status(500).json({ error: 'Folder picker dialog timed out' });
    }
    if (error && !stdout.trim()) {
      return res.status(500).json({ error: 'Folder dialog cancelled or failed', details: stderr || error.message });
    }
    const selectedPath = stdout.trim();
    res.json({ path: selectedPath });
  });
});

// Open Windows File Picker for .mmcollection
app.get('/api/browse-file', (req, res) => {
  const exePath = getExecutablePath('koyofile.exe');
  if (!exePath) {
    state.addLog('koyofile.exe file picker utility not found');
    return res.status(500).json({ error: 'koyofile.exe utility not found' });
  }
  
  exec(`"${exePath}"`, { timeout: 120000 }, (error, stdout, stderr) => {
    if (error && error.killed) {
      return res.status(500).json({ error: 'File picker dialog timed out' });
    }
    if (error && !stdout.trim()) {
      return res.status(500).json({ error: 'File dialog cancelled or failed', details: stderr || error.message });
    }
    const selectedPath = stdout.trim();
    res.json({ path: selectedPath });
  });
});

// 6. Collections API

// Get all active collections from state
app.get('/api/active-collections', (req, res) => {
  try {
    const folders = state.getFolders();
    const collectionIds = [...new Set(folders.map(f => f.collectionId).filter(Boolean))];
    res.json({ collections: collectionIds });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Load mmcollection file path
app.post('/api/collections/load', (req, res) => {
  const { filePath } = req.body;
  if (!filePath) {
    return res.status(400).json({ error: 'filePath is required' });
  }
  
  try {
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'Collection file not found at specified path' });
    }
    
    const xmlContent = fs.readFileSync(filePath, 'utf8');
    const locations = parseCollectionXml(xmlContent);
    
    state.addLog(`Loaded ${locations.length} location paths from collection file: ${filePath}`);
    res.json({ locations, filePath, collectionName: path.basename(filePath, '.mmcollection') });
  } catch (err) {
    res.status(500).json({ error: 'Failed to read collection file', details: err.message });
  }
});

// Direct Web Upload of mmcollection file
app.post('/api/collections/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }
  
  try {
    const xmlContent = req.file.buffer.toString('utf8');
    const locations = parseCollectionXml(xmlContent);
    const collectionName = path.basename(req.file.originalname, '.mmcollection');
    
    state.addLog(`Uploaded and parsed collection "${collectionName}" with ${locations.length} locations`);
    res.json({ locations, collectionName });
  } catch (err) {
    res.status(500).json({ error: 'Failed to parse uploaded collection file', details: err.message });
  }
});


// Server Initialization
app.listen(PORT, () => {
  console.log(`==========================================`);
  console.log(` Koyomail Admin Indexer Backend API Server`);
  console.log(` Running on http://localhost:${PORT}`);
  console.log(`==========================================`);
  state.addLog(`Indexer API server started on port ${PORT}`);
});
