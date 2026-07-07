const express = require('express');
const cors = require('cors');
const multer = require('multer');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const { XMLParser } = require('fast-xml-parser');

const state = require('./state');
const uploader = require('./uploader');
const { runMeiliDiagnostics } = require('./meiliDiagnostics');
const { runRepair: runMetadataRepair } = require('./repairMetadata');
const { runRetryErrors } = require('./retryErrors');
const pkg = require('../package.json');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

let electronDialog = null;
let electronBrowserWindow = null;
try {
  const electron = require('electron');
  electronDialog = electron.dialog;
  electronBrowserWindow = electron.BrowserWindow;
} catch (err) {
  console.log('Electron not available, native dialogs disabled.');
}

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

app.get('/api/version', (req, res) => {
  res.json({ version: pkg.version, name: pkg.name });
});

// Meilisearch connection + local vs remote document count check
app.get('/api/diagnostics', async (req, res) => {
  try {
    const report = await runMeiliDiagnostics({ state, pkg });
    res.json(report);
  } catch (err) {
    res.status(500).json({
      ok: false,
      error: 'Diagnostics failed',
      details: err.message,
    });
  }
});

// 1. Get State (excludes uploaded file ledger — can be millions of paths)
app.get('/api/state', (req, res) => {
  try {
    res.json(state.getPublicState());
  } catch (err) {
    res.status(500).json({ error: 'Failed to load indexer state', details: err.message });
  }
});

// 2. Add Location
app.post('/api/state/folders', (req, res) => {
  let { path: folderPath, type, description } = req.body;
  if (!folderPath) {
    return res.status(400).json({ error: 'path is required' });
  }

  // Expand Windows environment variables (e.g. %USERPROFILE%)
  folderPath = folderPath.replace(/%([^%]+)%/g, (match, n) => {
    const key = Object.keys(process.env).find(k => k.toLowerCase() === n.toLowerCase());
    return key ? process.env[key] : match;
  });

  // Connectivity and type check
  try {
    if (type !== 'collection' && !fs.existsSync(folderPath)) {
      return res.status(400).json({ error: `Path does not exist or is inaccessible: ${folderPath}` });
    }
    if (fs.existsSync(folderPath)) {
      const stat = fs.statSync(folderPath);
      if (!stat.isDirectory()) {
        return res.status(400).json({ error: `Path must be a directory, not a file: ${folderPath}` });
      }
    }
  } catch (err) {
    if (type !== 'collection') {
      return res.status(400).json({ error: `Cannot access path (${err.message})` });
    }
  }
  
  try {
    const added = state.addFolder(folderPath, type, description);
    res.json({ success: true, added, state: state.getPublicState() });
  } catch (err) {
    res.status(500).json({ error: 'Failed to add folder', details: err.message });
  }
});

// 3. Remove Location
app.delete('/api/state/folders', (req, res) => {
  const { path: folderPath } = req.body;
  if (!folderPath) {
    return res.status(400).json({ error: 'path is required' });
  }

  const success = state.removeFolder(folderPath);
  if (success) {
    res.json({ message: 'Removed successfully', state: state.getPublicState() });
  } else {
    res.status(404).json({ error: 'Folder not found' });
  }
});

// 4. Update Folder Permissions
app.put('/api/state/folders/permissions', (req, res) => {
  const { path: folderPath, isPublic, allowedUsers } = req.body;
  
  if (!folderPath) {
    return res.status(400).json({ error: 'path is required' });
  }
  
  const success = state.updateFolderPermissions(folderPath, isPublic, allowedUsers);
  if (success) {
    res.json({ message: 'Permissions updated successfully', state: state.getPublicState() });
  } else {
    res.status(404).json({ error: 'Folder not found' });
  }
});

// 5. Indexer Controls
app.post('/api/indexer/start', (req, res) => {
  try {
    const { targetPaths } = req.body || {};
    uploader.start(targetPaths);
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
    const { folders = [] } = req.body;
    uploader.reset(folders);
    res.json({ success: true, status: 'reset', targeted: folders.length > 0 });
  } catch (err) {
    res.status(500).json({ error: 'Failed to reset progress', details: err.message });
  }
});

app.post('/api/indexer/repair-metadata', (req, res) => {
  try {
    const s = state.loadState();
    if (s.indexingStatus === 'scanning' || s.indexingStatus === 'uploading' || s.indexingStatus === 'repairing') {
      return res.status(409).json({ error: 'Indexer or repair is already running. Wait for it to finish.' });
    }

    state.updateIndexingStatus('repairing');
    state.updateStats({
      totalFilesFound: 0,
      filesIndexedThisSession: 0,
      filesSkipped: 0,
      currentFilePath: 'Preparing metadata repair...',
      speed: 0,
    }, { immediate: true });
    state.addLog('Starting metadata repair: fixing missing To / Cc / Date fields...');

    (async () => {
      try {
        const result = await runMetadataRepair({
          log: (msg) => state.addLog(msg),
          onProgress: ({ total, scanned, repaired, skipped, currentFilePath }) => {
            state.updateStats({
              totalFilesFound: total || scanned,
              filesIndexedThisSession: scanned,
              filesSkipped: skipped,
              currentFilePath: currentFilePath || `Checking emails... (${scanned}${total ? ` / ${total}` : ''}, ${repaired} fixed)`,
            }, { persist: scanned % 50 === 0 });
          },
          shouldStop: () => state.getIndexingStatus() === 'paused',
        });
        state.addLog(
          `Metadata repair finished — ${result.repaired} emails updated, ${result.skipped} already OK.` +
          (result.stopped ? ' (stopped early)' : '')
        );
      } catch (err) {
        console.error('Metadata repair error:', err);
        state.addLog(`Metadata repair failed: ${err.message}`);
      } finally {
        state.updateStats({ currentFilePath: '', speed: 0 }, { immediate: true });
        state.updateIndexingStatus('idle');
      }
    })();

    res.json({ success: true, status: 'started' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to start metadata repair', details: err.message });
  }
});

app.post('/api/indexer/retry-errors', async (req, res) => {
  try {
    const s = state.loadState();
    if (s.indexingStatus === 'scanning' || s.indexingStatus === 'uploading' || s.indexingStatus === 'repairing' || s.indexingStatus === 'retrying') {
      return res.status(409).json({ error: 'Indexer, repair or error recovery is already running. Wait for it to finish.' });
    }

    // Run recovery asynchronously so endpoint returns immediately and doesn't block the UI
    runRetryErrors().catch(err => {
      console.error('Error recovery task failed:', err);
      state.addLog(`Error recovery failed: ${err.message}`);
      state.updateIndexingStatus('idle');
    });

    res.json({ success: true, status: 'started' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to start error recovery', details: err.message });
  }
});

app.post('/api/indexer/fast-sync', (req, res) => {
  try {
    const s = state.loadState();
    const folders = s.folders || [];
    
    state.updateIndexingStatus('uploading');
    state.addLog('Starting Fast Sync: Syncing permissions for all folders...');
    
    (async () => {
      let totalUpdated = 0;
      for (const folder of folders) {
        let offset = 0;
        const limit = 1000;
        let hasMore = true;
        
        while (hasMore) {
          try {
            const searchResponse = await emailIndex.search('', {
              filter: `indexedRootPath = "${folder.path.replace(/\\/g, '\\\\')}"`,
              limit,
              offset,
              attributesToRetrieve: ['id']
            });
            
            if (searchResponse.hits.length === 0) {
              hasMore = false;
              break;
            }
            
            const isPublic = folder.isPublic !== false;
            const allowedUsers = (folder.allowedUsers || []).map(u => u.toLowerCase());
            
            const updatePayload = searchResponse.hits.map(hit => ({
              id: hit.id,
              isPublic,
              allowedUsers
            }));
            
            await emailIndex.updateDocuments(updatePayload);
            totalUpdated += updatePayload.length;
            
            if (searchResponse.hits.length < limit) {
              hasMore = false;
            } else {
              offset += limit;
            }
          } catch (err) {
            console.error(`Fast Sync error on folder ${folder.path}:`, err);
            state.addLog(`Error syncing ${folder.path}: ${err.message}`);
            hasMore = false;
          }
        }
      }
      
      state.updateIndexingStatus('idle');
      state.addLog(`Fast Sync completed successfully. Updated ${totalUpdated} documents in Meilisearch.`);
    })();
    
    res.json({ success: true, status: 'started' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to start fast sync', details: err.message });
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
app.get('/api/browse-folder', async (req, res) => {
  if (!electronDialog) {
    return res.status(500).json({ error: 'Native dialogs only available when running via Electron' });
  }
  
  try {
    const win = electronBrowserWindow.getFocusedWindow();
    const result = await electronDialog.showOpenDialog(win, {
      title: 'Select Directory to Index',
      properties: ['openDirectory']
    });
    
    if (!result.canceled && result.filePaths.length > 0) {
      res.json({ path: result.filePaths[0] });
    } else {
      res.status(400).json({ error: 'Folder dialog cancelled' });
    }
  } catch (error) {
    res.status(500).json({ error: 'Folder picker failed', details: error.message });
  }
});

// Open Windows File Picker for .mmcollection
app.get('/api/browse-file', async (req, res) => {
  if (!electronDialog) {
    return res.status(500).json({ error: 'Native dialogs only available when running via Electron' });
  }
  
  try {
    const win = electronBrowserWindow.getFocusedWindow();
    const result = await electronDialog.showOpenDialog(win, {
      title: 'Select .mmcollection File',
      properties: ['openFile'],
      filters: [{ name: 'MailManager Collections', extensions: ['mmcollection'] }]
    });
    
    if (!result.canceled && result.filePaths.length > 0) {
      res.json({ path: result.filePaths[0] });
    } else {
      res.status(400).json({ error: 'File dialog cancelled' });
    }
  } catch (error) {
    res.status(500).json({ error: 'File picker failed', details: error.message });
  }
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

// Admin Login
app.post('/api/admin/login', (req, res) => {
  const { email, password } = req.body;
  
  const expectedEmail = process.env.ADMIN_EMAIL;
  const expectedPassword = process.env.ADMIN_PASSWORD;

  if (email === expectedEmail && password === expectedPassword) {
    res.json({ token: 'koyo-admin-token-123', success: true });
  } else {
    res.status(401).json({ error: 'Invalid credentials', success: false });
  }
});

// Serve the React Admin Dashboard (built files from indexer/public)
const publicPath = path.join(__dirname, '..', 'public');
if (fs.existsSync(publicPath)) {
  app.use(express.static(publicPath));
  // Catch-all to serve index.html for React Router (Express v5 compatible)
  app.get('/{*path}', (req, res) => {
    res.sendFile(path.join(publicPath, 'index.html'));
  });
}


// Server Initialization
app.listen(PORT, async () => {
  console.log(`==========================================`);
  console.log(` Koyomail Admin Indexer Backend API Server`);
  console.log(` Running on http://localhost:${PORT}`);
  console.log(` Diagnostics: http://localhost:${PORT}/api/diagnostics`);
  console.log(`==========================================`);

  try {
    const diag = await runMeiliDiagnostics({ state, pkg });
    const host = diag.meilisearch.configuredHost;
    const docs = diag.meilisearch.documentCount;
    const connected = diag.meilisearch.connected ? 'connected' : 'NOT connected';
    console.log(` Meilisearch: ${host} (${connected}, ${docs ?? '?'} documents)`);
    if (diag.meilisearch.usingLocalhostFallback) {
      console.log(` WARNING: MEILI_URL missing — using localhost fallback!`);
    }
    if (diag.local?.documentCountMismatch) {
      console.log(
        ` WARNING: Local indexed (${diag.local.filesIndexed}) != Meilisearch (${docs})`
      );
    }
  } catch (err) {
    console.log(` Meilisearch diagnostics failed: ${err.message}`);
  }

  state.addLog(`Indexer API server started on port ${PORT}`);
  
  // Auto-resume Live Scheduler if it was active
  const currentState = state.loadState();
  if (currentState.schedulerStatus === 'active') {
    uploader.startScheduler(true);
  }
});
