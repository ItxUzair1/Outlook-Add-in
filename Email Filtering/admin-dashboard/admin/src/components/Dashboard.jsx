import { useState, useEffect } from 'react';
import Header from './Header';
import MetricsRow from './MetricsRow';
import AddLocationPanel from './AddLocationPanel';
import FoldersTable from './FoldersTable';
import IndexingControls from './IndexingControls';
import LogViewer from './LogViewer';
import Toast from './Toast';
import ErrorModal from './ErrorModal';

const API_BASE_URL = 'http://localhost:4001/api';

export default function Dashboard({ onLogout }) {
  // Dashboard State
  const [folders, setFolders] = useState([]);
  const [indexingStatus, setIndexingStatus] = useState('idle');
  const [schedulerStatus, setSchedulerStatus] = useState('inactive');
  const [selectedFolders, setSelectedFolders] = useState([]);
  const [stats, setStats] = useState({
    totalFilesFound: 0,
    filesIndexed: 0,
    filesFailed: 0,
    currentFilePath: '',
    speed: 0
  });
  const [logs, setLogs] = useState([]);
  const [recentErrors, setRecentErrors] = useState([]);
  const [showErrorModal, setShowErrorModal] = useState(false);
  const [isUploadingCollection, setIsUploadingCollection] = useState(false);
  const [toast, setToast] = useState({ message: '', type: 'success' });
  const [appVersion, setAppVersion] = useState('');

  const showToast = (message, type = 'success') => setToast({ message, type });

  useEffect(() => {
    fetch(`${API_BASE_URL}/version`)
      .then(resp => resp.ok ? resp.json() : null)
      .then(data => { if (data?.version) setAppVersion(data.version); })
      .catch(() => {});
  }, []);

  // Poll state — faster while indexing for live updates
  useEffect(() => {
    fetchState();

    const isActive = indexingStatus === 'uploading' || indexingStatus === 'scanning';
    const pollMs = isActive ? 500 : 1500;
    const interval = setInterval(() => {
      fetchState();
    }, pollMs);

    return () => clearInterval(interval);
  }, [indexingStatus]);

  async function fetchState() {
    try {
      const resp = await fetch(`${API_BASE_URL}/state`);
      if (resp.ok) {
        const data = await resp.json();
        setFolders(data.folders || []);
        setIndexingStatus(data.indexingStatus || 'idle');
        setSchedulerStatus(data.schedulerStatus || 'inactive');
        setStats(data.stats || {});
        setLogs(data.logs || []);
        setRecentErrors(data.recentErrors || []);
      }
    } catch (err) {
      console.error('Failed to connect to indexer API:', err);
    }
  }

  // API Actions
  const handleAddManualFolder = async (pathToAdd) => {
    try {
      const resp = await fetch(`${API_BASE_URL}/state/folders`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          path: pathToAdd,
          type: 'local',
          description: pathToAdd.split(/[\\/]/).pop() || pathToAdd
        })
      });

      if (resp.ok) {
        const data = await resp.json();
        if (data.added === false) {
          showToast(`Folder already exists in the list!`, 'warning');
        } else {
          showToast(`Folder added successfully!`);
        }
        fetchState();
      } else {
        const errData = await resp.json();
        showToast(`Failed to add folder: ${errData.error}`, 'error');
      }
    } catch (err) {
      showToast('Error adding folder', 'error');
      console.error(err);
    }
  };

  const handleRemoveFolder = async (folderPath) => {
    try {
      const resp = await fetch(`${API_BASE_URL}/state/folders`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path: folderPath })
      });
      if (resp.ok) {
        showToast('Folder removed successfully');
        fetchState();
      } else {
        showToast('Failed to remove folder', 'error');
      }
    } catch (err) {
      showToast('Error removing folder', 'error');
      console.error(err);
    }
  };

  const handleStartIndexing = async () => {
    try {
      await fetch(`${API_BASE_URL}/indexer/start`, { 
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ targetPaths: selectedFolders })
      });
      showToast(selectedFolders.length > 0 ? 'Started targeted indexing' : 'Started indexing all folders');
      fetchState();
    } catch (err) {
      showToast('Failed to start indexing', 'error');
      console.error(err);
    }
  };

  const handlePause = async () => {
    try {
      await fetch(`${API_BASE_URL}/indexer/pause`, { method: 'POST' });
      showToast('Indexing paused', 'warning');
      fetchState();
    } catch (err) {
      showToast('Failed to pause indexer', 'error');
      console.error(err);
    }
  };

  const handleReset = async () => {
    if (window.confirm('Are you sure you want to reset the indexing progress? This will clear Meilisearch uploaded status but keep your folders.')) {
      try {
        await fetch(`${API_BASE_URL}/indexer/reset`, { method: 'POST' });
        showToast('Indexer progress reset successfully');
        fetchState();
      } catch (err) {
        showToast('Failed to reset indexer', 'error');
        console.error(err);
      }
    }
  };

  const handleFastSync = async () => {
    try {
      await fetch(`${API_BASE_URL}/indexer/fast-sync`, { method: 'POST' });
      showToast('Fast Sync started! Check console for progress.', 'success');
      fetchState();
    } catch (err) {
      showToast('Failed to start Fast Sync', 'error');
      console.error(err);
    }
  };

  const handleRepairMetadata = async () => {
    if (!window.confirm(
      'Repair missing To / Cc / Date fields for already-indexed emails?\n\n' +
      'This reads email files from this PC and updates the search index. ' +
      'It does NOT re-index everything and usually takes a few minutes.'
    )) {
      return;
    }

    try {
      const resp = await fetch(`${API_BASE_URL}/indexer/repair-metadata`, { method: 'POST' });
      if (resp.ok) {
        showToast('Metadata repair started — watch the log below for progress.', 'success');
        fetchState();
      } else {
        const data = await resp.json();
        showToast(data.error || 'Failed to start metadata repair', 'error');
      }
    } catch (err) {
      showToast('Failed to start metadata repair', 'error');
      console.error(err);
    }
  };

  const handleStartScheduler = async () => {
    try {
      await fetch(`${API_BASE_URL}/scheduler/start`, { method: 'POST' });
      showToast('Live Scheduler started');
      fetchState();
    } catch (err) {
      showToast('Failed to start Live Scheduler', 'error');
      console.error(err);
    }
  };

  const handleStopScheduler = async () => {
    try {
      await fetch(`${API_BASE_URL}/scheduler/stop`, { method: 'POST' });
      showToast('Live Scheduler stopped', 'warning');
      fetchState();
    } catch (err) {
      showToast('Failed to stop Live Scheduler', 'error');
      console.error(err);
    }
  };

  const handleBrowseFolder = async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/browse-folder`);
      if (resp.ok) {
        const data = await resp.json();
        return data.path;
      }
    } catch (err) {
      console.error('Folder picker browse failed:', err);
      showToast('Failed to open folder picker', 'error');
    }
    return '';
  };

  const handleBrowseCollectionFile = async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/browse-file`);
      if (resp.ok) {
        const data = await resp.json();
        if (data.path) {
          await loadCollectionFromFilePath(data.path);
          return data.path;
        }
      }
    } catch (err) {
      console.error('File picker browse failed:', err);
      showToast('Failed to open file picker', 'error');
    }
    return '';
  };

  const loadCollectionFromFilePath = async (filePath) => {
    setIsUploadingCollection(true);
    try {
      const resp = await fetch(`${API_BASE_URL}/collections/load`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filePath })
      });

      if (resp.ok) {
        const data = await resp.json();
        const locations = data.locations || [];
        
        let addedCount = 0;
        let errors = [];
        for (const loc of locations) {
          const pathToAdd = loc.folder || loc.path;
          if (pathToAdd) {
            const addResp = await fetch(`${API_BASE_URL}/state/folders`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                path: pathToAdd,
                type: 'collection',
                description: loc.description || pathToAdd.split(/[\\/]/).pop() || pathToAdd
              })
            });
            if (addResp.ok) {
              const addData = await addResp.json();
              if (addData.added !== false) addedCount++;
            } else {
              const errData = await addResp.json();
              errors.push(errData.error);
            }
          }
        }
        
        if (addedCount > 0) {
          showToast(`Collection loaded successfully! Added ${addedCount} folders.`);
        } else if (errors.length > 0) {
          showToast(`No folders added. Example error: ${errors[0]}`, 'error');
        } else {
          showToast('No valid folders found in collection.', 'warning');
        }
        
        fetchState();
      } else {
        const errData = await resp.json();
        showToast(`Failed to load collection: ${errData.error}`, 'error');
      }
    } catch (err) {
      console.error(err);
      showToast('Error loading collection', 'error');
    } finally {
      setIsUploadingCollection(false);
    }
  };

  const processUploadedFile = async (file) => {
    setIsUploadingCollection(true);
    const formData = new FormData();
    formData.append('file', file);

    try {
      const resp = await fetch(`${API_BASE_URL}/collections/upload`, {
        method: 'POST',
        body: formData
      });

      if (resp.ok) {
        const data = await resp.json();
        const locations = data.locations || [];

        let addedCount = 0;
        let errors = [];
        for (const loc of locations) {
          const pathToAdd = loc.folder || loc.path;
          if (pathToAdd) {
            const addResp = await fetch(`${API_BASE_URL}/state/folders`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                path: pathToAdd,
                type: 'collection',
                description: loc.description || pathToAdd.split(/[\\/]/).pop() || pathToAdd
              })
            });
            if (addResp.ok) {
              const addData = await addResp.json();
              if (addData.added !== false) addedCount++;
            } else {
              const errData = await addResp.json();
              errors.push(errData.error);
            }
          }
        }
        
        if (addedCount > 0) {
          showToast(`Collection uploaded successfully! Added ${addedCount} folders.`);
        } else if (errors.length > 0) {
          showToast(`No folders added. Example error: ${errors[0]}`, 'error');
        } else {
          showToast('No valid folders found in collection.', 'warning');
        }
        
        fetchState();
      } else {
        const errData = await resp.json();
        showToast(`Failed to parse collection file: ${errData.error}`, 'error');
      }
    } catch (err) {
      console.error(err);
      showToast('Error parsing collection', 'error');
    } finally {
      setIsUploadingCollection(false);
    }
  };

  const handleUpdatePermissions = async (path, isPublic, allowedUsers) => {
    try {
      const resp = await fetch(`${API_BASE_URL}/state/folders/permissions`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path, isPublic, allowedUsers })
      });
      if (resp.ok) {
        showToast('Folder permissions updated successfully!');
        fetchState();
      } else {
        const data = await resp.json();
        showToast(`Failed to update permissions: ${data.error}`, 'error');
      }
    } catch (err) {
      console.error(err);
      showToast('Error updating permissions', 'error');
    }
  };

  return (
    <div className="dashboard-layout animated-fade">
      <Header onLogout={onLogout} version={appVersion} />

      <main className="container">
        <MetricsRow 
          foldersCount={folders.length} 
          stats={stats} 
          indexingStatus={indexingStatus} 
          onErrorClick={() => setShowErrorModal(true)}
        />

        <div className="dashboard-grid">
          <div className="dashboard-section">
            <AddLocationPanel 
              onAddManualFolder={handleAddManualFolder}
              onBrowseFolder={handleBrowseFolder}
              onBrowseCollectionFile={handleBrowseCollectionFile}
              onProcessUploadedFile={processUploadedFile}
              isUploadingCollection={isUploadingCollection}
            />
            <FoldersTable 
              folders={folders} 
              onRemoveFolder={handleRemoveFolder} 
              onUpdatePermissions={handleUpdatePermissions}
              selectedFolders={selectedFolders}
              onSelectionChange={setSelectedFolders}
            />
          </div>

          <div className="dashboard-section">
            <IndexingControls
              indexingStatus={indexingStatus}
              schedulerStatus={schedulerStatus}
              stats={stats}
              foldersCount={folders.length}
              onStart={handleStartIndexing}
              onPause={handlePause}
              onReset={handleReset}
              onFastSync={handleFastSync}
              onRepairMetadata={handleRepairMetadata}
              onStartScheduler={handleStartScheduler}
              onStopScheduler={handleStopScheduler}
            />
            <LogViewer logs={logs} onClearLogs={() => setLogs([])} />
          </div>
        </div>
      </main>

      {showErrorModal && (
        <ErrorModal 
          errors={recentErrors} 
          onClose={() => setShowErrorModal(false)} 
        />
      )}

      <Toast 
        message={toast.message} 
        type={toast.type} 
        onClose={() => setToast({ message: '', type: 'success' })} 
      />
    </div>
  );
}
