import { useState } from 'react';
import { 
  Database, 
  FolderOpen, 
  Plus, 
  Layers, 
  RefreshCw, 
  Upload 
} from 'lucide-react';

export default function AddLocationPanel({
  onAddManualFolder,
  onBrowseFolder,
  onBrowseCollectionFile,
  onProcessUploadedFile,
  isUploadingCollection
}) {
  const [manualPath, setManualPath] = useState('');
  const [collectionPath, setCollectionPath] = useState('');
  const [dragActive, setDragActive] = useState(false);

  const handleAddSubmit = (e) => {
    e.preventDefault();
    if (!manualPath.trim()) return;
    onAddManualFolder(manualPath.trim());
    setManualPath('');
  };

  const handleBrowseFolderClick = async () => {
    const path = await onBrowseFolder();
    if (path) setManualPath(path);
  };

  const handleBrowseCollectionClick = async () => {
    const path = await onBrowseCollectionFile();
    if (path) setCollectionPath(path);
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.mmcollection')) {
        onProcessUploadedFile(file);
      } else {
        alert('Invalid file format. Please upload a .mmcollection file.');
      }
    }
  };

  const handleFileInput = (e) => {
    if (e.target.files && e.target.files[0]) {
      onProcessUploadedFile(e.target.files[0]);
    }
  };

  return (
    <div className="add-locations-panel">
      <h3 className="panel-title">
        <Database size={18} style={{ color: 'var(--primary-color)' }} /> Load Email Locations
      </h3>
      
      {/* Option 1: Browse or Paste folder path */}
      <form onSubmit={handleAddSubmit} className="input-group-row">
        <div className="input-with-btn">
          <input 
            type="text" 
            className="input-control" 
            placeholder="Paste absolute path to local or network folder (e.g. C:\Emails)" 
            value={manualPath}
            onChange={e => setManualPath(e.target.value)}
          />
          <button type="button" className="action-btn secondary" onClick={handleBrowseFolderClick}>
            <FolderOpen size={16} /> Browse Folder
          </button>
        </div>
        <button type="submit" className="action-btn primary" disabled={!manualPath.trim()}>
          <Plus size={16} /> Add Folder
        </button>
      </form>

      {/* Option 2: Upload Collection */}
      <div className="collection-upload-container">
        {/* Browse via Native koyofile */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
          <label style={{ fontSize: '13px', fontWeight: '600' }}>Select .mmcollection File (Local Server)</label>
          <div style={{ display: 'flex', gap: '8px' }}>
            <input 
              type="text" 
              className="input-control" 
              placeholder="Browse or paste collection path" 
              value={collectionPath}
              onChange={e => setCollectionPath(e.target.value)}
              readOnly
            />
            <button 
              type="button" 
              className="action-btn secondary" 
              onClick={handleBrowseCollectionClick} 
              disabled={isUploadingCollection}
            >
              {isUploadingCollection ? <RefreshCw size={16} className="pulsing" /> : <Layers size={16} />} Browse Collection
            </button>
          </div>
        </div>

        {/* Direct Upload Dropzone */}
        <div 
          className={`dropzone ${dragActive ? 'active' : ''}`}
          onDragEnter={handleDrag}
          onDragOver={handleDrag}
          onDragLeave={handleDrag}
          onDrop={handleDrop}
          onClick={() => document.getElementById('collection-file-input').click()}
        >
          <Upload className="dropzone-icon" size={24} />
          <span style={{ fontSize: '13px', fontWeight: '600' }}>Drag & Drop .mmcollection file</span>
          <span style={{ fontSize: '11px', color: 'var(--text-light)' }}>or click to browse from browser</span>
          <input 
            type="file" 
            id="collection-file-input" 
            style={{ display: 'none' }} 
            accept=".mmcollection" 
            onChange={handleFileInput}
          />
        </div>
      </div>
    </div>
  );
}
