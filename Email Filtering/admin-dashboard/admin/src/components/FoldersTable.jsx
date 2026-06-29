import { useState } from 'react';
import { Layers, Search, Trash2, Edit2, X } from 'lucide-react';

export default function FoldersTable({ folders, onRemoveFolder, onUpdatePermissions, selectedFolders = [], onSelectionChange = () => {} }) {
  const [tableSearch, setTableSearch] = useState('');
  
  // Modal state
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingFolder, setEditingFolder] = useState(null);
  const [emailsInput, setEmailsInput] = useState('');

  const handleAccessChange = (folder, newAccessValue) => {
    if (newAccessValue === 'public') {
      onUpdatePermissions(folder.path, true, []);
    } else {
      // Set default input string
      setEmailsInput((folder.allowedUsers || []).join(', '));
      setEditingFolder(folder);
      setIsModalOpen(true);
    }
  };

  const handleEditClick = (folder) => {
    setEmailsInput((folder.allowedUsers || []).join(', '));
    setEditingFolder(folder);
    setIsModalOpen(true);
  };

  const handleSavePermissions = () => {
    if (!editingFolder) return;
    const emails = emailsInput.split(',').map(e => e.trim()).filter(e => e);
    onUpdatePermissions(editingFolder.path, false, emails);
    setIsModalOpen(false);
    setEditingFolder(null);
  };

  const filteredFolders = folders.filter(f => 
    f.path.toLowerCase().includes(tableSearch.toLowerCase()) || 
    f.description.toLowerCase().includes(tableSearch.toLowerCase())
  );

  const handleSelectAll = (e) => {
    if (e.target.checked) {
      onSelectionChange(filteredFolders.map(f => f.path));
    } else {
      onSelectionChange([]);
    }
  };

  const handleSelectOne = (path, isChecked) => {
    if (isChecked) {
      onSelectionChange([...selectedFolders, path]);
    } else {
      onSelectionChange(selectedFolders.filter(p => p !== path));
    }
  };

  const isAllSelected = filteredFolders.length > 0 && filteredFolders.every(f => selectedFolders.includes(f.path));

  return (
    <div className="locations-table-panel">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
        <h3 className="panel-title" style={{ margin: 0 }}>
          <Layers size={18} style={{ color: 'var(--primary-color)' }} /> Configured Folder Paths
        </h3>
        <div style={{ position: 'relative', width: '240px' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: 'var(--text-light)' }} />
          <input 
            type="text" 
            className="input-control" 
            style={{ paddingLeft: 30, paddingTop: 6, paddingBottom: 6, fontSize: '12px' }}
            placeholder="Search folder paths..." 
            value={tableSearch}
            onChange={e => setTableSearch(e.target.value)}
          />
        </div>
      </div>

      <div className="table-responsive">
        <table className="locations-table">
          <thead>
            <tr>
              <th style={{ width: '40px', textAlign: 'center' }}>
                <input 
                  type="checkbox" 
                  checked={isAllSelected}
                  onChange={handleSelectAll}
                />
              </th>
              <th>Folder Path</th>
              <th>Origin</th>
              <th>Access</th>
              <th style={{ width: '80px', textAlign: 'center' }}>Action</th>
            </tr>
          </thead>
          <tbody>
            {filteredFolders.length === 0 ? (
              <tr>
                <td colSpan="3" style={{ textAlign: 'center', color: 'var(--text-light)', padding: '24px' }}>
                  No locations configured yet. Add folders manually or import a .mmcollection file.
                </td>
              </tr>
            ) : (
              filteredFolders.map((item, idx) => (
                <tr key={idx} className={selectedFolders.includes(item.path) ? 'selected-row' : ''}>
                  <td style={{ textAlign: 'center' }}>
                    <input 
                      type="checkbox" 
                      checked={selectedFolders.includes(item.path)}
                      onChange={(e) => handleSelectOne(item.path, e.target.checked)}
                    />
                  </td>
                  <td style={{ wordBreak: 'break-all', fontWeight: '500' }}>{item.path}</td>
                  <td>
                    <span className={`type-badge ${item.type}`}>
                      {item.type}
                    </span>
                  </td>
                  <td>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <select 
                        value={item.isPublic === false ? 'restricted' : 'public'}
                        onChange={(e) => handleAccessChange(item, e.target.value)}
                        style={{ padding: '4px 8px', borderRadius: '4px', border: '1px solid #ddd', fontSize: '13px' }}
                      >
                        <option value="public">Public</option>
                        <option value="restricted">Restricted</option>
                      </select>
                      {item.isPublic === false && (
                        <button 
                          className="action-btn secondary" 
                          style={{ padding: '4px 8px' }}
                          onClick={() => handleEditClick(item)}
                          title="Edit Allowed Users"
                        >
                          <Edit2 size={14} />
                        </button>
                      )}
                    </div>
                  </td>
                  <td>
                    <div style={{ display: 'flex', justifyContent: 'center' }}>
                      <button 
                        className="row-action-btn" 
                        onClick={() => onRemoveFolder(item.path)}
                        title="Remove path"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {isModalOpen && editingFolder && (
        <div className="modal-overlay">
          <div className="modal-content" style={{ maxWidth: '500px' }}>
            <div className="modal-header">
              <h3>Restricted Access</h3>
              <button className="close-btn" onClick={() => setIsModalOpen(false)}>
                <X size={20} />
              </button>
            </div>
            <div className="modal-body">
              <div style={{ marginBottom: '16px', padding: '12px', background: '#f1f5f9', borderRadius: '6px', wordBreak: 'break-all' }}>
                <span style={{ fontSize: '12px', color: '#64748b', fontWeight: '600', display: 'block', marginBottom: '4px' }}>TARGET LOCATION</span>
                <span style={{ fontSize: '14px', color: '#0f172a', fontWeight: '500' }}>{editingFolder.path}</span>
              </div>
              
              <p style={{ fontSize: '14px', marginBottom: '12px', color: '#334155' }}>
                Enter the email addresses of the users who are allowed to search inside this folder. Separate multiple emails with a comma.
              </p>
              <textarea 
                className="input-control" 
                rows="4"
                placeholder="e.g. paul@company.com, suhail@company.com"
                value={emailsInput}
                onChange={(e) => setEmailsInput(e.target.value)}
                style={{ width: '100%', padding: '12px', fontSize: '14px', borderRadius: '6px', border: '1px solid #cbd5e1' }}
              />
            </div>
            <div className="modal-footer" style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px' }}>
              <button className="action-btn secondary" onClick={() => setIsModalOpen(false)}>Cancel</button>
              <button className="action-btn primary" onClick={handleSavePermissions}>Save</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
