import { useState } from 'react';
import { Layers, Search, Trash2 } from 'lucide-react';

export default function FoldersTable({ folders, onRemoveFolder }) {
  const [tableSearch, setTableSearch] = useState('');

  const filteredFolders = folders.filter(f => 
    f.path.toLowerCase().includes(tableSearch.toLowerCase()) || 
    f.description.toLowerCase().includes(tableSearch.toLowerCase())
  );

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
              <th>Folder Path</th>
              <th>Origin</th>
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
                <tr key={idx}>
                  <td style={{ wordBreak: 'break-all', fontWeight: '500' }}>{item.path}</td>
                  <td>
                    <span className={`type-badge ${item.type}`}>
                      {item.type}
                    </span>
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
    </div>
  );
}
