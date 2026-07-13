import { useState } from 'react';
import { X } from 'lucide-react';

export default function ApproveRequestModal({ request, onClose, onSubmit }) {
  const [networkPath, setNetworkPath] = useState('');

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!networkPath.trim()) return;
    onSubmit(networkPath);
  };

  return (
    <div className="modal-overlay">
      <div className="modal-content" style={{ width: 480, maxWidth: '90%' }}>
        <div className="modal-header">
          <h3 className="modal-title">Approve Indexing Request</h3>
          <button className="icon-btn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="modal-body">
          <p style={{ margin: '0 0 16px 0', fontSize: 14, color: '#605e5c', lineHeight: 1.5 }}>
            You are approving the request for <strong>{request.projectName}</strong>. 
            Please provide the exact network path where this project is located. 
            Once approved, it will be added to the index and an email will be sent to <strong>{request.userEmail}</strong>.
          </p>
          
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div className="input-group">
              <label>Project Network Path</label>
              <input 
                type="text" 
                value={networkPath}
                onChange={e => setNetworkPath(e.target.value)}
                placeholder="e.g. Z:\Projects\12345"
                autoFocus
                className="path-input"
              />
            </div>
            
            <div style={{ display: 'flex', gap: 12, justifyContent: 'flex-end', marginTop: 8 }}>
              <button type="button" className="btn btn-secondary" onClick={onClose}>Cancel</button>
              <button type="submit" className="btn btn-primary" disabled={!networkPath.trim()}>
                Approve & Index
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}
