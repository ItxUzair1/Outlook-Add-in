import { useState } from 'react';
import { X } from 'lucide-react';

export default function RejectRequestModal({ request, onClose, onSubmit }) {
  const [message, setMessage] = useState('');

  const handleRejectNoEmail = () => {
    onSubmit(message, false);
  };

  const handleRejectWithEmail = () => {
    if (message.trim()) {
      onSubmit(message, true);
    }
  };

  return (
    <div className="modal-overlay">
      <div className="modal-content" style={{ width: 500, maxWidth: '90%' }}>
        <div className="modal-header">
          <h3 className="modal-title">Reject Indexing Request</h3>
          <button className="icon-btn" onClick={onClose}>
            <X size={18} />
          </button>
        </div>
        
        <div className="modal-body">
          <p style={{ margin: '0 0 16px 0', fontSize: 14, color: '#605e5c', lineHeight: 1.5 }}>
            You are rejecting the request for <strong>{request.projectName}</strong>. 
            Please provide a reason for the rejection below, which can be sent to <strong>{request.userEmail}</strong>.
          </p>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div className="input-group">
              <label style={{ display: 'block', marginBottom: '6px', fontWeight: 500 }}>Rejection Reason</label>
              <textarea 
                value={message}
                onChange={(e) => setMessage(e.target.value)}
                placeholder="e.g. This project number does not exist on the specified drive..."
                autoFocus
                rows={4}
                style={{ 
                  width: '100%', 
                  padding: '8px 12px', 
                  borderRadius: 6, 
                  border: '1px solid #e2e8f0',
                  resize: 'vertical',
                  fontSize: 14,
                  fontFamily: 'inherit'
                }}
              />
            </div>
            
            <div style={{ display: 'flex', gap: 12, justifyContent: 'flex-end', marginTop: 8 }}>
              <button 
                type="button" 
                className="btn btn-secondary" 
                onClick={onClose}
              >
                Cancel
              </button>
              <button 
                type="button" 
                className="btn btn-secondary" 
                onClick={handleRejectNoEmail}
                style={{ color: '#ef4444', borderColor: '#ef4444' }}
              >
                Reject (No Email)
              </button>
              <button 
                type="button" 
                className="btn btn-danger" 
                onClick={handleRejectWithEmail}
                disabled={!message.trim()}
              >
                Reject & Send Email
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
