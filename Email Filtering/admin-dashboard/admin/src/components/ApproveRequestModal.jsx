import { X } from 'lucide-react';

export default function ApproveRequestModal({ request, onClose, onSubmit }) {
  const handleSubmit = (e) => {
    e.preventDefault();
    onSubmit();
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
            Are you sure you indexed the project <strong>{request.projectName}</strong>? 
            Once approved, an email will be sent to <strong>{request.userEmail}</strong> confirming that the project is now searchable.
          </p>
          
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div style={{ display: 'flex', gap: 12, justifyContent: 'flex-end', marginTop: 8 }}>
              <button type="button" className="btn btn-secondary" onClick={onClose}>Cancel</button>
              <button type="submit" className="btn btn-primary">
                Approve
              </button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}
