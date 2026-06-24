import { X, AlertCircle } from 'lucide-react';

export default function ErrorModal({ errors, onClose }) {
  if (!errors || errors.length === 0) return null;

  return (
    <div style={overlayStyle}>
      <div style={modalStyle}>
        <div style={headerStyle}>
          <div style={titleStyle}>
            <AlertCircle size={20} color="#e11d48" />
            <span style={{ fontWeight: 600 }}>Recent Error Logs</span>
          </div>
          <button style={closeBtnStyle} onClick={onClose}><X size={18} /></button>
        </div>
        
        <div style={contentStyle}>
          {errors.map((err, i) => (
            <div key={i} style={errorItemStyle}>
              <div style={{ fontSize: '11px', color: '#666', marginBottom: '4px' }}>
                {err.timestamp} • {err.filePath}
              </div>
              <div style={{ fontSize: '13px', color: '#e11d48', fontWeight: 500 }}>
                {err.error}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

const overlayStyle = {
  position: 'fixed',
  top: 0, left: 0, right: 0, bottom: 0,
  backgroundColor: 'rgba(0,0,0,0.5)',
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'center',
  zIndex: 1000
};

const modalStyle = {
  backgroundColor: '#fff',
  borderRadius: '8px',
  width: '90%',
  maxWidth: '600px',
  maxHeight: '80vh',
  display: 'flex',
  flexDirection: 'column',
  boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1)',
  overflow: 'hidden'
};

const headerStyle = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  padding: '16px 24px',
  borderBottom: '1px solid #eee',
  backgroundColor: '#f8fafc'
};

const titleStyle = {
  display: 'flex',
  alignItems: 'center',
  gap: '8px',
  color: '#0f172a',
  fontSize: '16px'
};

const closeBtnStyle = {
  background: 'transparent',
  border: 'none',
  cursor: 'pointer',
  color: '#64748b',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center'
};

const contentStyle = {
  padding: '24px',
  overflowY: 'auto',
  display: 'flex',
  flexDirection: 'column',
  gap: '12px'
};

const errorItemStyle = {
  padding: '12px',
  backgroundColor: '#fff1f2',
  border: '1px solid #ffe4e6',
  borderRadius: '6px',
  wordBreak: 'break-all'
};
