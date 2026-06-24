import { useEffect } from 'react';

export default function Toast({ message, type = 'success', onClose }) {
  useEffect(() => {
    if (!message) return;
    const timer = setTimeout(() => {
      onClose();
    }, 3000);
    return () => clearTimeout(timer);
  }, [message, onClose]);

  if (!message) return null;

  const getColors = () => {
    if (type === 'error') return { bg: '#fde7e9', text: '#a80000', border: '#a80000' };
    if (type === 'warning') return { bg: '#fff4ce', text: '#d83b01', border: '#d83b01' };
    return { bg: '#dff6dd', text: '#107c41', border: '#107c41' };
  };

  const colors = getColors();

  return (
    <div style={{
      position: 'fixed',
      top: '80px', // Just below the header (header is ~72px)
      right: '24px',
      backgroundColor: colors.bg,
      color: colors.text,
      borderLeft: `4px solid ${colors.border}`,
      padding: '16px 24px',
      borderRadius: '4px',
      boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
      zIndex: 9999,
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
      fontWeight: '500',
      fontSize: '14px',
      animation: 'fadeIn 0.3s ease-out forwards'
    }}>
      {message}
      <button 
        onClick={onClose} 
        style={{
          background: 'transparent',
          border: 'none',
          color: colors.text,
          cursor: 'pointer',
          marginLeft: '12px',
          fontWeight: 'bold',
          opacity: 0.7
        }}
      >
        ✕
      </button>
    </div>
  );
}
