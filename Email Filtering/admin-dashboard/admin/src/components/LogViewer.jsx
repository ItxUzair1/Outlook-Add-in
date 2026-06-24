import { useEffect, useRef, useState } from 'react';
import { Terminal, Lock, Unlock } from 'lucide-react';

export default function LogViewer({ logs, onClearLogs }) {
  const containerRef = useRef(null);
  const [autoScroll, setAutoScroll] = useState(true);

  // Handle manual scroll to determine if we should auto-scroll
  const handleScroll = () => {
    if (!containerRef.current) return;
    const { scrollTop, scrollHeight, clientHeight } = containerRef.current;
    // Check if user is near bottom (within 100px)
    const isNearBottom = scrollHeight - scrollTop - clientHeight < 100;
    
    if (autoScroll !== isNearBottom) {
      setAutoScroll(isNearBottom);
    }
  };

  // Autoscroll logs
  useEffect(() => {
    if (autoScroll && containerRef.current) {
      // Use instant scroll to prevent smooth animation conflicts during fast polling
      containerRef.current.scrollTop = containerRef.current.scrollHeight;
    }
  }, [logs, autoScroll]);

  return (
    <div className="log-viewer-panel">
      <div className="log-viewer-header">
        <div className="log-title">
          <Terminal size={16} /> Console Output logs
          <span style={{ 
            marginLeft: '12px', 
            fontSize: '11px', 
            color: autoScroll ? '#dff6dd' : '#fde7e9', 
            opacity: 0.8,
            display: 'flex', 
            alignItems: 'center', 
            gap: '4px',
            fontWeight: 'normal'
          }}>
            {autoScroll ? <Lock size={12} /> : <Unlock size={12} />}
            {autoScroll ? 'Auto-scroll ON' : 'Auto-scroll PAUSED'}
          </span>
        </div>
        <button 
          className="log-clear-btn" 
          onClick={onClearLogs}
        >
          Clear Console
        </button>
      </div>
      <div 
        className="log-content" 
        ref={containerRef} 
        onScroll={handleScroll}
        style={{ overflowAnchor: 'none' }}
      >
        {logs.length === 0 ? (
          <div style={{ color: '#858585', fontStyle: 'italic' }}>Console output idle... Start indexing to stream live log feeds.</div>
        ) : (
          logs.map((logLine, idx) => (
            <div className="log-line" key={idx}>
              {logLine}
            </div>
          ))
        )}
      </div>
    </div>
  );
}
