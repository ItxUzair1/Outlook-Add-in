import { useState, useMemo } from 'react';
import { X, AlertCircle, Search, Copy, CheckCheck, FileWarning, ChevronDown, ChevronUp } from 'lucide-react';

// Plain helper — NOT a component, so React won't complain about it being
// created during render. Receives sortField/sortDir as plain arguments.
function renderSortIcon(activeSortField, sortDir, field) {
  if (activeSortField !== field) return <ChevronDown size={13} style={{ opacity: 0.3 }} />;
  return sortDir === 'asc'
    ? <ChevronUp size={13} style={{ color: '#0078d4' }} />
    : <ChevronDown size={13} style={{ color: '#0078d4' }} />;
}

export default function ErrorModal({ errors, onClose }) {
  const [search, setSearch] = useState('');
  const [sortField, setSortField] = useState('timestamp');
  const [sortDir, setSortDir] = useState('desc');
  const [copied, setCopied] = useState(false);
  const [expandedRow, setExpandedRow] = useState(null);

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return (errors || []).filter(e =>
      !q ||
      (e.filePath || '').toLowerCase().includes(q) ||
      (e.error || '').toLowerCase().includes(q)
    );
  }, [errors, search]);

  const sorted = useMemo(() => {
    return [...filtered].sort((a, b) => {
      const va = a[sortField] || '';
      const vb = b[sortField] || '';
      const cmp = va < vb ? -1 : va > vb ? 1 : 0;
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [filtered, sortField, sortDir]);

  function toggleSort(field) {
    if (sortField === field) {
      setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDir('asc');
    }
  }



  function getFileName(filePath) {
    if (!filePath) return '—';
    return filePath.replace(/\\/g, '/').split('/').pop() || filePath;
  }

  function copyAll() {
    const text = (errors || []).map(e =>
      `[${e.timestamp}] ${e.filePath}\n  ERROR: ${e.error}`
    ).join('\n\n');
    navigator.clipboard.writeText(text).catch(() => {});
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  const total = (errors || []).length;
  const showing = sorted.length;

  return (
    <div style={overlayStyle} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={modalStyle}>

        {/* ── Header ─────────────────────────────────────────────── */}
        <div style={headerStyle}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={iconWrapStyle}>
              <AlertCircle size={18} color="#e11d48" />
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16, color: '#0f172a' }}>
                Parse Error Report
              </div>
              <div style={{ fontSize: 12, color: '#64748b', marginTop: 1 }}>
                {total} email{total !== 1 ? 's' : ''} failed to parse during indexing
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <button
              style={copyBtnStyle}
              onClick={copyAll}
              title="Copy all errors to clipboard"
            >
              {copied ? <CheckCheck size={15} color="#16a34a" /> : <Copy size={15} />}
              <span>{copied ? 'Copied!' : 'Copy All'}</span>
            </button>
            <button style={closeBtnStyle} onClick={onClose} title="Close">
              <X size={18} />
            </button>
          </div>
        </div>

        {/* ── Search bar ─────────────────────────────────────────── */}
        <div style={searchBarStyle}>
          <Search size={14} color="#94a3b8" style={{ flexShrink: 0 }} />
          <input
            type="text"
            placeholder="Filter by filename or error message…"
            value={search}
            onChange={e => setSearch(e.target.value)}
            style={searchInputStyle}
          />
          {search && (
            <span style={{ fontSize: 12, color: '#94a3b8', whiteSpace: 'nowrap' }}>
              {showing} of {total}
            </span>
          )}
        </div>

        {/* ── Table ──────────────────────────────────────────────── */}
        <div style={tableWrapStyle}>
          {sorted.length === 0 ? (
            <div style={emptyStyle}>
              <FileWarning size={40} color="#cbd5e1" />
              <div style={{ color: '#94a3b8', fontSize: 14, marginTop: 8 }}>
                {search ? 'No errors match your filter.' : 'No errors recorded.'}
              </div>
            </div>
          ) : (
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={{ ...thStyle, width: 42 }}>#</th>
                  <th
                    style={{ ...thStyle, cursor: 'pointer', userSelect: 'none' }}
                    onClick={() => toggleSort('timestamp')}
                  >
                    <span style={thInnerStyle}>
                      Time {renderSortIcon(sortField, sortDir, 'timestamp')}
                    </span>
                  </th>
                  <th
                    style={{ ...thStyle, cursor: 'pointer', userSelect: 'none' }}
                    onClick={() => toggleSort('filePath')}
                  >
                    <span style={thInnerStyle}>
                      Email File {renderSortIcon(sortField, sortDir, 'filePath')}
                    </span>
                  </th>
                  <th
                    style={{ ...thStyle, cursor: 'pointer', userSelect: 'none' }}
                    onClick={() => toggleSort('error')}
                  >
                    <span style={thInnerStyle}>
                      Error Message {renderSortIcon(sortField, sortDir, 'error')}
                    </span>
                  </th>
                </tr>
              </thead>
              <tbody>
                {sorted.map((err, i) => {
                  const isExpanded = expandedRow === i;
                  const fileName = getFileName(err.filePath);
                  return (
                    <>
                      <tr
                        key={`row-${i}`}
                        style={{
                          ...trStyle,
                          background: i % 2 === 0 ? '#ffffff' : '#fafafa',
                          cursor: 'pointer'
                        }}
                        onClick={() => setExpandedRow(isExpanded ? null : i)}
                      >
                        <td style={{ ...tdStyle, color: '#94a3b8', fontSize: 11, textAlign: 'center' }}>
                          {i + 1}
                        </td>
                        <td style={{ ...tdStyle, whiteSpace: 'nowrap', color: '#64748b', fontSize: 11 }}>
                          {err.timestamp || '—'}
                        </td>
                        <td style={tdStyle}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                            <span style={fileTagStyle} title={err.filePath}>
                              {fileName}
                            </span>
                          </div>
                        </td>
                        <td style={{ ...tdStyle, color: '#b91c1c', fontSize: 12 }}>
                          <div style={{
                            maxWidth: 360,
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            whiteSpace: isExpanded ? 'normal' : 'nowrap'
                          }}>
                            {err.error || '—'}
                          </div>
                        </td>
                      </tr>
                      {isExpanded && (
                        <tr key={`expand-${i}`}>
                          <td colSpan={4} style={expandCellStyle}>
                            <div style={expandInnerStyle}>
                              <div style={{ marginBottom: 6 }}>
                                <span style={labelStyle}>Full Path:</span>
                                <span style={monoStyle}>{err.filePath || '—'}</span>
                              </div>
                              <div>
                                <span style={labelStyle}>Error:</span>
                                <span style={{ ...monoStyle, color: '#dc2626' }}>{err.error || '—'}</span>
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>

        {/* ── Footer ─────────────────────────────────────────────── */}
        <div style={footerStyle}>
          <span style={{ fontSize: 12, color: '#94a3b8' }}>
            Click any row to expand full path &amp; error • Showing last 50 errors per session
          </span>
          <button style={closeFooterBtnStyle} onClick={onClose}>
            Close
          </button>
        </div>
      </div>
    </div>
  );
}

/* ── Styles ──────────────────────────────────────────────────────────────── */

const overlayStyle = {
  position: 'fixed',
  inset: 0,
  backgroundColor: 'rgba(15,23,42,0.6)',
  backdropFilter: 'blur(4px)',
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'center',
  zIndex: 1000,
  padding: 16
};

const modalStyle = {
  backgroundColor: '#ffffff',
  borderRadius: 12,
  width: '100%',
  maxWidth: 900,
  maxHeight: '88vh',
  display: 'flex',
  flexDirection: 'column',
  boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)',
  overflow: 'hidden',
  animation: 'fadeInUp 0.2s ease'
};

const headerStyle = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  padding: '18px 24px',
  borderBottom: '1px solid #f1f5f9',
  background: 'linear-gradient(to right, #fff1f2, #ffffff)',
  flexShrink: 0
};

const iconWrapStyle = {
  width: 36,
  height: 36,
  borderRadius: 8,
  background: '#fee2e2',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center'
};

const copyBtnStyle = {
  display: 'flex',
  alignItems: 'center',
  gap: 5,
  padding: '6px 12px',
  fontSize: 12,
  fontWeight: 600,
  background: '#f8fafc',
  border: '1px solid #e2e8f0',
  borderRadius: 6,
  cursor: 'pointer',
  color: '#475569',
  transition: 'all 0.15s'
};

const closeBtnStyle = {
  width: 32,
  height: 32,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  background: 'transparent',
  border: '1px solid #e2e8f0',
  borderRadius: 6,
  cursor: 'pointer',
  color: '#64748b'
};

const searchBarStyle = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '10px 24px',
  borderBottom: '1px solid #f1f5f9',
  background: '#f8fafc',
  flexShrink: 0
};

const searchInputStyle = {
  flex: 1,
  border: 'none',
  background: 'transparent',
  fontSize: 13,
  color: '#0f172a',
  outline: 'none'
};

const tableWrapStyle = {
  flex: 1,
  overflowY: 'auto',
  overflowX: 'auto'
};

const tableStyle = {
  width: '100%',
  borderCollapse: 'collapse',
  fontSize: 13
};

const thStyle = {
  background: '#f8fafc',
  padding: '10px 14px',
  borderBottom: '1px solid #e2e8f0',
  fontWeight: 600,
  color: '#374151',
  textAlign: 'left',
  fontSize: 12,
  position: 'sticky',
  top: 0,
  zIndex: 10
};

const thInnerStyle = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 4
};

const trStyle = {
  borderBottom: '1px solid #f1f5f9',
  transition: 'background 0.1s'
};

const tdStyle = {
  padding: '10px 14px',
  verticalAlign: 'middle'
};

const fileTagStyle = {
  display: 'inline-block',
  background: '#eff6ff',
  color: '#1d4ed8',
  border: '1px solid #bfdbfe',
  borderRadius: 4,
  padding: '2px 7px',
  fontSize: 12,
  fontFamily: 'Consolas, monospace',
  maxWidth: 240,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

const expandCellStyle = {
  padding: 0,
  borderBottom: '2px solid #fee2e2'
};

const expandInnerStyle = {
  padding: '12px 24px',
  background: '#fff5f5',
  fontSize: 12
};

const labelStyle = {
  fontWeight: 700,
  color: '#64748b',
  marginRight: 8,
  fontSize: 11,
  textTransform: 'uppercase',
  letterSpacing: '0.05em'
};

const monoStyle = {
  fontFamily: 'Consolas, monospace',
  color: '#0f172a',
  wordBreak: 'break-all'
};

const emptyStyle = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  padding: '60px 24px'
};

const footerStyle = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  padding: '12px 24px',
  borderTop: '1px solid #f1f5f9',
  background: '#f8fafc',
  flexShrink: 0
};

const closeFooterBtnStyle = {
  padding: '7px 20px',
  fontSize: 13,
  fontWeight: 600,
  background: '#0f172a',
  color: '#ffffff',
  border: 'none',
  borderRadius: 6,
  cursor: 'pointer'
};
