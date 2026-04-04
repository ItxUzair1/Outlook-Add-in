import * as React from "react";
import {
  Search20Regular,
  FolderOpen20Regular,
  Dismiss20Regular,
  Filter20Regular,
  ArrowClockwise20Regular,
  ArrowCounterclockwise20Regular,
  Mail20Regular,
  Attach20Regular,
  CalendarMonth20Regular,
  Person20Regular,
  MailTemplate20Regular,
  TextBulletList20Regular,
  Settings20Regular,
  QuestionCircle20Regular,
  MoreHorizontal20Regular,
  ChevronLeft20Regular,
  Desktop20Regular,
  Checkmark20Regular,
  MailSettings20Regular,
  ChevronDown20Regular,
  ArrowSync20Regular,
} from "@fluentui/react-icons";

const BASE_URL = "http://localhost:4000";

const DATE_RANGES = [
  { label: "Past Month", value: "1m" },
  { label: "Past 3 Months", value: "3m" },
  { label: "Past 6 Months", value: "6m" },
  { label: "Past Year", value: "1y" },
  { label: "All Time", value: "all" },
];

function relativeDate(dateStr) {
  if (!dateStr) return "";
  const date = new Date(dateStr);
  const now = new Date();
  const diff = now - date;
  const days = Math.floor(diff / (1000 * 60 * 60 * 24));
  if (days === 0) return "Today";
  if (days === 1) return "Yesterday";
  if (days < 7) return `${days} days ago`;
  if (days < 14) return "One week ago";
  if (days < 21) return "Two weeks ago";
  if (days < 31) return "Three weeks ago";
  if (days < 60) return "Last month";
  if (days < 180) return "A few months ago";
  return date.toLocaleDateString("en-GB", { year: "numeric", month: "short", day: "numeric" });
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return d.toLocaleString("en-GB", {
    day: "2-digit", month: "short", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: false
  }).replace(",", "");
}

function groupByRelativeDate(results) {
  const groups = {};
  results.forEach(r => {
    const key = relativeDate(r.sentAt || r.filedAt);
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });
  return groups;
}

export default function SearchDialog({ onClose }) {
  const [dateRange, setDateRange] = React.useState("6m");
  const [from, setFrom] = React.useState("");
  const [to, setTo] = React.useState("");
  const [cc, setCc] = React.useState("");
  const [subject, setSubject] = React.useState("");
  const [location, setLocation] = React.useState("");
  const [keywords, setKeywords] = React.useState("");
  const [hasAttachments, setHasAttachments] = React.useState(false);
  const [isIncludingEnabled, setIsIncludingEnabled] = React.useState(false);
  const [selectedType, setSelectedType] = React.useState("emails");
  const [selectedRowIds, setSelectedRowIds] = React.useState(new Set());
  const [isHelpOpen, setIsHelpOpen] = React.useState(false);
  const [isSyncing, setIsSyncing] = React.useState(false);
  const [syncMessage, setSyncMessage] = React.useState("");
  const [activeMenuId, setActiveMenuId] = React.useState(null);
  const [itemToDelete, setItemToDelete] = React.useState(null);
  const [results, setResults] = React.useState(null);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState("");

  const getProtocol = (path = "") => {
    if (path.startsWith("\\\\")) return { label: "Network", icon: <Desktop20Regular style={{ color: "#0078d4" }} /> };
    if (/^[a-zA-Z]:/.test(path)) return { label: "Local", icon: <Desktop20Regular style={{ color: "#ffb900" }} /> };
    return { label: "Cloud", icon: <FolderOpen20Regular style={{ color: "#0078d4" }} /> };
  };

  const handleSyncIndex = async () => {
      setIsSyncing(true);
      setSyncMessage("");
      try {
          const resp = await fetch(`${BASE_URL}/api/search/sync`, { method: "POST" });
          if (resp.ok) {
              const data = await resp.json();
              setSyncMessage(`Index Synced! Removed ${data.removedCount} stale entries.`);
              runSearch(); // Refresh the list
              setTimeout(() => setSyncMessage(""), 5000); // Clear message after 5s
          } else {
              alert("Sync failed. Server might be unreachable.");
          }
      } catch (err) {
          alert(`Sync failed: ${err.message}`);
      } finally {
          setIsSyncing(false);
      }
  };

  const handleOpenItem = async (r) => {
      try {
          const resp = await fetch(`${BASE_URL}/api/search/open`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ filePath: r.filePath })
          });
          if (resp.status === 404) {
              const encodedId = encodeURIComponent(r.id);
              if (window.confirm("The file was not found at its original location. It may have been moved or deleted.\n\nWould you like to remove this search entry from the history?")) {
                  const delResp = await fetch(`${BASE_URL}/api/search/${encodedId}`, { method: "DELETE" });
                  if (delResp.ok) runSearch();
              }
              setActiveMenuId(null);
              return;
          }
          if (!resp.ok) {
              const data = await resp.json();
              alert(`Error: ${data.error || "Could not open file"}`);
          }
      } catch (err) {
          alert(`Open failed: ${err.message}`);
      }
      setActiveMenuId(null);
  };

  const handleDeleteItem = (r) => {
      setItemToDelete(r);
      setActiveMenuId(null);
  };

  const handleConfirmDelete = async () => {
      if (!itemToDelete) return;
      try {
          const encodedId = encodeURIComponent(itemToDelete.id);
          const resp = await fetch(`${BASE_URL}/api/search/${encodedId}`, { method: "DELETE" });
          if (resp.ok) {
              runSearch();
          } else {
              const data = await resp.json();
              alert(`Delete failed: ${data.error}`);
          }
      } catch (err) {
          alert(`Delete failed: ${err.message}`);
      } finally {
          setItemToDelete(null);
      }
  };

  async function runSearch() {
    setLoading(true);
    setError("");
    try {
      const params = new URLSearchParams();
      if (dateRange) params.set("dateRange", dateRange);
      if (from.trim()) params.set("from", from.trim());
      if (to.trim()) params.set("to", to.trim());
      if (cc.trim()) params.set("cc", cc.trim());
      if (subject.trim()) params.set("subject", subject.trim());
      if (location.trim()) params.set("location", location.trim());
      if (keywords.trim()) params.set("keywords", keywords.trim());
      if (hasAttachments) params.set("hasAttachments", "true");
      if (isIncludingEnabled) params.set("including", "true");

      const resp = await fetch(`${BASE_URL}/api/search?${params.toString()}`);
      if (!resp.ok) throw new Error(`Search failed: ${resp.statusText}`);
      const data = await resp.json();
      setResults(data);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  function clearFilters() {
    setDateRange("6m");
    setFrom("");
    setTo("");
    setCc("");
    setSubject("");
    setLocation("");
    setKeywords("");
    setHasAttachments(false);
    setIsIncludingEnabled(false);
    setSelectedRowIds(new Set());
    setResults(null);
    setError("");
  }

  const handleSelectAll = (e) => {
    if (e.target.checked && results?.results) {
      setSelectedRowIds(new Set(results.results.map(r => r.id)));
    } else {
      setSelectedRowIds(new Set());
    }
  };

  const handleSelectRow = (id) => {
    const next = new Set(selectedRowIds);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    setSelectedRowIds(next);
  };

  const grouped = results ? groupByRelativeDate(results.results || []) : {};

  return (
    <div style={{
      display: "flex", flexDirection: "column", height: "100vh",
      fontFamily: "Segoe UI, sans-serif", backgroundColor: "#f8f8f8",
      overflow: "hidden",
    }}>

      {/* ── Top Search Bar ──────────────────────────────────────────────── */}
      <div style={{
        display: "flex", alignItems: "center", gap: 8,
        padding: "8px 12px", backgroundColor: "#ffffff",
        borderBottom: "1px solid #edebe9", flexShrink: 0,
      }}>
        {/* Location search */}
        <div style={{
          display: "flex", alignItems: "center", gap: 6, flex: 1,
          backgroundColor: "#f3f2f1", borderRadius: 4, padding: "6px 10px",
          border: "1px solid transparent",
        }}>
          <FolderOpen20Regular style={{ color: "#0078d4" }} />
          <input
            placeholder="Search By Filed Location"
            value={location}
            onChange={e => setLocation(e.target.value)}
            onKeyDown={e => e.key === "Enter" && runSearch()}
            style={{ border: "none", background: "transparent", outline: "none", flex: 1, fontSize: 13, fontFamily: "Segoe UI" }}
          />
        </div>

        {/* Keyword search */}
        <div style={{
          display: "flex", alignItems: "center", gap: 6, flex: 1.5,
          backgroundColor: "#f3f2f1", borderRadius: 4, padding: "6px 10px",
          border: "1px solid transparent",
        }}>
          <Search20Regular style={{ color: "#0078d4" }} />
          <input
            placeholder="Search For Any Keywords"
            value={keywords}
            onChange={e => setKeywords(e.target.value)}
            onKeyDown={e => e.key === "Enter" && runSearch()}
            style={{ border: "none", background: "transparent", outline: "none", flex: 1, fontSize: 13, fontFamily: "Segoe UI" }}
          />
          <ArrowCounterclockwise20Regular 
              style={{ color: "#605e5c", cursor: "pointer" }} 
              onClick={() => { setKeywords(""); setLocation(""); }}
          />
        </div>

        {/* Actions */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, color: "#605e5c" }}>
          {syncMessage && (
              <span style={{ fontSize: 12, color: "#107c10", fontWeight: 600 }}>{syncMessage}</span>
          )}
          <ArrowSync20Regular 
              style={{ 
                  cursor: "pointer", 
                  animation: isSyncing ? "spin 1s linear infinite" : "none"
              }} 
              onClick={handleSyncIndex}
              title="Sync Index (Cleanup missing files)"
          />
          <style>{`
              @keyframes spin {
                  from { transform: rotate(0deg); }
                  to { transform: rotate(360deg); }
              }
          `}</style>
          <Settings20Regular style={{ cursor: "pointer" }} />
          <QuestionCircle20Regular 
              style={{ cursor: "pointer" }} 
              onClick={() => setIsHelpOpen(true)}
              title="Help and Search Guide"
          />
          <button onClick={runSearch}
            style={{ 
              background: "#0078d4", border: "none", borderRadius: 4, 
              padding: "6px 16px", color: "#fff", cursor: "pointer", 
              fontSize: 13, fontWeight: 600, fontFamily: "Segoe UI",
              display: "flex", alignItems: "center", gap: 4 
            }}>
            Search
          </button>
        </div>
      </div>

      {/* ── Help Modal ─────────────────────────────────────────────────── */}
      {isHelpOpen && (
          <div style={{
              position: "fixed", inset: 0, zIndex: 10000,
              display: "flex", alignItems: "center", justifyContent: "center",
              backgroundColor: "rgba(0,0,0,0.4)"
          }}>
              <div style={{
                  width: 500, backgroundColor: "#fff", borderRadius: 8,
                  boxShadow: "0 8px 32px rgba(0,0,0,0.2)", padding: 24,
                  position: "relative"
              }}>
                  <Dismiss20Regular 
                      style={{ position: "absolute", top: 16, right: 16, cursor: "pointer", color: "#605e5c" }} 
                      onClick={() => setIsHelpOpen(false)}
                  />
                  <h2 style={{ marginTop: 0, fontSize: 18 }}>How Search Works</h2>
                  <div style={{ fontSize: 13, color: "#323130", lineHeight: "1.6" }}>
                      <p>Use the search interface to find emails you've filed previously across various locations.</p>
                      
                      <div style={{ marginBottom: 16 }}>
                          <strong>🔍 Search Bars:</strong>
                          <ul style={{ margin: "4px 0" }}>
                              <li><b>Filed Location:</b> Search for specific paths where emails were saved.</li>
                              <li><b>Keywords:</b> Searches across Subject, Sender, Recipients, and Path.</li>
                          </ul>
                      </div>
                      
                      <div style={{ marginBottom: 16 }}>
                          <strong>📂 Filtering:</strong>
                          <ul style={{ margin: "4px 0" }}>
                              <li><b>Date Range:</b> Quickly filter by the last month, 6 months, etc.</li>
                              <li><b>Including:</b> When enabled, keyword searches will also look inside filing <b>comments</b>.</li>
                              <li><b>Specific Fields:</b> Refine by From, To, CC, or Subject exact matches.</li>
                          </ul>
                      </div>
                      
                      <div style={{ marginBottom: 16 }}>
                          <strong>⚙️ Actions:</strong>
                          <ul style={{ margin: "4px 0" }}>
                              <li><b>Open:</b> Launch the filed email directly in Outlook.</li>
                              <li><b>Delete:</b> Permanently remove the filed file and clearing it from history.</li>
                          </ul>
                      </div>
                  </div>
                  <button 
                      onClick={() => setIsHelpOpen(false)}
                      style={{ width: "100%", padding: "8px", backgroundColor: "#0078d4", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", fontWeight: 600 }}
                  >
                      Got it!
                  </button>
              </div>
          </div>
      )}

      {/* ── Body: Sidebar + Results ──────────────────────────────────────── */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>

        {/* ── Left Sidebar Filters ──────────────────────────────────────── */}
        <div style={{
          width: 260, flexShrink: 0, backgroundColor: "#ffffff",
          borderRight: "1px solid #edebe9", display: "flex", flexDirection: "column",
        }}>
          <div style={{ 
            display: "flex", justifyContent: "space-between", alignItems: "center", 
            padding: "16px 16px 12px 16px" 
          }}>
            <span style={{ fontWeight: 600, fontSize: 14, color: "#323130" }}>Filter By</span>
            <div style={{ display: "flex", gap: 12, color: "#605e5c" }}>
              <MoreHorizontal20Regular style={{ cursor: "pointer" }} />
              <ChevronLeft20Regular style={{ cursor: "pointer" }} onClick={onClose} />
            </div>
          </div>

          <div style={{ flex: 1, overflowY: "auto", padding: "0 16px 16px 16px" }}>
            {/* Date Range Selector */}
            <div style={{ 
              marginBottom: 16, border: "1px solid #0078d4", borderRadius: 6, 
              padding: "10px 12px", display: "flex", alignItems: "center", gap: 10 
            }}>
                <CalendarMonth20Regular style={{ color: "#0078d4" }} />
                <select
                  value={dateRange}
                  onChange={e => setDateRange(e.target.value)}
                  style={{ border: "none", background: "none", outline: "none", fontSize: 13, fontWeight: 600, flex: 1, color: "#323130" }}
                >
                  {DATE_RANGES.map(r => <option key={r.value} value={r.value}>{r.label}</option>)}
                </select>
            </div>

            {/* Including Toggle */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <MailSettings20Regular style={{ color: "#ffb900" }} />
                    <span style={{ fontSize: 13, color: "#605e5c" }}>Including</span>
                </div>
                <div
                    onClick={() => setIsIncludingEnabled(!isIncludingEnabled)}
                    style={{
                      width: 32, height: 16, borderRadius: 8, cursor: "pointer",
                      backgroundColor: isIncludingEnabled ? "#0078d4" : "#c8c6c4",
                      position: "relative", transition: "background 0.2s",
                    }}
                >
                    <div style={{
                      position: "absolute", top: 2, left: isIncludingEnabled ? 18 : 2,
                      width: 12, height: 12, borderRadius: "50%",
                      backgroundColor: "#fff", transition: "left 0.2s",
                    }} />
                </div>
            </div>

            {/* Field Filters */}
            {[
                { label: "From", value: from, setter: setFrom, icon: <MailSettings20Regular style={{ color: "#0078d4" }} /> },
                { label: "To", value: to, setter: setTo, icon: <MailSettings20Regular style={{ color: "#0078d4" }} /> },
                { label: "CC", value: cc, setter: setCc, icon: <MailSettings20Regular style={{ color: "#0078d4" }} /> },
                { label: "Subject", value: subject, setter: setSubject, icon: <TextBulletList20Regular style={{ color: "#605e5c" }} /> },
                { label: "Body", icon: <TextBulletList20Regular style={{ color: "#605e5c" }} /> },
            ].map((f, idx) => (
                <div key={idx} style={{ marginBottom: 16 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                        {f.icon}
                        <span style={{ fontSize: 13, color: "#605e5c" }}>{f.label}</span>
                    </div>
                </div>
            ))}

            {/* Attachments Toggle */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <Attach20Regular style={{ color: "#605e5c" }} />
                    <span style={{ fontSize: 13, color: "#605e5c" }}>Attachments</span>
                </div>
                <div
                    onClick={() => setHasAttachments(!hasAttachments)}
                    style={{
                      width: 32, height: 16, borderRadius: 8, cursor: "pointer",
                      backgroundColor: hasAttachments ? "#0078d4" : "#c8c6c4",
                      position: "relative", transition: "background 0.2s",
                    }}
                >
                    <div style={{
                      position: "absolute", top: 2, left: hasAttachments ? 18 : 2,
                      width: 12, height: 12, borderRadius: "50%",
                      backgroundColor: "#fff", transition: "left 0.2s",
                    }} />
                </div>
            </div>

            {/* Search Types (bottom) */}
            <div style={{ marginTop: "auto", borderTop: "1px solid #edebe9", paddingTop: 16 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <Mail20Regular style={{ color: "#0078d4" }} />
                    <select 
                        value={selectedType}
                        onChange={e => setSelectedType(e.target.value)}
                        style={{ border: "none", background: "none", outline: "none", fontSize: 13, fontWeight: 600, color: "#323130" }}
                    >
                        <option value="emails">All Types</option>
                        <option value="files">Only Files</option>
                    </select>
                </div>
            </div>
          </div>
        </div>

        {/* ── Results Pane ─────────────────────────────────────────────── */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", backgroundColor: "#ffffff" }}>

          {/* Results Header */}
          <div style={{
            padding: "16px 20px", display: "flex", alignItems: "center", 
            justifyContent: "space-between", borderBottom: "1px solid #edebe9"
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontWeight: 600, fontSize: 16, color: "#323130" }}>Results</span>
              {results && (
                <span style={{ fontSize: 13, color: "#0078d4", fontWeight: 600 }}>
                  {results.count} {results.count === 1 ? "item" : "items"} found
                </span>
              )}
            </div>
            <div style={{ display: "flex", gap: 16, color: "#605e5c" }}>
                <ArrowClockwise20Regular style={{ cursor: "pointer" }} onClick={runSearch} />
                <MoreHorizontal20Regular style={{ cursor: "pointer" }} />
            </div>
          </div>

          <div style={{ flex: 1, overflowY: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead style={{ backgroundColor: "#ffffff", position: "sticky", top: 0, zIndex: 1 }}>
                <tr>
                  <th style={{ padding: "12px 20px", textAlign: "left", width: 40 }}>
                    <input 
                      type="checkbox" 
                      onChange={handleSelectAll}
                      checked={results?.results?.length > 0 && selectedRowIds.size === results.results.length}
                    />
                  </th>
                  <th style={thStyle}>Protocol</th>
                  <th style={thStyle}>Type</th>
                  <th style={thStyle}><Attach20Regular /></th>
                  <th style={thStyle}>
                      <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
                          Sent Date <ChevronDown20Regular style={{ fontSize: 12 }} />
                      </span>
                  </th>
                  <th style={thStyle}>From</th>
                  <th style={thStyle}>To</th>
                  <th style={{ width: 40 }}></th>
                </tr>
              </thead>
              <tbody style={{ fontSize: 13 }}>
                {Object.entries(grouped).map(([groupLabel, items]) => (
                  <React.Fragment key={groupLabel}>
                    <tr>
                      <td colSpan={8} style={{ padding: "16px 20px 8px 20px", fontWeight: 600, color: "#8a2a21", fontSize: 12 }}>
                        {groupLabel}
                      </td>
                    </tr>
                    {items.map(r => {
                      const protocol = getProtocol(r.filePath);
                      return (
                        <tr key={r.id} style={{ borderBottom: "1px solid #f3f3f3" }}>
                          <td style={{ padding: "10px 20px" }}>
                            <input 
                              type="checkbox" 
                              checked={selectedRowIds.has(r.id)}
                              onChange={() => handleSelectRow(r.id)}
                            />
                          </td>
                          <td style={tdStyle}>{protocol.icon}</td>
                          <td style={tdStyle}><Mail20Regular style={{ color: "#0078d4" }} /></td>
                          <td style={tdStyle}>
                            {r.hasAttachments && <Attach20Regular style={{ color: "#605e5c" }} />}
                          </td>
                          <td style={{ ...tdStyle, whiteSpace: "nowrap" }}>{formatDate(r.sentAt)}</td>
                          <td style={tdStyle} title={r.sender}>
                            <div style={{ maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                {r.sender}
                            </div>
                          </td>
                          <td style={tdStyle} title={Array.isArray(r.recipients) ? r.recipients.join(", ") : r.recipients}>
                            <div style={{ maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                {Array.isArray(r.recipients) ? r.recipients[0] : r.recipients}
                            </div>
                          </td>
                          <td style={{ paddingRight: 20, textAlign: "right" }}>
                              <div style={{ position: "relative", display: "inline-block" }}>
                                  <MoreHorizontal20Regular 
                                      style={{ color: "#605e5c", cursor: "pointer" }} 
                                      onClick={(e) => {
                                          e.stopPropagation();
                                          setActiveMenuId(activeMenuId === r.id ? null : r.id);
                                      }}
                                  />
                                  {activeMenuId === r.id && (
                                      <div style={{
                                          position: "absolute", top: 25, right: 0, zIndex: 100,
                                          backgroundColor: "#fff", borderRadius: 4, padding: "4px 0",
                                          boxShadow: "0 4px 12px rgba(0,0,0,0.15)", border: "1px solid #edebe9",
                                          minWidth: 100
                                      }}>
                                          <div 
                                              onClick={(e) => { e.stopPropagation(); handleOpenItem(r); }}
                                              style={{ 
                                                  padding: "8px 12px", cursor: "pointer", fontSize: 13, 
                                                  textAlign: "left", color: "#323130" 
                                              }}
                                              onMouseOver={e => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                                              onMouseOut={e => e.currentTarget.style.backgroundColor = ""}
                                          >Open</div>
                                          <div 
                                              onClick={(e) => { e.stopPropagation(); handleDeleteItem(r); }}
                                              style={{ 
                                                  padding: "8px 12px", cursor: "pointer", fontSize: 13, 
                                                  textAlign: "left", color: "#a4262c" 
                                              }}
                                              onMouseOver={e => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                                              onMouseOut={e => e.currentTarget.style.backgroundColor = ""}
                                          >Delete</div>
                                      </div>
                                  )}
                              </div>
                          </td>
                        </tr>
                      );
                    })}
                  </React.Fragment>
                ))}
              </tbody>
            </table>

            {/* Error / Loading / Placeholder states */}
            {loading && <div style={{ padding: 40, textAlign: "center", color: "#605e5c" }}>Searching...</div>}
            {!loading && !results && (
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100%", color: "#a19f9d" }}>
                <Search20Regular style={{ fontSize: 64, marginBottom: 16 }} />
                <span>Search for emails or locations above</span>
              </div>
            )}
          </div>
        </div>

        {/* Delete Confirmation Overlay */}
        {itemToDelete && (
            <div style={{
                position: "absolute", top: 0, left: 0, right: 0, bottom: 0,
                backgroundColor: "rgba(0,0,0,0.4)", display: "flex", justifyContent: "center",
                alignItems: "center", zIndex: 1000, borderRadius: 8
            }}>
                <div style={{
                    backgroundColor: "#fff", padding: 24, borderRadius: 8, maxWidth: 320,
                    boxShadow: "0 8px 32px rgba(0,0,0,0.2)", textAlign: "center"
                }}>
                    <h3 style={{ marginTop: 0, color: "#323130" }}>Confirm Delete</h3>
                    <p style={{ fontSize: 14, color: "#605e5c", lineHeight: "1.5" }}>
                        Are you sure you want to permanently delete this filed email from disk and our records?
                    </p>
                    <div style={{ marginTop: 24, display: "flex", gap: 12, justifyContent: "center" }}>
                        <button 
                            onClick={handleConfirmDelete}
                            style={{
                                padding: "8px 20px", borderRadius: 4, border: "none",
                                backgroundColor: "#a4262c", color: "#fff", cursor: "pointer",
                                fontWeight: 600
                            }}
                        >Delete</button>
                        <button 
                            onClick={() => setItemToDelete(null)}
                            style={{
                                padding: "8px 20px", borderRadius: 4, border: "1px solid #8a8886",
                                backgroundColor: "#fff", color: "#323130", cursor: "pointer",
                                fontWeight: 600
                            }}
                        >Cancel</button>
                    </div>
                </div>
            </div>
        )}
      </div>
    </div>
  );
}

const thStyle = {
  padding: "12px 10px", textAlign: "left", fontWeight: 600,
  fontSize: 12, color: "#605e5c", borderBottom: "1px solid #edebe9",
};

const tdStyle = {
  padding: "10px 10px", verticalAlign: "middle", color: "#323130",
};
