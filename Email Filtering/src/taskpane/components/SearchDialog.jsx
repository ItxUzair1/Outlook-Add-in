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
  ChevronRight20Regular,
  Desktop20Regular,
  Checkmark20Regular,
  MailSettings20Regular,
  ChevronDown20Regular,
  ArrowSync20Regular,
} from "@fluentui/react-icons";

import { API_BASE_URL } from "../services/backendApi.js";

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

/** Parent folder + file name for the Location column (full path in title). */
function formatFileLocation(filePath) {
  if (!filePath) return "—";
  const parts = filePath.split(/[\\/]/).filter(Boolean);
  if (parts.length === 0) return filePath;
  const file = parts[parts.length - 1];
  if (parts.length === 1) return file;
  const parent = parts[parts.length - 2];
  return `${parent} › ${file}`;
}

function parentDir(filePath) {
  if (!filePath) return "";
  const i = Math.max(filePath.lastIndexOf("\\"), filePath.lastIndexOf("/"));
  return i >= 0 ? filePath.slice(0, i) : "";
}

export default function SearchDialog({ onClose, onOpenSearchOptions }) {
  const [dateRange, setDateRange] = React.useState("6m");
  const [from, setFrom] = React.useState("");
  const [to, setTo] = React.useState("");
  const [cc, setCc] = React.useState("");
  const [subject, setSubject] = React.useState("");
  const [location, setLocation] = React.useState("");
  const [keywords, setKeywords] = React.useState("");
  const [attachmentFilter, setAttachmentFilter] = React.useState("any"); // any | with | without
  const [body, setBody] = React.useState("");
  const [isIncludingEnabled, setIsIncludingEnabled] = React.useState(false);
  const [selectedType, setSelectedType] = React.useState("emails");
  const [selectedRowIds, setSelectedRowIds] = React.useState(new Set());
  const [isHelpOpen, setIsHelpOpen] = React.useState(false);
  const [isSyncing, setIsSyncing] = React.useState(false);
  const [syncMessage, setSyncMessage] = React.useState("");
  const [activeMenuId, setActiveMenuId] = React.useState(null);
  const [itemToDelete, setItemToDelete] = React.useState(null);
  const [bulkDeleteRows, setBulkDeleteRows] = React.useState(null);
  const [filtersCollapsed, setFiltersCollapsed] = React.useState(false);

  const getSelectedResultRows = React.useCallback(() => {
    if (!results?.results?.length || selectedRowIds.size === 0) return [];
    return results.results.filter((r) => selectedRowIds.has(r.id));
  }, [results, selectedRowIds]);

  async function removeIndexEntry(id) {
    const enc = encodeURIComponent(id);
    const delResp = await fetch(`${API_BASE_URL}/api/search/${enc}`, { method: "DELETE" });
    return delResp.ok;
  }

  /**
   * @returns {"ok"|"removed"|"cancelled"|"missing"}
   */
  async function openIndexedFile(r, { bulk = false } = {}) {
    const resp = await fetch(`${API_BASE_URL}/api/search/open`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filePath: r.filePath }),
    });
    if (resp.status === 404) {
      if (bulk) return "missing";
      if (
        window.confirm(
          "The file was not found at its original location. It may have been moved or deleted.\n\nRemove this entry from search history?"
        )
      ) {
        if (await removeIndexEntry(r.id)) return "removed";
      }
      return "cancelled";
    }
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.error || "Could not open file");
    }
    return "ok";
  }

  /**
   * @returns {"ok"|"removed"|"cancelled"|"missing"}
   */
  async function openIndexedFolder(r, { bulk = false } = {}) {
    const resp = await fetch(`${API_BASE_URL}/api/search/open-folder`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filePath: r.filePath }),
    });
    if (resp.status === 404) {
      if (bulk) return "missing";
      if (
        window.confirm(
          "The folder was not found at its original location.\n\nRemove this entry from search history?"
        )
      ) {
        if (await removeIndexEntry(r.id)) return "removed";
      }
      return "cancelled";
    }
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.error || "Could not open folder");
    }
    return "ok";
  }
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
          const resp = await fetch(`${API_BASE_URL}/api/search/sync`, { method: "POST" });
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
      const out = await openIndexedFile(r, { bulk: false });
      if (out === "removed") await runSearch();
    } catch (err) {
      alert(`Open failed: ${err.message}`);
    }
    setActiveMenuId(null);
  };

  const handleOpenFolder = async (r) => {
    try {
      const out = await openIndexedFolder(r, { bulk: false });
      if (out === "removed") await runSearch();
    } catch (err) {
      alert(`Open folder failed: ${err.message}`);
    }
    setActiveMenuId(null);
  };

  const handleBulkOpen = async () => {
    const rows = getSelectedResultRows();
    if (rows.length === 0) return;
    if (rows.length > 1 && !window.confirm(`Open ${rows.length} files in their default applications?`)) return;
    setActiveMenuId(null);
    let removed = 0;
    const failures = [];
    for (const r of rows) {
      try {
        const out = await openIndexedFile(r, { bulk: true });
        if (out === "removed") removed++;
      } catch (e) {
        failures.push(e.message);
      }
    }
    if (removed) await runSearch();
    if (failures.length) alert(failures.slice(0, 5).join("\n") + (failures.length > 5 ? "\n…" : ""));
  };

  const handleBulkOpenFolders = async () => {
    const rows = getSelectedResultRows();
    if (rows.length === 0) return;
    const byDir = new Map();
    for (const r of rows) {
      const d = parentDir(r.filePath);
      if (d && !byDir.has(d)) byDir.set(d, r);
    }
    const reps = [...byDir.values()];
    if (reps.length === 0) return;
    if (reps.length > 1 && !window.confirm(`Open ${reps.length} unique folders in File Explorer?`)) return;
    setActiveMenuId(null);
    let removed = 0;
    const failures = [];
    for (const r of reps) {
      try {
        const out = await openIndexedFolder(r, { bulk: true });
        if (out === "removed") removed++;
      } catch (e) {
        failures.push(e.message);
      }
    }
    if (removed) await runSearch();
    if (failures.length) alert(failures.slice(0, 5).join("\n") + (failures.length > 5 ? "\n…" : ""));
  };

  const handleBulkDeleteClick = () => {
    const rows = getSelectedResultRows();
    if (rows.length === 0) return;
    setItemToDelete(null);
    setBulkDeleteRows(rows);
    setActiveMenuId(null);
  };

  const handleDeleteItem = (r) => {
    setBulkDeleteRows(null);
    setItemToDelete(r);
    setActiveMenuId(null);
  };

  const handleConfirmDelete = async () => {
    if (!itemToDelete) return;
    try {
      const encodedId = encodeURIComponent(itemToDelete.id);
      const resp = await fetch(`${API_BASE_URL}/api/search/${encodedId}`, { method: "DELETE" });
      if (resp.ok) {
        setSelectedRowIds((prev) => {
          const next = new Set(prev);
          next.delete(itemToDelete.id);
          return next;
        });
        await runSearch();
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

  const handleConfirmBulkDelete = async () => {
    if (!bulkDeleteRows?.length) return;
    let ok = 0;
    let fail = 0;
    const deletedIds = [];
    for (const r of bulkDeleteRows) {
      try {
        const enc = encodeURIComponent(r.id);
        const resp = await fetch(`${API_BASE_URL}/api/search/${enc}`, { method: "DELETE" });
        if (resp.ok) {
          ok++;
          deletedIds.push(r.id);
        } else fail++;
      } catch {
        fail++;
      }
    }
    setBulkDeleteRows(null);
    if (deletedIds.length) {
      setSelectedRowIds((prev) => {
        const next = new Set(prev);
        deletedIds.forEach((id) => next.delete(id));
        return next;
      });
    }
    await runSearch();
    if (fail > 0) alert(`Deleted ${ok} item(s). ${fail} failed.`);
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
      if (body.trim()) params.set("body", body.trim());
      if (location.trim()) params.set("location", location.trim());
      if (keywords.trim()) params.set("keywords", keywords.trim());
      if (attachmentFilter === "with") params.set("hasAttachments", "true");
      if (attachmentFilter === "without") params.set("hasAttachments", "false");
      if (isIncludingEnabled) params.set("including", "true");
      if (selectedType === "files") params.set("resultKind", "files");

      const resp = await fetch(`${API_BASE_URL}/api/search?${params.toString()}`);
      if (!resp.ok) {
        const raw = await resp.text();
        let msg = `Search failed (${resp.status} ${resp.statusText})`;
        try {
          const j = JSON.parse(raw);
          if (j.error) msg = j.error;
        } catch {
          if (raw?.trim()) msg = raw.trim().slice(0, 240);
        }
        throw new Error(msg);
      }
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
    setBody("");
    setLocation("");
    setKeywords("");
    setAttachmentFilter("any");
    setIsIncludingEnabled(false);
    setSelectedType("emails");
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
          <Settings20Regular style={{ cursor: "pointer" }} onClick={onOpenSearchOptions} title="Search Options" />
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
          <Dismiss20Regular
            style={{ cursor: "pointer", color: "#605e5c", marginLeft: 4 }}
            onClick={onClose}
            title="Close search"
          />
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
                              <li><b>Keywords:</b> Searches Subject, Sender, Recipients, filed Path, and indexed message body.</li>
                          </ul>
                      </div>
                      
                      <div style={{ marginBottom: 16 }}>
                          <strong>📂 Filtering:</strong>
                          <ul style={{ margin: "4px 0" }}>
                              <li><b>Date Range:</b> Quickly filter by the last month, 6 months, etc.</li>
                              <li><b>Including:</b> When enabled, keyword searches also match your filing <b>comments</b>.</li>
                              <li><b>Attachments:</b> Filter to rows with or without attachments, or leave as Any.</li>
                              <li><b>All Types / Only Files:</b> Limit results to saved non-message files (e.g. attachments) vs all index rows.</li>
                              <li><b>Specific Fields:</b> Refine by From, To, CC, Subject, or Body (substring match).</li>
                          </ul>
                      </div>
                      
                      <div style={{ marginBottom: 16 }}>
                          <strong>⚙️ Results:</strong>
                          <ul style={{ margin: "4px 0" }}>
                              <li><b>Location</b> shows the save folder and file name; hover for the full path.</li>
                              <li>Click <b>⋯</b> on a row for <b>Open</b>, <b>Open folder</b>, or <b>Delete</b>.</li>
                              <li>Select rows with checkboxes, then use <b>Open selected</b>, <b>Open folders</b>, or <b>Delete selected</b>.</li>
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

        {/* ── Left Sidebar Filters (chevron < hides panel; strip with > restores) ── */}
        {filtersCollapsed ? (
          <div
            style={{
              width: 44,
              flexShrink: 0,
              backgroundColor: "#ffffff",
              borderRight: "1px solid #edebe9",
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              paddingTop: 16,
            }}
          >
            <ChevronRight20Regular
              style={{ color: "#605e5c", cursor: "pointer" }}
              title="Show filters"
              onClick={() => setFiltersCollapsed(false)}
            />
          </div>
        ) : (
        <div style={{
          width: 260, flexShrink: 0, backgroundColor: "#ffffff",
          borderRight: "1px solid #edebe9", display: "flex", flexDirection: "column",
        }}>
          <div style={{ 
            display: "flex", justifyContent: "space-between", alignItems: "center", 
            padding: "16px 16px 12px 16px" 
          }}>
            <span style={{ fontWeight: 600, fontSize: 14, color: "#323130" }}>Filter By</span>
            <ChevronLeft20Regular
              style={{ color: "#605e5c", cursor: "pointer" }}
              title="Hide filters"
              onClick={() => setFiltersCollapsed(true)}
            />
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
                { label: "Subject", value: subject, setter: setSubject, icon: <TextBulletList20Regular style={{ color: "#0078d4" }} /> },
                { label: "Body", value: body, setter: setBody, icon: <TextBulletList20Regular style={{ color: "#ffb900" }} /> },
            ].map((f, idx) => (
                <div key={idx} style={{ marginBottom: 16 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                        {f.icon}
                        <span style={{ fontSize: 13, color: "#605e5c" }}>{f.label}</span>
                    </div>
                    {f.setter && (
                        <div style={{ backgroundColor: "#f3f2f1", borderRadius: 4, padding: "4px 8px", border: "1px solid transparent" }}>
                            <input
                                value={f.value}
                                onChange={e => f.setter(e.target.value)}
                                onKeyDown={e => e.key === "Enter" && runSearch()}
                                style={{ border: "none", background: "transparent", outline: "none", width: "100%", fontSize: 12, fontFamily: "Segoe UI" }}
                                placeholder={`Enter ${f.label.toLowerCase()}...`}
                            />
                        </div>
                    )}
                </div>
            ))}

            {/* Attachments filter */}
            <div style={{ marginBottom: 20 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                    <Attach20Regular style={{ color: "#605e5c" }} />
                    <span style={{ fontSize: 13, color: "#605e5c" }}>Attachments</span>
                </div>
                <select
                    value={attachmentFilter}
                    onChange={e => setAttachmentFilter(e.target.value)}
                    style={{
                        width: "100%", fontSize: 13, padding: "6px 8px", borderRadius: 4,
                        border: "1px solid #edebe9", backgroundColor: "#f3f2f1", color: "#323130",
                        fontFamily: "Segoe UI",
                    }}
                >
                    <option value="any">Any</option>
                    <option value="with">With attachments</option>
                    <option value="without">Without attachments</option>
                </select>
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
                        <option value="emails">All types</option>
                        <option value="files">Only files (non-.msg/.eml)</option>
                    </select>
                </div>
            </div>
          </div>
        </div>
        )}

        {/* ── Results Pane (minWidth:0 so flex does not block horizontal scroll) ── */}
        <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column", backgroundColor: "#ffffff" }}>

          {/* Results Header */}
          <div style={{
            padding: "16px 20px", display: "flex", alignItems: "center", 
            justifyContent: "space-between", borderBottom: "1px solid #edebe9", flexShrink: 0,
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", minWidth: 0 }}>
              <span style={{ fontWeight: 600, fontSize: 16, color: "#323130" }}>Results</span>
              {results && (
                <span style={{ fontSize: 13, color: "#0078d4", fontWeight: 600 }}>
                  {results.count} {results.count === 1 ? "item" : "items"} found
                </span>
              )}
            </div>
            <ArrowClockwise20Regular
              style={{ cursor: "pointer", color: "#605e5c", flexShrink: 0 }}
              onClick={runSearch}
              title="Refresh results"
            />
          </div>

          {error && (
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                gap: 10,
                margin: "0 16px 8px 16px",
                padding: "10px 12px",
                backgroundColor: "#fde7e9",
                border: "1px solid #a4262c",
                borderRadius: 4,
                color: "#323130",
                fontSize: 13,
                flexShrink: 0,
              }}
              role="alert"
            >
              <span style={{ flex: 1, lineHeight: 1.4 }}>{error}</span>
              <Dismiss20Regular
                style={{ cursor: "pointer", color: "#a4262c", flexShrink: 0 }}
                title="Dismiss"
                onClick={() => setError("")}
              />
            </div>
          )}

          {results && selectedRowIds.size > 0 && (
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                alignItems: "center",
                gap: 8,
                padding: "8px 16px 12px 16px",
                borderBottom: "1px solid #edebe9",
                backgroundColor: "#f3f2f1",
                flexShrink: 0,
              }}
            >
              <span style={{ fontSize: 13, fontWeight: 600, color: "#323130", marginRight: 4 }}>
                {selectedRowIds.size} selected
              </span>
              <button type="button" onClick={handleBulkOpen} style={bulkBtnPrimary}>
                Open selected
              </button>
              <button type="button" onClick={handleBulkOpenFolders} style={bulkBtnSecondary}>
                Open folders
              </button>
              <button type="button" onClick={handleBulkDeleteClick} style={bulkBtnDanger}>
                Delete selected
              </button>
            </div>
          )}

          <div
            style={{
              flex: 1,
              minHeight: 0,
              minWidth: 0,
              overflowY: "auto",
              overflowX: "scroll",
              WebkitOverflowScrolling: "touch",
            }}
            className="search-results-scroll"
          >
            <style>{`
              .search-results-scroll { scrollbar-width: thin; scrollbar-color: #c8c6c4 #f3f2f1; }
              .search-results-scroll::-webkit-scrollbar { height: 12px; width: 12px; }
              .search-results-scroll::-webkit-scrollbar-thumb { background: #c8c6c4; border-radius: 6px; }
              .search-results-scroll::-webkit-scrollbar-track { background: #f3f2f1; }
            `}</style>
            {/* No width:100% — otherwise the table shrinks and never gains a horizontal scrollbar */}
            <table style={{ minWidth: 1040, width: "max-content", borderCollapse: "collapse", tableLayout: "auto" }}>
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
                  <th style={{ ...thStyle, minWidth: 160 }}>From</th>
                  <th style={{ ...thStyle, minWidth: 160 }}>To</th>
                  <th style={{ ...thStyle, minWidth: 240 }}>Location</th>
                  <th style={{ minWidth: 48, width: 48, padding: "12px 8px", textAlign: "right", fontSize: 12, fontWeight: 600, color: "#605e5c", borderBottom: "1px solid #edebe9" }} aria-label="Open or delete" title="⋯ Open / Delete">⋯</th>
                </tr>
              </thead>
              <tbody style={{ fontSize: 13 }}>
                {Object.entries(grouped).map(([groupLabel, items]) => (
                  <React.Fragment key={groupLabel}>
                    <tr>
                      <td colSpan={9} style={{ padding: "16px 20px 8px 20px", fontWeight: 600, color: "#8a2a21", fontSize: 12 }}>
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
                          <td style={tdStyle}>
                              {r.filePath.toLowerCase().endsWith(".eml") || r.filePath.toLowerCase().endsWith(".msg") 
                                  ? <Mail20Regular style={{ color: "#0078d4" }} /> 
                                  : <Attach20Regular style={{ color: "#ffb900" }} />
                              }
                          </td>
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
                          <td style={tdStyle} title={r.filePath}>
                            <div style={{ maxWidth: 320, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "flex", alignItems: "center", gap: 6 }}>
                                <FolderOpen20Regular style={{ fontSize: 16, color: "#605e5c", flexShrink: 0 }} />
                                <span style={{ fontSize: 12, color: "#323130" }}>{formatFileLocation(r.filePath)}</span>
                            </div>
                          </td>
                          <td style={{ ...tdStyle, textAlign: "right", paddingRight: 16, whiteSpace: "nowrap", verticalAlign: "middle" }}>
                              <div style={{ position: "relative", display: "inline-block" }}>
                                  <MoreHorizontal20Regular 
                                      style={{ color: "#605e5c", cursor: "pointer" }} 
                                      title="Open, open folder, delete"
                                      onClick={(e) => {
                                          e.stopPropagation();
                                          setActiveMenuId(activeMenuId === r.id ? null : r.id);
                                      }}
                                  />
                                  {activeMenuId === r.id && (
                                      <div style={{
                                          position: "absolute", top: 26, right: 0, zIndex: 200,
                                          backgroundColor: "#fff", borderRadius: 4, padding: "4px 0",
                                          boxShadow: "0 4px 12px rgba(0,0,0,0.15)", border: "1px solid #edebe9",
                                          minWidth: 140
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
                                              onClick={(e) => { e.stopPropagation(); handleOpenFolder(r); }}
                                              style={{ 
                                                  padding: "8px 12px", cursor: "pointer", fontSize: 13, 
                                                  textAlign: "left", color: "#323130" 
                                              }}
                                              onMouseOver={e => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                                              onMouseOut={e => e.currentTarget.style.backgroundColor = ""}
                                          >Open folder</div>
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
            {!loading && results && results.count === 0 && (
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100%", color: "#a19f9d", padding: 40 }}>
                <Dismiss20Regular style={{ fontSize: 48, marginBottom: 16, color: "#a4262c" }} />
                <span style={{ fontWeight: 600, color: "#323130" }}>No results found</span>
                <span style={{ fontSize: 13, marginTop: 4 }}>Try adjusting your filters or keywords</span>
              </div>
            )}
          </div>
        </div>

        {/* Delete Confirmation Overlay (single or bulk) */}
        {(itemToDelete || (bulkDeleteRows && bulkDeleteRows.length > 0)) && (
            <div style={{
                position: "absolute", top: 0, left: 0, right: 0, bottom: 0,
                backgroundColor: "rgba(0,0,0,0.4)", display: "flex", justifyContent: "center",
                alignItems: "center", zIndex: 1000, borderRadius: 8
            }}>
                <div style={{
                    backgroundColor: "#fff", padding: 24, borderRadius: 8, maxWidth: 400,
                    boxShadow: "0 8px 32px rgba(0,0,0,0.2)", textAlign: "center"
                }}>
                    <h3 style={{ marginTop: 0, color: "#323130" }}>Confirm Delete</h3>
                    <p style={{ fontSize: 14, color: "#605e5c", lineHeight: "1.5" }}>
                        {bulkDeleteRows?.length
                          ? `Permanently delete ${bulkDeleteRows.length} file(s) from disk and remove them from search history?`
                          : "Are you sure you want to permanently delete this filed email from disk and our records?"}
                    </p>
                    <div style={{ marginTop: 24, display: "flex", gap: 12, justifyContent: "center" }}>
                        <button 
                            onClick={bulkDeleteRows?.length ? handleConfirmBulkDelete : handleConfirmDelete}
                            style={{
                                padding: "8px 20px", borderRadius: 4, border: "none",
                                backgroundColor: "#a4262c", color: "#fff", cursor: "pointer",
                                fontWeight: 600
                            }}
                        >Delete</button>
                        <button 
                            onClick={() => { setItemToDelete(null); setBulkDeleteRows(null); }}
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

const bulkBtnPrimary = {
  padding: "6px 14px",
  borderRadius: 4,
  border: "none",
  backgroundColor: "#0078d4",
  color: "#fff",
  cursor: "pointer",
  fontWeight: 600,
  fontSize: 12,
  fontFamily: "Segoe UI, sans-serif",
};

const bulkBtnSecondary = {
  ...bulkBtnPrimary,
  backgroundColor: "#fff",
  color: "#323130",
  border: "1px solid #8a8886",
};

const bulkBtnDanger = {
  ...bulkBtnPrimary,
  backgroundColor: "#a4262c",
};
