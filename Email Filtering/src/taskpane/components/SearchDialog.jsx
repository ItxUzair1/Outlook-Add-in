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
  Calendar20Regular,
} from "@fluentui/react-icons";

import { API_BASE_URL } from "../services/backendApi.js";



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

function renderHighlightedText(text, keyword) {
  if (!text) return "";
  if (!keyword || !keyword.trim()) return text;

  const normalizedKeyword = keyword.trim();
  const escapedKeyword = normalizedKeyword.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  const regex = new RegExp(`(${escapedKeyword})`, "gi");
  const parts = text.split(regex);

  return parts.map((part, index) => {
    return regex.test(part) ? (
      <mark key={index} style={{ backgroundColor: "#ffeb3b", color: "#000", padding: "0 2px", borderRadius: "2px" }}>
        {part}
      </mark>
    ) : (
      part
    );
  });
}

const getSavedFilter = (key, fallback) => {
  try {
    const saved = localStorage.getItem("koyomail_last_search_filters");
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed[key] !== undefined) {
        if (key === "searchScope" && parsed[key] === "personal_only") {
          return "all_personal";
        }
        return parsed[key];
      }
    }
  } catch (e) {}
  return fallback;
};

export default function SearchDialog({ onClose, onOpenSearchOptions }) {

  const [from, setFrom] = React.useState(() => getSavedFilter("from", ""));
  const [to, setTo] = React.useState(() => getSavedFilter("to", ""));
  const [cc, setCc] = React.useState(() => getSavedFilter("cc", ""));
  const [subject, setSubject] = React.useState(() => getSavedFilter("subject", ""));
  const [location, setLocation] = React.useState(() => getSavedFilter("location", ""));
  const [keywords, setKeywords] = React.useState(() => getSavedFilter("keywords", ""));
  const [attachmentFilter, setAttachmentFilter] = React.useState(() => getSavedFilter("attachmentFilter", "any")); // any | with | without
  const [dateFilter, setDateFilter] = React.useState(() => getSavedFilter("dateFilter", "all"));
  const [selectedType, setSelectedType] = React.useState(() => getSavedFilter("selectedType", "emails"));
  const [selectedRowIds, setSelectedRowIds] = React.useState(new Set());
  const [previewItem, setPreviewItem] = React.useState(null);
  const [previewBodyLoadingId, setPreviewBodyLoadingId] = React.useState(null);
  const [previewBodyError, setPreviewBodyError] = React.useState(null);
  const previewBodyCache = React.useRef(new Map());
  const [isHelpOpen, setIsHelpOpen] = React.useState(false);

  const [activeMenuId, setActiveMenuId] = React.useState(null);
  const [itemToDelete, setItemToDelete] = React.useState(null);
  const [bulkDeleteRows, setBulkDeleteRows] = React.useState(null);
  const [filtersCollapsed, setFiltersCollapsed] = React.useState(false);
  const [options, setOptions] = React.useState({ enableSearching: true, disableDelete: false, disableMoveTo: false });
  const [timeSpan, setTimeSpan] = React.useState(() => getSavedFilter("timeSpan", "all_time"));
  const [isLocationDropdownOpen, setIsLocationDropdownOpen] = React.useState(false);


  const [moveTargetItem, setMoveTargetItem] = React.useState(null);
  const [moveDestinationPath, setMoveDestinationPath] = React.useState("");
  
  React.useEffect(() => {
    const loadOptions = () => {
      try {
        const stored = localStorage.getItem('koyomail_options');
        if (stored) {
          const parsed = JSON.parse(stored);
          setOptions(parsed);
        }
      } catch (e) {
        console.error("Could not load options", e);
      }
    };
    
    loadOptions();
    window.addEventListener('koyomail_options_updated', loadOptions);
    return () => window.removeEventListener('koyomail_options_updated', loadOptions);
  }, []);

  React.useEffect(() => {
    try {
      const filters = {
        from,
        to,
        cc,
        subject,
        location,
        keywords,
        attachmentFilter,
        dateFilter,
        selectedType,
        timeSpan,
      };
      localStorage.setItem("koyomail_last_search_filters", JSON.stringify(filters));
    } catch (e) {
      console.error("Failed to save search filters", e);
    }
  }, [from, to, cc, subject, location, keywords, attachmentFilter, dateFilter, selectedType, timeSpan]);

  function getSearchUserEmail() {
    return new URLSearchParams(window.location.search).get("userEmail") || "";
  }

  const fetchPreviewBody = React.useCallback(async (id) => {
    if (previewBodyCache.current.has(id)) {
      return previewBodyCache.current.get(id);
    }
    const params = new URLSearchParams({ id });
    const userEmail = getSearchUserEmail();
    if (userEmail) params.set("userEmail", userEmail);
    const resp = await fetch(`${API_BASE_URL}/api/search/preview?${params.toString()}`);
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.error || `Preview failed (${resp.status})`);
    }
    const data = await resp.json();
    const bodyText = data.body || "";
    previewBodyCache.current.set(id, bodyText);
    return bodyText;
  }, []);

  const prefetchPreviewBody = React.useCallback((id) => {
    if (!id || previewBodyCache.current.has(id)) return;
    fetchPreviewBody(id).catch(() => {});
  }, [fetchPreviewBody]);

  const handlePreviewRow = React.useCallback(async (row) => {
    const isCached = previewBodyCache.current.has(row.id);
    setPreviewItem({ ...row, body: isCached ? previewBodyCache.current.get(row.id) : null });
    setPreviewBodyError(null);

    if (isCached) return;

    setPreviewBodyLoadingId(row.id);
    try {
      const bodyText = await fetchPreviewBody(row.id);
      setPreviewItem((prev) => (prev && prev.id === row.id ? { ...prev, body: bodyText } : prev));
    } catch (e) {
      setPreviewBodyError(e.message);
    } finally {
      setPreviewBodyLoadingId((current) => (current === row.id ? null : current));
    }
  }, [fetchPreviewBody]);

  const skipServerFilterRefresh = React.useRef(true);
  // Dropdown scope changes do NOT auto-trigger a search.
  // The user must click "Search" or press Enter to run a new query.

  // Re-run search when attachment or timeSpan filters change (server-side filters).
  React.useEffect(() => {
    if (skipServerFilterRefresh.current) {
      skipServerFilterRefresh.current = false;
      return;
    }
    if (results === null) return;
    if (!keywords.trim() && !location.trim()) return;
    runSearch();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [attachmentFilter, timeSpan]);

  React.useEffect(() => {
    const handleDocumentClick = () => {
      setActiveMenuId(null);
    };
    document.addEventListener("click", handleDocumentClick);
    return () => document.removeEventListener("click", handleDocumentClick);
  }, []);

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
  const [loadingMore, setLoadingMore] = React.useState(false);
  const isSearchBusy = loading || loadingMore;
  const [error, setError] = React.useState("");





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

  const handleCopyItem = async (r, e) => {
    if (e) {
      e.stopPropagation();
      const target = e.currentTarget;
      const originalText = target.innerText;
      target.innerText = "Copied!";
      setTimeout(() => {
        if (target) target.innerText = originalText;
        setActiveMenuId(null);
      }, 1000);
    } else {
      setActiveMenuId(null);
    }
    
    try {
      const resp = await fetch(`${API_BASE_URL}/api/search/copy`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filePath: r.filePath }),
      });
      if (!resp.ok) {
        const data = await resp.json().catch(() => ({}));
        throw new Error(data.error || "Could not copy file");
      }
    } catch (err) {
      alert(`Copy failed: ${err.message}`);
    }
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

  const handleMoveItem = (r) => {
    setActiveMenuId(null);
    setMoveTargetItem(r);
    setMoveDestinationPath("");
  };

  const movePathInputRef = React.useRef(null);

  const handleBrowseFolder = async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/api/search/browse-folder`);
      if (!resp.ok) {
        throw new Error("Unable to open folder picker");
      }
      const data = await resp.json();
      if (data?.path) {
        setMoveDestinationPath(String(data.path).trim());
        // Force WebView2 to completely repaint after native dialog closes
        setTimeout(() => {
          if (movePathInputRef.current) {
            movePathInputRef.current.blur();
            movePathInputRef.current.focus();
          }
          window.dispatchEvent(new Event('resize'));
        }, 150);
      }
    } catch (err) {
      alert(`Browse failed: ${err.message}`);
    }
  };

  const handlePasteFolder = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        setMoveDestinationPath(text.trim());
      }
    } catch (err) {
      console.error("Failed to read clipboard:", err);
    }
  };

  const submitMoveItem = async () => {
    const r = moveTargetItem;
    const destDir = moveDestinationPath.trim();
    if (!r || !destDir) return;
    
    setMoveTargetItem(null);

    try {
      const resp = await fetch(`${API_BASE_URL}/api/search/move`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: r.id, destinationDir: destDir })
      });
      if (resp.ok) {
         await runSearch();
      } else {
         const data = await resp.json();
         alert(`Move failed: ${data.error}`);
      }
    } catch (e) {
      alert(`Move failed: ${e.message}`);
    }
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

  async function fetchSearchPage(offset, { forceDisk = false } = {}) {
    const params = new URLSearchParams();
    params.set("offset", String(offset));
    params.set("limit", "50");

    if (location.trim()) params.set("location", location.trim());
    if (keywords.trim()) params.set("keywords", keywords.trim());
    if (attachmentFilter === "with") params.set("hasAttachments", "true");
    if (attachmentFilter === "without") params.set("hasAttachments", "false");
    if (selectedType === "files") params.set("resultKind", "files");
    if (timeSpan && timeSpan !== "all_time") params.set("timeSpan", timeSpan);
    params.set("includeBody", "true"); // Always include body
    if (forceDisk) params.set("forceDynamicScan", "true");

    const userEmail = getSearchUserEmail();
    if (userEmail) params.set("userEmail", userEmail);

    const resp = await fetch(`${API_BASE_URL}/api/search?${params.toString()}`);
    if (!resp.ok) {
      const raw = await resp.text();
      let msg = `Search failed (${resp.status} ${resp.statusText})`;
      try {
        const j = JSON.parse(raw);
        if (j.code === "EMPTY_QUERY" || j.error) msg = j.error;
      } catch {
        if (raw?.trim()) msg = raw.trim().slice(0, 240);
      }
      throw new Error(msg);
    }
    return resp.json();
  }

  async function runSearch({ forceDisk = false } = {}) {
    if (isSearchBusy) return;
    setLoading(true);
    setError("");
    try {
      const hasAnyInput = location.trim() || keywords.trim();
      if (!hasAnyInput) {
        setError("Please enter a keyword or location to search.");
        setLoading(false);
        return;
      }

      const data = await fetchSearchPage(0, { forceDisk });
      previewBodyCache.current.clear();
      setPreviewItem(null);
      setPreviewBodyError(null);
      setResults(data);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  async function loadMoreResults() {
    if (isSearchBusy || !results?.hasMore) return;
    setLoadingMore(true);
    setError("");
    try {
      const offset = results.results?.length ?? 0;
      const data = await fetchSearchPage(offset);
      setResults((prev) => {
        const merged = [...(prev.results || []), ...(data.results || [])];
        return {
          ...data,
          results: merged,
          count: merged.length,
          loadedCount: merged.length,
          hasMore: data.hasMore,
          estimatedTotalHits: data.estimatedTotalHits,
        };
      });
    } catch (e) {
      setError(e.message);
    } finally {
      setLoadingMore(false);
    }
  }

  function clearFilters() {
    setFrom("");
    setTo("");
    setCc("");
    setSubject("");
    setLocation("");
    setKeywords("");
    setAttachmentFilter("any");
    setDateFilter("all");
    setIncludeBodyInSearch(false);
    setSelectedType("emails");
    setSelectedRowIds(new Set());
    setPreviewItem(null);
    setPreviewBodyError(null);
    previewBodyCache.current.clear();
    setResults(null);
    setError("");
    skipServerFilterRefresh.current = true;
  }

  const handleSelectAll = (e) => {
    if (visibleResults.length > 0 && selectedRowIds.size > 0 && selectedRowIds.size < visibleResults.length) {
      setSelectedRowIds(new Set());
    } else if (e.target.checked && visibleResults.length > 0) {
      setSelectedRowIds(new Set(visibleResults.map(r => r.id)));
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

  // ── Client-side post-filters (From, To, CC, Subject) ──────────────
  // Date range and attachments are filtered on the server when Search runs.
  const visibleResults = React.useMemo(() => {
    if (!results?.results) return [];
    let filtered = results.results;
    if (from.trim()) {
      const q = from.trim().toLowerCase();
      filtered = filtered.filter(r => (r.sender || "").toLowerCase().includes(q));
    }
    if (to.trim()) {
      const q = to.trim().toLowerCase();
      filtered = filtered.filter(r => (r.recipients || "").toLowerCase().includes(q));
    }
    if (cc.trim()) {
      const q = cc.trim().toLowerCase();
      filtered = filtered.filter(r => (r.cc || "").toLowerCase().includes(q));
    }
    if (subject.trim()) {
      const q = subject.trim().toLowerCase();
      filtered = filtered.filter(r => (r.subject || "").toLowerCase().includes(q));
    }
    return filtered.sort((a, b) => {
      const ta = new Date(a.sentAt || a.filedAt || 0).getTime();
      const tb = new Date(b.sentAt || b.filedAt || 0).getTime();
      if (dateFilter === "oldest_first") return ta - tb;
      return tb - ta;
    });
  }, [results, from, to, cc, subject, dateFilter]);

  const grouped = results ? groupByRelativeDate(visibleResults) : {};

  if (!options.enableSearching) {
    return (
      <div style={{
        display: "flex", flexDirection: "column", height: "100vh",
        fontFamily: "Segoe UI, sans-serif", backgroundColor: "#f8f8f8",
        alignItems: "center", justifyContent: "center",
      }}>
        <Dismiss20Regular style={{ fontSize: 48, marginBottom: 16, color: "#605e5c" }} />
        <span style={{ fontWeight: 600, color: "#323130", fontSize: 18 }}>Search is Disabled</span>
        <span style={{ fontSize: 14, color: "#605e5c", marginTop: 8 }}>You can enable searching from the Options window.</span>
        <div style={{ marginTop: 24, display: "flex", gap: 12 }}>
           <button 
               onClick={onOpenSearchOptions}
               style={{ padding: "8px 20px", borderRadius: 4, border: "1px solid #0078d4", backgroundColor: "#0078d4", color: "#fff", cursor: "pointer", fontWeight: 600 }}
           >Open Options</button>
           <button 
               onClick={onClose}
               style={{ padding: "8px 20px", borderRadius: 4, border: "1px solid #8a8886", backgroundColor: "#fff", color: "#323130", cursor: "pointer", fontWeight: 600 }}
           >Close</button>
        </div>
      </div>
    );
  }

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
          border: "1px solid transparent", position: "relative",
        }}>
          <FolderOpen20Regular style={{ color: "#0078d4" }} />
          <input
            placeholder="Search By Filed Location"
            value={location}
            onChange={e => {
              setLocation(e.target.value);
              setIsLocationDropdownOpen(true);
            }}
            onFocus={() => {
              setIsLocationDropdownOpen(true);
            }}
            onBlur={() => setTimeout(() => setIsLocationDropdownOpen(false), 200)}
            onKeyDown={e => e.key === "Enter" && !isSearchBusy && runSearch({ forceDisk: true })}
            style={{ border: "none", background: "transparent", outline: "none", flex: 1, fontSize: 13, fontFamily: "Segoe UI" }}
          />
          <ChevronDown20Regular 
            style={{ color: "#605e5c", cursor: "pointer" }} 
            onClick={() => setIsLocationDropdownOpen(!isLocationDropdownOpen)}
          />

          {isLocationDropdownOpen && (
            <div style={{
              position: "absolute",
              top: "100%", left: 0, right: 0,
              backgroundColor: "#fff",
              border: "1px solid #edebe9",
              boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
              borderRadius: 4,
              maxHeight: 250,
              overflowY: "auto",
              zIndex: 1000,
              marginTop: 4,
            }}>
              {scopePaths.filter(p => p.toLowerCase().includes(location.toLowerCase())).map((p, i) => (
                <div 
                  key={i} 
                  style={{ padding: "8px 12px", cursor: "pointer", fontSize: 13, color: "#323130", wordBreak: "break-all" }}
                  onMouseEnter={e => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                  onMouseLeave={e => e.currentTarget.style.backgroundColor = "transparent"}
                  onMouseDown={(e) => {
                    e.preventDefault(); // Prevents input from losing focus
                    setLocation(p);
                    setIsLocationDropdownOpen(false);
                  }}
                >
                  {p}
                </div>
              ))}
              {scopePaths.filter(p => p.toLowerCase().includes(location.toLowerCase())).length === 0 && (
                <div style={{ padding: "8px 12px", fontSize: 13, color: "#a19f9d", fontStyle: "italic" }}>
                  No matching locations...
                </div>
              )}
            </div>
          )}
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
            onKeyDown={e => e.key === "Enter" && !isSearchBusy && runSearch({ forceDisk: true })}
            style={{ border: "none", background: "transparent", outline: "none", flex: 1, fontSize: 13, fontFamily: "Segoe UI" }}
          />
          <ArrowCounterclockwise20Regular 
              style={{ color: "#605e5c", cursor: "pointer" }} 
              onClick={() => { setKeywords(""); setLocation(""); }}
          />
        </div>

        {/* Actions */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, color: "#605e5c" }}>

          <Settings20Regular style={{ cursor: "pointer" }} onClick={onOpenSearchOptions} title="Search Options" />
          <QuestionCircle20Regular 
              style={{ cursor: "pointer" }} 
              onClick={() => setIsHelpOpen(true)}
              title="Help and Search Guide"
          />
          <button onClick={clearFilters}
            style={{ 
              background: "#fff", border: "1px solid #8a8886", borderRadius: 4, 
              padding: "5px 15px", color: "#323130", cursor: "pointer", 
              fontSize: 13, fontWeight: 600, fontFamily: "Segoe UI",
              display: "flex", alignItems: "center", gap: 4 
            }}>
            Clear
          </button>
          <button onClick={() => runSearch({ forceDisk: true })}
            disabled={isSearchBusy}
            style={{ 
              background: isSearchBusy ? "#c8c6c4" : "#0078d4", border: "none", borderRadius: 4, 
              padding: "6px 16px", color: "#fff", cursor: isSearchBusy ? "not-allowed" : "pointer", 
              fontSize: 13, fontWeight: 600, fontFamily: "Segoe UI",
              display: "flex", alignItems: "center", gap: 4 
            }}>
            {loading ? "Searching…" : "Search"}
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
                              <li><b>Keywords:</b> Searches Subject, Sender, To, Cc, filed Path, and comments by default.</li>
                              <li><b>Include email body:</b> Optional checkbox to also search inside message text (slower). Preview still shows the full body when you click a result.</li>
                          </ul>
                      </div>
                      
                      <div style={{ marginBottom: 16 }}>
                          <strong>📂 Filtering:</strong>
                          <ul style={{ margin: "4px 0" }}>
                              <li><b>Date Range:</b> Applied on the server when you click Search (re-searches automatically if you change it after results load).</li>
                              <li><b>Attachments:</b> Applied on the server when you click Search.</li>
                              <li><b>Search scope:</b> &quot;Locations I Use&quot; is fastest. &quot;Search All Locations&quot; searches the entire index and may be slower.</li>
                              <li><b>All Types / Only Files:</b> Limit results to saved non-message files (e.g. attachments) vs all index rows.</li>
                              <li><b>Specific Fields:</b> Refine by From, To, CC, or Subject (substring match on loaded results).</li>
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
            {/* Field Filters — active only after a search has returned results */}
            {results === null && (
              <div style={{
                marginBottom: 16, padding: "8px 10px", backgroundColor: "#f3f2f1",
                borderRadius: 4, fontSize: 11, color: "#8a8886", lineHeight: "1.4"
              }}>
                Set date and attachments below, then search. From / To / Subject narrow results after they load.
              </div>
            )}

            {/* Search Scope Selector */}
            <div style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                    <FolderOpen20Regular style={{ color: "#0078d4" }} />
                    <span style={{ fontSize: 13, color: "#605e5c", fontWeight: 600 }}>Search Scope</span>
                </div>
                <select
                    value={searchScope}
                    onChange={e => setSearchScope(e.target.value)}
                    style={{
                        width: "100%", fontSize: 13, padding: "6px 8px", borderRadius: 4,
                        border: "1px solid #edebe9", backgroundColor: "#f3f2f1", color: "#323130",
                        fontFamily: "Segoe UI", fontWeight: 600
                    }}
                >
                    <option value="locations_i_use">Locations I Use (All)</option>
                    <option value="all_personal">All Personal</option>
                    <option value="all_locations">Search All Locations</option>
                    {Array.from(new Set(loadedCollections.map(p => getCollectionName(p))))
                        .filter(name => name && name.toLowerCase() !== "personal")
                        .map(name => (
                            <option key={name} value={`collection:${name}`}>
                                Collection: {name}
                            </option>
                        ))}
                </select>
                {searchScope === "all_locations" && (
                  <div style={{
                    marginTop: 8, fontSize: 11, color: "#a4262c", lineHeight: 1.4,
                    padding: "6px 8px", backgroundColor: "#fef6f6", borderRadius: 4,
                    border: "1px solid #f1bbbc",
                  }}>
                    Searching entire index — may be slower on large databases.
                  </div>
                )}
            </div>



            {[
                { label: "From", value: from, setter: setFrom, icon: <MailSettings20Regular style={{ color: results ? "#0078d4" : "#c8c6c4" }} /> },
                { label: "To", value: to, setter: setTo, icon: <MailSettings20Regular style={{ color: results ? "#0078d4" : "#c8c6c4" }} /> },
                { label: "CC", value: cc, setter: setCc, icon: <MailSettings20Regular style={{ color: results ? "#0078d4" : "#c8c6c4" }} /> },
                { label: "Subject", value: subject, setter: setSubject, icon: <TextBulletList20Regular style={{ color: results ? "#0078d4" : "#c8c6c4" }} /> },
            ].map((f, idx) => (
                <div key={idx} style={{ marginBottom: 16 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                        {f.icon}
                        <span style={{ fontSize: 13, color: results ? "#605e5c" : "#c8c6c4" }}>{f.label}</span>
                    </div>
                    {f.setter && (
                        <div style={{
                          backgroundColor: results ? "#f3f2f1" : "#faf9f8",
                          borderRadius: 4, padding: "4px 8px",
                          border: `1px solid ${results ? "transparent" : "#edebe9"}`
                        }}>
                            <input
                                value={f.value}
                                onChange={e => f.setter(e.target.value)}
                                disabled={!results}
                                style={{
                                  border: "none", background: "transparent", outline: "none",
                                  width: "100%", fontSize: 12, fontFamily: "Segoe UI",
                                  cursor: results ? "text" : "not-allowed",
                                  color: results ? "#323130" : "#c8c6c4"
                                }}
                                placeholder={results ? `Filter by ${f.label.toLowerCase()}...` : (f.label === "Subject" ? "Enter subject..." : "Enter email address...")}
                            />
                        </div>
                    )}
                </div>
            ))}

            {/* Time Span Filter */}
            <div style={{ marginBottom: 20 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                    <Calendar20Regular style={{ color: "#605e5c" }} />
                    <span style={{ fontSize: 13, color: "#605e5c", fontWeight: 600 }}>Time Span</span>
                </div>
                <select
                    value={timeSpan}
                    onChange={e => setTimeSpan(e.target.value)}
                    style={{
                        width: "100%", fontSize: 13, padding: "6px 8px", borderRadius: 4,
                        border: "1px solid #edebe9", backgroundColor: "#f3f2f1",
                        color: "#323130", fontFamily: "Segoe UI", cursor: "pointer",
                        fontWeight: 600
                    }}
                >
                    <option value="all_time">All Time</option>
                    <option value="past_month">Past Month</option>
                    <option value="past_3_months">Past 3 Months</option>
                    <option value="past_6_months">Past 6 Months</option>
                    <option value="past_year">Past Year</option>
                </select>
            </div>

            {/* Attachments filter — applied on the server when Search runs */}
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
                        border: "1px solid #edebe9",
                        backgroundColor: "#f3f2f1",
                        color: "#323130",
                        fontFamily: "Segoe UI", cursor: "pointer"
                    }}
                >
                    <option value="any">Any</option>
                    <option value="with">With attachments</option>
                    <option value="without">Without attachments</option>
                </select>
            </div>

          </div>
        </div>
        )}

        {/* ── Split Container for Results & Preview ── */}
        <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "row", backgroundColor: "#ffffff" }}>
        
        {/* ── Results Pane (minWidth:0 so flex does not block horizontal scroll) ── */}
        <div style={{ flex: previewItem ? "0 0 50%" : 1, minWidth: 0, display: "flex", flexDirection: "column", backgroundColor: "#ffffff", borderRight: previewItem ? "1px solid #edebe9" : "none" }}>

          {/* Results Header */}
          <div style={{
            padding: "16px 20px", display: "flex", alignItems: "center", 
            justifyContent: "space-between", borderBottom: "1px solid #edebe9", flexShrink: 0,
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", minWidth: 0 }}>
              <span style={{ fontWeight: 600, fontSize: 16, color: "#323130" }}>Results</span>
              {results && results.results?.length > 0 && (
                <span style={{ fontSize: 13, color: "#0078d4", fontWeight: 600 }}>
                  Showing {results.results.length.toLocaleString()}
                  {results.estimatedTotalHits != null && (
                    <> of {results.estimatedTotalHits.toLocaleString()}{results.estimatedTotalHits >= 1000 ? "+" : ""}</>
                  )}
                </span>
              )}
              {results && visibleResults.length !== results.results.length && (
                <span style={{ fontSize: 12, color: "#605e5c" }}>
                  ({visibleResults.length.toLocaleString()} match filters)
                </span>
              )}
            </div>
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
              {!options.disableDelete && (
                <button type="button" onClick={handleBulkDeleteClick} style={bulkBtnDanger}>
                  Delete selected
                </button>
              )}
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
              position: "relative"
            }}
            className="search-results-scroll"
          >
            {loading && (
                <div style={{
                    position: "absolute", top: 0, left: 0, right: 0, bottom: 0,
                    backgroundColor: "rgba(255, 255, 255, 0.7)", zIndex: 10,
                    display: "flex", flexDirection: "column", alignItems: "center", paddingTop: 60
                }}>
                    <ArrowSync20Regular style={{ fontSize: 32, color: "#0078d4", animation: "spin 1s linear infinite", marginBottom: 12 }} />
                    <span style={{ fontSize: 14, fontWeight: 600, color: "#323130" }}>Searching database...</span>
                </div>
            )}
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
                      ref={input => {
                        if (input) {
                          input.indeterminate = results?.results?.length > 0 && selectedRowIds.size > 0 && selectedRowIds.size < results.results.length;
                        }
                      }}
                      onChange={handleSelectAll}
                      checked={results?.results?.length > 0 && selectedRowIds.size === results.results.length}
                    />
                  </th>
                  <th style={{ minWidth: 60, width: 60, padding: "12px 8px", textAlign: "center", fontSize: 12, fontWeight: 600, color: "#605e5c", borderBottom: "1px solid #edebe9" }}>Actions</th>
                  <th style={thStyle}>Type</th>
                  <th style={thStyle}><Attach20Regular /></th>
                  <th style={thStyle}>
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <select 
                        value={dateFilter}
                        onChange={e => setDateFilter(e.target.value)}
                        title="Filter by date range"
                        style={{ border: "none", background: "transparent", fontSize: 12, color: "#605e5c", cursor: "pointer", outline: "none", fontWeight: 600, fontFamily: "Segoe UI", appearance: "none", paddingRight: "16px", backgroundImage: "url(\"data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%23605e5c%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right center", backgroundSize: "8px" }}
                      >
                        <option value="all">Most Recent</option>
                        <option value="oldest_first">Oldest First</option>
                      </select>
                    </div>
                  </th>
                  <th style={{ ...thStyle, minWidth: 160 }}>From</th>
                  <th style={{ ...thStyle, minWidth: 160 }}>To</th>
                  <th style={{ ...thStyle, minWidth: 480 }}>Email</th>
                </tr>
              </thead>
              <tbody style={{ fontSize: 13 }}>
                {results !== null && results.results?.length > 0 && visibleResults.length === 0 && (
                  <tr>
                    <td colSpan={8} style={{ padding: "40px 20px", textAlign: "center", color: "#605e5c", fontSize: 14 }}>
                      No results match your filters.
                    </td>
                  </tr>
                )}
                {Object.entries(grouped).map(([groupLabel, items]) => (
                  <React.Fragment key={groupLabel}>
                    <tr>
                      <td colSpan={8} style={{ padding: "16px 20px 8px 20px", fontWeight: 600, color: "#8a2a21", fontSize: 12 }}>
                        {groupLabel}
                      </td>
                    </tr>
                    {items.map(r => {
                      return (
                        <tr 
                          key={r.id} 
                          style={{ 
                            borderBottom: "1px solid #f3f3f3",
                            cursor: "pointer",
                            backgroundColor: previewItem && previewItem.id === r.id ? "#f3f2f1" : "transparent"
                          }}
                          onClick={() => handlePreviewRow(r)}
                          onMouseEnter={() => prefetchPreviewBody(r.id)}
                          onDoubleClick={() => handleOpenItem(r)}
                        >
                          <td style={{ padding: "10px 20px" }}>
                            <input 
                              type="checkbox" 
                              checked={selectedRowIds.has(r.id)}
                              onChange={() => handleSelectRow(r.id)}
                            />
                          </td>
                          <td style={{ ...tdStyle, textAlign: "center", whiteSpace: "nowrap", verticalAlign: "middle" }}>
                              <div style={{ position: "relative", display: "inline-block" }}>
                                  <MoreHorizontal20Regular 
                                      style={{ color: "#605e5c", cursor: "pointer" }} 
                                      title="Actions"
                                      onClick={(e) => {
                                          e.stopPropagation();
                                          setActiveMenuId(activeMenuId === r.id ? null : r.id);
                                      }}
                                  />
                                  {activeMenuId === r.id && (
                                      <div style={{
                                          position: "absolute", top: 26, left: 0, zIndex: 200,
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
                                              onClick={(e) => { handleCopyItem(r, e); }}
                                              style={{ 
                                                  padding: "8px 12px", cursor: "pointer", fontSize: 13, 
                                                  textAlign: "left", color: "#323130" 
                                              }}
                                              onMouseOver={e => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                                              onMouseOut={e => e.currentTarget.style.backgroundColor = ""}
                                          >Copy</div>
                                          {!options.disableMoveTo && (
                                            <div 
                                                onClick={(e) => { e.stopPropagation(); handleMoveItem(r); }}
                                                style={{ padding: "8px 12px", cursor: "pointer", fontSize: 13, textAlign: "left", color: "#323130" }}
                                                onMouseOver={e => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                                                onMouseOut={e => e.currentTarget.style.backgroundColor = ""}
                                            >Transfer to..</div>
                                          )}
                                          {!options.disableDelete && (
                                            <div 
                                                onClick={(e) => { e.stopPropagation(); handleDeleteItem(r); }}
                                                style={{ padding: "8px 12px", cursor: "pointer", fontSize: 13, textAlign: "left", color: "#a4262c" }}
                                                onMouseOver={e => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                                                onMouseOut={e => e.currentTarget.style.backgroundColor = ""}
                                            >Delete</div>
                                          )}
                                      </div>
                                  )}
                              </div>
                          </td>
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
                            <div style={{ maxWidth: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "flex", alignItems: "center", gap: 6 }}>
                                <Mail20Regular style={{ fontSize: 16, color: "#605e5c", flexShrink: 0 }} />
                                <span style={{ fontSize: 12, color: "#323130" }}>{r.filePath ? r.filePath.split(/[\\/]/).pop() : ""}</span>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </React.Fragment>
                ))}
              </tbody>
            </table>

            {/* Pagination footer */}
            {results && results.results?.length > 0 && (
              <div style={{
                padding: "16px 20px",
                borderTop: "1px solid #edebe9",
                backgroundColor: "#faf9f8",
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 12,
                flexWrap: "wrap",
              }}>
                <div style={{ fontSize: 13, color: "#605e5c" }}>
                  Loaded <strong style={{ color: "#323130" }}>{results.results.length.toLocaleString()}</strong>
                  {results.estimatedTotalHits != null && (
                    <> of <strong style={{ color: "#323130" }}>{results.estimatedTotalHits.toLocaleString()}{results.estimatedTotalHits >= 1000 ? "+" : ""}</strong> matches</>
                  )}
                </div>
                {results.hasMore ? (
                  <button
                    type="button"
                    onClick={loadMoreResults}
                    disabled={isSearchBusy}
                    style={{
                      background: isSearchBusy ? "#c8c6c4" : "#0078d4",
                      border: "none",
                      borderRadius: 4,
                      padding: "8px 20px",
                      color: "#fff",
                      fontSize: 13,
                      fontWeight: 600,
                      cursor: isSearchBusy ? "not-allowed" : "pointer",
                      fontFamily: "Segoe UI",
                      minWidth: 140,
                    }}
                  >
                    {loadingMore ? "Loading…" : "Load more"}
                  </button>
                ) : (
                  <span style={{ fontSize: 12, color: "#8a8886", fontStyle: "italic" }}>
                    All matching results loaded
                  </span>
                )}
              </div>
            )}

            {/* Error / Loading / Placeholder states */}
            {loading && <div style={{ padding: 40, textAlign: "center", color: "#605e5c" }}>Searching...</div>}
            {!loading && !results && (
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100%", color: "#a19f9d" }}>
                <Search20Regular style={{ fontSize: 64, marginBottom: 16 }} />
                <span>Search for emails or locations above</span>
              </div>
            )}
            {!loading && results && results.results?.length === 0 && (
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100%", color: "#a19f9d", padding: 40 }}>
                <Dismiss20Regular style={{ fontSize: 48, marginBottom: 16, color: "#a4262c" }} />
                <span style={{ fontWeight: 600, color: "#323130" }}>No results found</span>
                <span style={{ fontSize: 13, marginTop: 4 }}>Try adjusting your filters or keywords</span>
              </div>
            )}
          </div>
        </div>

        {/* ── Preview Pane ── */}
        {previewItem && (
          <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column", backgroundColor: "#faf9f8", overflowY: "auto" }}>
             <div style={{ padding: "16px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: "1px solid #edebe9", backgroundColor: "#ffffff", position: "sticky", top: 0, zIndex: 2 }}>
               <div style={{ fontWeight: 600, fontSize: 16, color: "#323130" }}>Email Preview</div>
               <div style={{ display: "flex", gap: 16, alignItems: "center" }}>
                 <button onClick={() => handleOpenItem(previewItem)} style={bulkBtnPrimary}>Open in Outlook</button>
                 <Dismiss20Regular style={{ cursor: "pointer", color: "#605e5c", fontSize: 20 }} onClick={() => setPreviewItem(null)} title="Close Preview" />
               </div>
             </div>
             <div style={{ padding: "24px", display: "flex", flexDirection: "column", flex: 1 }}>
               <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 16 }}>
                 <h2 style={{ margin: 0, fontSize: 20, fontWeight: 600, color: "#323130", wordBreak: "break-word", lineHeight: 1.3 }}>{renderHighlightedText(previewItem.subject || "(No Subject)", keywords)}</h2>
                 {previewItem.hasAttachments && <Attach20Regular style={{ fontSize: 24, color: "#605e5c", flexShrink: 0, marginLeft: 16 }} title="Has Attachments" />}
               </div>
               <div style={{ fontSize: 13, color: "#605e5c", marginBottom: 24, paddingBottom: 16, borderBottom: "1px solid #edebe9", display: "flex", flexDirection: "column", gap: 6 }}>
                 <div><strong style={{ color: "#323130", fontWeight: 600 }}>From:</strong> {renderHighlightedText(previewItem.sender, keywords)}</div>
                 <div><strong style={{ color: "#323130", fontWeight: 600 }}>To:</strong> {renderHighlightedText(Array.isArray(previewItem.recipients) ? previewItem.recipients.join(", ") : previewItem.recipients, keywords)}</div>
                 {previewItem.cc && <div><strong style={{ color: "#323130", fontWeight: 600 }}>Cc:</strong> {renderHighlightedText(Array.isArray(previewItem.cc) ? previewItem.cc.join(", ") : previewItem.cc, keywords)}</div>}
                 <div><strong style={{ color: "#323130", fontWeight: 600 }}>Date:</strong> {formatDate(previewItem.sentAt)}</div>
               </div>
               <div style={{ fontSize: 14, color: "#323130", lineHeight: "1.6", whiteSpace: "pre-wrap", wordBreak: "break-word", flex: 1, fontFamily: "Segoe UI, sans-serif" }}>
                 {previewBodyLoadingId === previewItem.id ? (
                    <span style={{ color: "#605e5c", fontStyle: "italic" }}>Loading message…</span>
                  ) : previewBodyError && !previewItem.body ? (
                    <span style={{ color: "#a4262c" }}>{previewBodyError}</span>
                  ) : previewItem.body ? (
                    renderHighlightedText(previewItem.body, keywords)
                  ) : (
                    <span style={{ color: "#a19f9d", fontStyle: "italic" }}>No content available.</span>
                  )}
               </div>
             </div>
          </div>
        )}
        
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

        {/* Move Confirmation Overlay */}
        {moveTargetItem && (
            <div style={{
                position: "absolute", top: 0, left: 0, right: 0, bottom: 0,
                backgroundColor: "rgba(0,0,0,0.4)", display: "flex", justifyContent: "center",
                alignItems: "center", zIndex: 1000, borderRadius: 8
            }}>
                <div style={{
                    backgroundColor: "#fff", padding: 24, borderRadius: 8, width: 400,
                    boxShadow: "0 8px 32px rgba(0,0,0,0.2)", textAlign: "left"
                }}>
                    <h3 style={{ marginTop: 0, color: "#323130" }}>Transfer File</h3>
                    <p style={{ fontSize: 13, color: "#605e5c", lineHeight: "1.5", marginBottom: 12 }}>
                        Transferring: <b>{formatFileLocation(moveTargetItem.filePath)}</b><br/>
                        Enter the exact network or local destination path:
                    </p>
                    <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
                        <input 
                            ref={movePathInputRef}
                            type="text"
                            value={moveDestinationPath}
                            onChange={(e) => setMoveDestinationPath(e.target.value)}
                            placeholder="e.g. C:\Archive\Project X"
                            style={{
                                flexGrow: 1, padding: "8px", border: "1px solid #8a8886", borderRadius: 4,
                                boxSizing: "border-box", fontFamily: "Segoe UI", fontSize: 13, minWidth: 0
                            }}
                        />
                        <button
                            type="button"
                            onClick={handlePasteFolder}
                            style={{
                                padding: "8px 16px", borderRadius: 4, border: "1px solid #c8c6c4",
                                backgroundColor: "#fff", color: "#323130",
                                cursor: "pointer", fontWeight: 600, flexShrink: 0,
                                fontFamily: "Segoe UI", fontSize: 13
                            }}
                        >Paste</button>
                        <button
                            type="button"
                            onClick={handleBrowseFolder}
                            style={{
                                padding: "8px 16px", borderRadius: 4, border: "1px solid #c8c6c4",
                                backgroundColor: "#fff", color: "#323130",
                                cursor: "pointer", fontWeight: 600, flexShrink: 0,
                                fontFamily: "Segoe UI", fontSize: 13
                            }}
                        >Browse...</button>
                    </div>
                    <div style={{ display: "flex", gap: 12, justifyContent: "flex-end" }}>
                        <button 
                            onClick={() => setMoveTargetItem(null)}
                            style={{
                                padding: "8px 20px", borderRadius: 4, border: "1px solid #8a8886",
                                backgroundColor: "#fff", color: "#323130", cursor: "pointer",
                                fontWeight: 600
                            }}
                        >Cancel</button>
                        <button 
                            onClick={submitMoveItem}
                            disabled={!moveDestinationPath.trim()}
                            style={{
                                padding: "8px 20px", borderRadius: 4, border: "none",
                                backgroundColor: moveDestinationPath.trim() ? "#0078d4" : "#c8c6c4", 
                                color: "#fff", cursor: moveDestinationPath.trim() ? "pointer" : "default",
                                fontWeight: 600
                            }}
                        >Transfer</button>
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
