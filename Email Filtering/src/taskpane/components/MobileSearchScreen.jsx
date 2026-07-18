/**
 * MobileSearchScreen.jsx
 *
 * Mobile search screen — wraps the existing GET /api/search endpoint.
 * Layout: collapsible filter bar at top, card list of results, inline preview.
 * Same data, same API — mobile-optimised display only.
 */

import * as React from "react";
import { getResolvedBaseUrl } from "../services/backendApi";

/* global Office */

const API = () => getResolvedBaseUrl();

const styles = {
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    background: "#f7f8fa",
  },
  topBar: {
    background: "#fff",
    borderBottom: "1px solid #e8e8e8",
    padding: "10px 12px 0",
  },
  searchRow: {
    display: "flex",
    gap: 8,
    marginBottom: 8,
  },
  searchInput: {
    flex: 1,
    padding: "10px 12px",
    borderRadius: 8,
    border: "1.5px solid #d0d0d0",
    fontSize: 14,
    outline: "none",
  },
  searchBtn: {
    padding: "10px 16px",
    borderRadius: 8,
    border: "none",
    background: "#0078d4",
    color: "#fff",
    fontWeight: 600,
    cursor: "pointer",
    fontSize: 14,
  },
  filterToggle: {
    fontSize: 12,
    color: "#0078d4",
    fontWeight: 600,
    cursor: "pointer",
    padding: "4px 0 8px",
    userSelect: "none",
  },
  filterPanel: {
    padding: "8px 0",
    borderTop: "1px solid #f0f0f0",
  },
  filterRow: {
    display: "flex",
    gap: 8,
    marginBottom: 8,
    alignItems: "center",
  },
  filterLabel: {
    fontSize: 12,
    fontWeight: 600,
    color: "#555",
    width: 60,
    flexShrink: 0,
  },
  filterInput: {
    flex: 1,
    padding: "8px 10px",
    borderRadius: 6,
    border: "1px solid #d0d0d0",
    fontSize: 13,
    outline: "none",
  },
  filterSelect: {
    flex: 1,
    padding: "8px 10px",
    borderRadius: 6,
    border: "1px solid #d0d0d0",
    fontSize: 13,
    outline: "none",
    background: "#fff",
  },
  resultsList: {
    flex: 1,
    overflowY: "auto",
    padding: "10px 12px 80px",
  },
  card: {
    background: "#fff",
    borderRadius: 10,
    border: "1px solid #e8e8e8",
    marginBottom: 8,
    overflow: "hidden",
    boxShadow: "0 1px 3px rgba(0,0,0,.05)",
  },
  cardHeader: {
    padding: "12px 14px",
    cursor: "pointer",
  },
  cardSubject: {
    fontWeight: 600,
    fontSize: 14,
    marginBottom: 4,
    color: "#1a1a1a",
  },
  cardMeta: {
    fontSize: 12,
    color: "#777",
    display: "flex",
    justifyContent: "space-between",
    marginBottom: 3,
  },
  cardPath: {
    fontSize: 11,
    color: "#999",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  previewPanel: {
    borderTop: "1px solid #f0f0f0",
    padding: "12px 14px",
    background: "#fafafa",
  },
  previewBody: {
    fontSize: 13,
    color: "#333",
    lineHeight: 1.6,
    whiteSpace: "pre-wrap",
    maxHeight: 240,
    overflowY: "auto",
    marginBottom: 10,
  },
  previewActions: {
    display: "flex",
    gap: 8,
  },
  previewBtn: {
    flex: 1,
    padding: "8px 0",
    borderRadius: 7,
    border: "1.5px solid #0078d4",
    background: "#fff",
    color: "#0078d4",
    fontWeight: 600,
    fontSize: 12,
    cursor: "pointer",
  },
  emptyState: {
    textAlign: "center",
    color: "#bbb",
    marginTop: 60,
    fontSize: 13,
    lineHeight: 1.6,
  },
  loadingState: {
    textAlign: "center",
    color: "#0078d4",
    marginTop: 60,
    fontSize: 13,
  },
  errorState: {
    background: "#fce8e6",
    color: "#c62828",
    borderRadius: 8,
    padding: 12,
    margin: "10px 12px",
    fontSize: 13,
  },
  countBar: {
    padding: "6px 12px",
    fontSize: 12,
    color: "#777",
    borderBottom: "1px solid #eee",
    background: "#fff",
  },
};

function formatDate(ts) {
  if (!ts) return "";
  try {
    return new Date(ts).toLocaleDateString("en-GB", {
      day: "2-digit", month: "short", year: "numeric"
    });
  } catch { return ""; }
}

function shortPath(path = "") {
  const parts = path.replace(/\\/g, "/").split("/").filter(Boolean);
  return parts.length > 3 ? `…/${parts.slice(-2).join("/")}` : path;
}

export default function MobileSearchScreen() {
  const userEmail = (() => {
    try { return Office?.context?.mailbox?.userProfile?.emailAddress || ""; } catch { return ""; }
  })();

  const [keywords, setKeywords] = React.useState("");
  const [location, setLocation] = React.useState("");
  const [from, setFrom] = React.useState("");
  const [subject, setSubject] = React.useState("");
  const [timeSpan, setTimeSpan] = React.useState("all_time");
  const [hasAttachments, setHasAttachments] = React.useState("");
  const [showFilters, setShowFilters] = React.useState(false);

  const [results, setResults] = React.useState(null); // null = no search yet
  const [total, setTotal] = React.useState(0);
  const [searching, setSearching] = React.useState(false);
  const [error, setError] = React.useState(null);

  const [expandedId, setExpandedId] = React.useState(null);
  const [previewCache, setPreviewCache] = React.useState({});
  const [previewLoading, setPreviewLoading] = React.useState(null);

  const runSearch = React.useCallback(async () => {
    if (!keywords.trim() && !location.trim()) {
      setError("Please enter keywords or a location to search.");
      return;
    }
    setSearching(true);
    setError(null);
    setResults(null);
    setExpandedId(null);
    try {
      const params = new URLSearchParams({ _t: Date.now() });
      if (keywords.trim()) params.set("keywords", keywords.trim());
      if (location.trim()) params.set("location", location.trim());
      if (from.trim()) params.set("from", from.trim());
      if (subject.trim()) params.set("subject", subject.trim());
      if (timeSpan && timeSpan !== "all_time") params.set("timeSpan", timeSpan);
      if (hasAttachments) params.set("hasAttachments", hasAttachments);
      if (userEmail) params.set("userEmail", userEmail);
      params.set("limit", "50");

      const resp = await fetch(`${API()}/api/search?${params.toString()}`);
      if (!resp.ok) {
        const data = await resp.json().catch(() => ({}));
        throw new Error(data.error || `Search failed (${resp.status})`);
      }
      const data = await resp.json();
      setResults(data.results || []);
      setTotal(data.estimatedTotalHits ?? (data.results || []).length);
    } catch (e) {
      setError(e.message);
      setResults([]);
    } finally {
      setSearching(false);
    }
  }, [keywords, location, from, subject, timeSpan, hasAttachments, userEmail]);

  const handleCardTap = async (row) => {
    if (expandedId === row.id) {
      setExpandedId(null);
      return;
    }
    setExpandedId(row.id);
    if (previewCache[row.id]) return;
    setPreviewLoading(row.id);
    try {
      const params = new URLSearchParams({ id: row.id });
      if (userEmail) params.set("userEmail", userEmail);
      const resp = await fetch(`${API()}/api/search/preview?${params.toString()}`);
      if (!resp.ok) throw new Error("Preview failed");
      const data = await resp.json();
      setPreviewCache((prev) => ({ ...prev, [row.id]: data.body || "(No body)" }));
    } catch {
      setPreviewCache((prev) => ({ ...prev, [row.id]: "(Preview unavailable)" }));
    } finally {
      setPreviewLoading(null);
    }
  };

  const handleDownload = (row) => {
    const params = new URLSearchParams({ filePath: row.filePath });
    if (userEmail) params.set("userEmail", userEmail);
    window.open(`${API()}/api/search/download?${params.toString()}`, "_blank");
  };

  return (
    <div style={styles.container}>
      {/* Top search bar */}
      <div style={styles.topBar}>
        <div style={styles.searchRow}>
          <input
            style={styles.searchInput}
            placeholder="Keywords…"
            value={keywords}
            onChange={(e) => setKeywords(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && runSearch()}
          />
          <button style={styles.searchBtn} onClick={runSearch}>Search</button>
        </div>

        {/* Filter toggle */}
        <div style={styles.filterToggle} onClick={() => setShowFilters((v) => !v)}>
          {showFilters ? "▲ Hide filters" : "▼ More filters"}
        </div>

        {/* Collapsible filter panel */}
        {showFilters && (
          <div style={styles.filterPanel}>
            <div style={styles.filterRow}>
              <span style={styles.filterLabel}>Location</span>
              <input
                style={styles.filterInput}
                placeholder="Project / folder"
                value={location}
                onChange={(e) => setLocation(e.target.value)}
              />
            </div>
            <div style={styles.filterRow}>
              <span style={styles.filterLabel}>From</span>
              <input
                style={styles.filterInput}
                placeholder="Sender"
                value={from}
                onChange={(e) => setFrom(e.target.value)}
              />
            </div>
            <div style={styles.filterRow}>
              <span style={styles.filterLabel}>Subject</span>
              <input
                style={styles.filterInput}
                placeholder="Subject contains"
                value={subject}
                onChange={(e) => setSubject(e.target.value)}
              />
            </div>
            <div style={styles.filterRow}>
              <span style={styles.filterLabel}>Date</span>
              <select
                style={styles.filterSelect}
                value={timeSpan}
                onChange={(e) => setTimeSpan(e.target.value)}
              >
                <option value="all_time">All time</option>
                <option value="past_week">Past week</option>
                <option value="past_month">Past month</option>
                <option value="past_3_months">Past 3 months</option>
                <option value="past_6_months">Past 6 months</option>
                <option value="past_year">Past year</option>
              </select>
            </div>
            <div style={styles.filterRow}>
              <span style={styles.filterLabel}>Attach.</span>
              <select
                style={styles.filterSelect}
                value={hasAttachments}
                onChange={(e) => setHasAttachments(e.target.value)}
              >
                <option value="">Any</option>
                <option value="true">Has attachments</option>
                <option value="false">No attachments</option>
              </select>
            </div>
          </div>
        )}
      </div>

      {/* Result count */}
      {results !== null && !searching && (
        <div style={styles.countBar}>
          {results.length === 0
            ? "No results found."
            : `Showing ${results.length} of ${total} result${total !== 1 ? "s" : ""}`}
        </div>
      )}

      {/* Body */}
      {searching && <div style={styles.loadingState}>Searching…</div>}
      {error && !searching && <div style={styles.errorState}>{error}</div>}

      {!searching && results === null && !error && (
        <div style={styles.emptyState}>
          Enter keywords or a location<br />and tap Search.
        </div>
      )}

      {!searching && results !== null && (
        <div style={styles.resultsList}>
          {results.map((row) => {
            const isExpanded = expandedId === row.id;
            return (
              <div key={row.id} style={styles.card}>
                {/* Card header — tap to expand */}
                <div style={styles.cardHeader} onClick={() => handleCardTap(row)}>
                  <div style={styles.cardSubject}>
                    {row.hasAttachments ? "📎 " : ""}{row.subject || "(No subject)"}
                  </div>
                  <div style={styles.cardMeta}>
                    <span>{row.sender || ""}</span>
                    <span>{formatDate(row.sentAt)}</span>
                  </div>
                  <div style={styles.cardPath}>{shortPath(row.filePath)}</div>
                </div>

                {/* Preview panel */}
                {isExpanded && (
                  <div style={styles.previewPanel}>
                    {previewLoading === row.id ? (
                      <div style={{ fontSize: 12, color: "#999", marginBottom: 10 }}>Loading preview…</div>
                    ) : (
                      <div style={styles.previewBody}>
                        {previewCache[row.id] || ""}
                      </div>
                    )}
                    <div style={styles.previewActions}>
                      {row.filePath && (
                        <button
                          style={styles.previewBtn}
                          onClick={() => handleDownload(row)}
                        >
                          Open in Outlook
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
