/**
 * MobileSearchScreen.jsx
 *
 * Mobile search screen — wraps the existing GET /api/search endpoint.
 * Layout: collapsible filter bar at top, results grouped by relative
 * date with keyword/location highlighting, inline email body preview.
 *
 * New in this version:
 *  - Results grouped by relative date (Today / Yesterday / Last week …)
 *  - Keyword + location highlighting in subject and preview body
 *  - Improved empty states and loading UX
 */

import * as React from "react";
import {
  Search24Regular,
  MailOff24Regular,
  ArrowSync20Regular,
  Attach20Regular,
  FolderOpen20Regular,
  DocumentDismiss24Regular,
} from "@fluentui/react-icons";
import { searchEmails, getSearchPreview, buildDownloadUrl } from "../services/backendApi";

/* global Office */

// ─── Date helpers ─────────────────────────────────────────────────────────────

function relativeDate(ts) {
  if (!ts) return "Unknown date";
  const date = new Date(ts);
  const now = new Date();
  const diff = now - date;
  const days = Math.floor(diff / (1000 * 60 * 60 * 24));
  if (days === 0) return "Today";
  if (days === 1) return "Yesterday";
  if (days < 7) return `${days} days ago`;
  if (days < 14) return "Last week";
  if (days < 21) return "Two weeks ago";
  if (days < 31) return "Three weeks ago";
  if (days < 60) return "Last month";
  if (days < 180) return "A few months ago";
  return date.toLocaleDateString("en-GB", { year: "numeric", month: "short" });
}

function groupByRelativeDate(results) {
  const groups = {};
  const order = [];
  results.forEach((r) => {
    const key = relativeDate(r.sentAt || r.filedAt);
    if (!groups[key]) { groups[key] = []; order.push(key); }
    groups[key].push(r);
  });
  return order.map((key) => ({ label: key, items: groups[key] }));
}

// ─── Keyword highlight helper ─────────────────────────────────────────────────

function renderHighlightedText(text, keyword) {
  if (!text || !keyword || !keyword.trim()) return text;
  try {
    const escaped = keyword.trim().replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&");
    const regex = new RegExp(`(${escaped})`, "gi");
    const parts = String(text).split(regex);
    return parts.map((part, i) =>
      regex.test(part) ? (
        <mark
          key={i}
          style={{
            background: "#fff3cd",
            color: "#333",
            borderRadius: 3,
            padding: "0 2px",
            fontWeight: 600,
          }}
        >
          {part}
        </mark>
      ) : part
    );
  } catch {
    return text;
  }
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const S = {
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    background: "#f7f8fa",
  },

  // ── Top bar ──────────────────────────────────────────────────────────────────
  topBar: {
    background: "#fff",
    borderBottom: "1px solid #e8e8e8",
    padding: "10px 12px 0",
    flexShrink: 0,
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
    background: "#fff",
    transition: "border-color .15s",
    WebkitTapHighlightColor: "transparent",
  },
  searchBtn: {
    padding: "10px 18px",
    borderRadius: 8,
    border: "none",
    background: "#0078d4",
    color: "#fff",
    fontWeight: 700,
    cursor: "pointer",
    fontSize: 14,
    flexShrink: 0,
    transition: "opacity .15s",
    WebkitTapHighlightColor: "transparent",
  },
  searchBtnDisabled: {
    background: "#a0c4e8",
    cursor: "default",
  },

  // ── Filter toggle ─────────────────────────────────────────────────────────
  filterToggle: {
    fontSize: 12,
    color: "#0078d4",
    fontWeight: 600,
    cursor: "pointer",
    padding: "4px 0 8px",
    userSelect: "none",
    display: "flex",
    alignItems: "center",
    gap: 5,
    WebkitTapHighlightColor: "transparent",
  },
  filterPanel: {
    padding: "8px 0 4px",
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
    width: 62,
    flexShrink: 0,
  },
  filterInput: {
    flex: 1,
    padding: "8px 10px",
    borderRadius: 6,
    border: "1px solid #d0d0d0",
    fontSize: 13,
    outline: "none",
    background: "#fff",
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

  // ── Count bar ─────────────────────────────────────────────────────────────
  countBar: {
    padding: "7px 12px",
    fontSize: 12,
    color: "#777",
    borderBottom: "1px solid #eee",
    background: "#fff",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexShrink: 0,
  },

  // ── Results list ─────────────────────────────────────────────────────────
  resultsList: {
    flex: 1,
    overflowY: "auto",
    padding: "10px 12px 80px",
  },

  // ── Date group header ─────────────────────────────────────────────────────
  dateGroupHeader: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    padding: "12px 2px 6px",
  },
  dateGroupLabel: {
    fontSize: 11,
    fontWeight: 700,
    color: "#aaa",
    textTransform: "uppercase",
    letterSpacing: "0.07em",
    whiteSpace: "nowrap",
    flexShrink: 0,
  },
  dateGroupLine: {
    flex: 1,
    height: 1,
    background: "#e8e8e8",
  },
  dateGroupCount: {
    fontSize: 10,
    color: "#bbb",
    fontWeight: 600,
    whiteSpace: "nowrap",
    flexShrink: 0,
  },

  // ── Result card ───────────────────────────────────────────────────────────
  card: {
    background: "#fff",
    borderRadius: 10,
    border: "1px solid #e8e8e8",
    marginBottom: 8,
    overflow: "hidden",
    boxShadow: "0 1px 3px rgba(0,0,0,.05)",
    transition: "box-shadow .15s",
  },
  cardHeader: {
    padding: "12px 14px",
    cursor: "pointer",
    WebkitTapHighlightColor: "transparent",
  },
  cardSubjectRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: 6,
    marginBottom: 4,
  },
  cardSubject: {
    fontWeight: 600,
    fontSize: 14,
    color: "#1a1a1a",
    lineHeight: 1.3,
    flex: 1,
    minWidth: 0,
  },
  attachBadge: {
    display: "inline-flex",
    alignItems: "center",
    gap: 2,
    fontSize: 10,
    fontWeight: 700,
    color: "#0078d4",
    background: "#e3f2fd",
    borderRadius: 4,
    padding: "2px 5px",
    flexShrink: 0,
    marginTop: 1,
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
    color: "#bbb",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    display: "flex",
    alignItems: "center",
    gap: 3,
  },

  // ── Preview panel ─────────────────────────────────────────────────────────
  previewPanel: {
    borderTop: "1px solid #f0f0f0",
    padding: "14px",
    background: "#f9fbfd",
  },
  previewBody: {
    fontSize: 13,
    color: "#444",
    lineHeight: 1.65,
    whiteSpace: "pre-wrap",
    maxHeight: 260,
    overflowY: "auto",
    marginBottom: 12,
    background: "#fff",
    border: "1px solid #e0e0e0",
    borderRadius: 8,
    padding: 12,
  },
  previewActions: {
    display: "flex",
    gap: 8,
    justifyContent: "flex-end",
  },
  previewBtn: {
    padding: "9px 18px",
    borderRadius: 8,
    border: "none",
    background: "#0078d4",
    color: "#fff",
    fontWeight: 600,
    fontSize: 13,
    cursor: "pointer",
    WebkitTapHighlightColor: "transparent",
  },

  // ── States ─────────────────────────────────────────────────────────────────
  stateContainer: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    padding: "40px 24px",
    textAlign: "center",
  },
  stateIcon: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 12,
    color: "#c8c6c4",
  },
  stateTitle: {
    fontSize: 15,
    fontWeight: 600,
    color: "#444",
    marginBottom: 6,
  },
  stateBody: {
    fontSize: 13,
    color: "#aaa",
    lineHeight: 1.6,
  },
  errorBox: {
    background: "#fce8e6",
    color: "#c62828",
    borderRadius: 8,
    padding: "12px 14px",
    margin: "10px 12px",
    fontSize: 13,
    lineHeight: 1.5,
  },
};

// ─── Formatters ───────────────────────────────────────────────────────────────

function formatDate(ts) {
  if (!ts) return "";
  try {
    return new Date(ts).toLocaleDateString("en-GB", {
      day: "2-digit", month: "short", year: "numeric",
    });
  } catch { return ""; }
}

function shortPath(path = "") {
  const parts = path.replace(/\\/g, "/").split("/").filter(Boolean);
  return parts.length > 3 ? `\u2026/${parts.slice(-2).join("/")}` : path;
}

// ─── Main component ────────────────────────────────────────────────────────────

export default function MobileSearchScreen() {
  const userEmail = (() => {
    try { return Office?.context?.mailbox?.userProfile?.emailAddress || ""; }
    catch { return ""; }
  })();

  // ── Filter state ──
  const [keywords, setKeywords] = React.useState("");
  const [location, setLocation] = React.useState("");
  const [from, setFrom] = React.useState("");
  const [subject, setSubject] = React.useState("");
  const [timeSpan, setTimeSpan] = React.useState("all_time");
  const [hasAttachments, setHasAttachments] = React.useState("");
  const [showFilters, setShowFilters] = React.useState(false);

  // ── Results state ──
  const [results, setResults] = React.useState(null); // null = no search yet
  const [total, setTotal] = React.useState(0);
  const [searching, setSearching] = React.useState(false);
  const [error, setError] = React.useState(null);
  // The term actually used for the search (for highlighting)
  const [activeHighlight, setActiveHighlight] = React.useState("");

  // ── Preview state ──
  const [expandedId, setExpandedId] = React.useState(null);
  const [previewCache, setPreviewCache] = React.useState({});
  const [previewLoading, setPreviewLoading] = React.useState(null);

  // ── Search ──────────────────────────────────────────────────────────────────
  const runSearch = React.useCallback(async () => {
    if (!keywords.trim() && !location.trim()) {
      setError("Please enter keywords or a location to search.");
      return;
    }
    setSearching(true);
    setError(null);
    setResults(null);
    setExpandedId(null);
    // Track highlight term for the results we're about to show
    setActiveHighlight(keywords.trim() || location.trim());
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

      const data = await searchEmails(params);
      setResults(data.results || []);
      setTotal(data.estimatedTotalHits ?? (data.results || []).length);
    } catch (e) {
      setError(e.message);
      setResults([]);
    } finally {
      setSearching(false);
    }
  }, [keywords, location, from, subject, timeSpan, hasAttachments, userEmail]);

  // ── Preview ─────────────────────────────────────────────────────────────────
  const handleCardTap = async (row) => {
    if (expandedId === row.id) { setExpandedId(null); return; }
    setExpandedId(row.id);
    if (previewCache[row.id]) return;
    setPreviewLoading(row.id);
    try {
      const params = new URLSearchParams({ id: row.id });
      if (userEmail) params.set("userEmail", userEmail);
      const data = await getSearchPreview(params);
      setPreviewCache((prev) => ({ ...prev, [row.id]: data.body || "(No body)" }));
    } catch {
      setPreviewCache((prev) => ({ ...prev, [row.id]: "(Preview unavailable)" }));
    } finally {
      setPreviewLoading(null);
    }
  };

  const handleDownload = (row) => {
    const params = new URLSearchParams({ id: row.id, filePath: row.filePath });
    if (userEmail) params.set("userEmail", userEmail);
    window.open(buildDownloadUrl(params), "_blank");
  };

  // ── Grouped results ─────────────────────────────────────────────────────────
  const groupedResults = React.useMemo(() => {
    if (!results || results.length === 0) return [];
    return groupByRelativeDate(results);
  }, [results]);

  // ── Render helpers ──────────────────────────────────────────────────────────
  const renderCard = (row) => {
    const isExpanded = expandedId === row.id;
    return (
      <div key={row.id} style={S.card}>
        {/* Card header — tap to expand */}
        <div style={S.cardHeader} onClick={() => handleCardTap(row)}>
          <div style={S.cardSubjectRow}>
            <span style={S.cardSubject}>
              {renderHighlightedText(row.subject || "(No subject)", activeHighlight)}
            </span>
            {row.hasAttachments && (
              <span style={S.attachBadge} title="Has attachments">
                <Attach20Regular style={{ fontSize: 11, width: 11, height: 11 }} />
              </span>
            )}
          </div>
          <div style={S.cardMeta}>
            <span>{row.sender || ""}</span>
            <span>{formatDate(row.sentAt)}</span>
          </div>
          <div style={S.cardPath}>
            <FolderOpen20Regular style={{ opacity: 0.4, width: 12, height: 12, flexShrink: 0 }} />
            <span>{shortPath(row.filePath)}</span>
          </div>
        </div>

        {/* Preview panel — shown when expanded */}
        {isExpanded && (
          <div style={S.previewPanel}>
            {previewLoading === row.id ? (
              <div style={{ fontSize: 12, color: "#999", marginBottom: 10, fontStyle: "italic" }}>
                Loading preview…
              </div>
            ) : (
              <div style={S.previewBody}>
                {renderHighlightedText(previewCache[row.id] || "", activeHighlight)}
              </div>
            )}
            <div style={S.previewActions}>
              {row.filePath && (
                <button style={S.previewBtn} onClick={() => handleDownload(row)}>
                  Open in Outlook
                </button>
              )}
            </div>
          </div>
        )}
      </div>
    );
  };

  // ── JSX ─────────────────────────────────────────────────────────────────────
  return (
    <div style={S.container}>

      {/* ── Top search bar ── */}
      <div style={S.topBar}>
        <div style={S.searchRow}>
          <input
            style={S.searchInput}
            placeholder="Search by location\u2026"
            value={location}
            onChange={(e) => setLocation(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && runSearch()}
          />
          <button
            style={{ ...S.searchBtn, ...(searching ? S.searchBtnDisabled : {}) }}
            onClick={runSearch}
            disabled={searching}
          >
            {searching ? "\u2026" : "Search"}
          </button>
        </div>

        {/* Filter toggle */}
        <div style={S.filterToggle} onClick={() => setShowFilters((v) => !v)}>
          <span style={{ fontSize: 13 }}>{showFilters ? "\u25b2" : "\u25bc"}</span>
          <span>{showFilters ? "Hide filters" : "More filters"}</span>
        </div>

        {/* Collapsible filter panel */}
        {showFilters && (
          <div style={S.filterPanel}>
            <div style={S.filterRow}>
              <span style={S.filterLabel}>Keywords</span>
              <input style={S.filterInput} placeholder="Keywords" value={keywords}
                onChange={(e) => setKeywords(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && runSearch()} />
            </div>
            <div style={S.filterRow}>
              <span style={S.filterLabel}>From</span>
              <input style={S.filterInput} placeholder="Sender email" value={from}
                onChange={(e) => setFrom(e.target.value)} />
            </div>
            <div style={S.filterRow}>
              <span style={S.filterLabel}>Subject</span>
              <input style={S.filterInput} placeholder="Subject contains" value={subject}
                onChange={(e) => setSubject(e.target.value)} />
            </div>
            <div style={S.filterRow}>
              <span style={S.filterLabel}>Date</span>
              <select style={S.filterSelect} value={timeSpan}
                onChange={(e) => setTimeSpan(e.target.value)}>
                <option value="all_time">All time</option>
                <option value="past_week">Past week</option>
                <option value="past_month">Past month</option>
                <option value="past_3_months">Past 3 months</option>
                <option value="past_6_months">Past 6 months</option>
                <option value="past_year">Past year</option>
              </select>
            </div>
            <div style={S.filterRow}>
              <span style={S.filterLabel}>Attach.</span>
              <select style={S.filterSelect} value={hasAttachments}
                onChange={(e) => setHasAttachments(e.target.value)}>
                <option value="">Any</option>
                <option value="true">Has attachments</option>
                <option value="false">No attachments</option>
              </select>
            </div>
          </div>
        )}
      </div>

      {/* ── Result count bar ── */}
      {results !== null && !searching && (
        <div style={S.countBar}>
          <span>
            {results.length === 0
              ? "No results found."
              : `${results.length}\u202fof\u202f${total}\u202fresult${total !== 1 ? "s" : ""}`}
          </span>
          {activeHighlight && results.length > 0 && (
            <span style={{ fontSize: 11, color: "#bbb" }}>
              Highlighting: <mark style={{ background: "#fff3cd", borderRadius: 3, padding: "0 4px", color: "#555", fontWeight: 700 }}>
                {activeHighlight}
              </mark>
            </span>
          )}
        </div>
      )}

      {/* ── Error ── */}
      {error && !searching && <div style={S.errorBox}>{error}</div>}

      {/* ── Searching spinner ── */}
      {searching && (
        <div style={S.stateContainer}>
          <div style={S.stateIcon}>
            <ArrowSync20Regular style={{ width: 44, height: 44 }} />
          </div>
          <div style={S.stateTitle}>Searching\u2026</div>
          <div style={S.stateBody}>Scanning your email archive</div>
        </div>
      )}

      {/* ── No search yet ── */}
      {!searching && results === null && !error && (
        <div style={S.stateContainer}>
          <div style={S.stateIcon}>
            <Search24Regular style={{ width: 44, height: 44 }} />
          </div>
          <div style={S.stateTitle}>Search your archive</div>
          <div style={S.stateBody}>
            Enter a location or keywords<br />and tap <strong>Search</strong>
          </div>
        </div>
      )}

      {/* ── No results ── */}
      {!searching && results !== null && results.length === 0 && (
        <div style={S.stateContainer}>
          <div style={S.stateIcon}>
            <MailOff24Regular style={{ width: 44, height: 44 }} />
          </div>
          <div style={S.stateTitle}>No results found</div>
          <div style={S.stateBody}>
            Try different keywords, broaden your<br />date range, or check the location filter.
          </div>
        </div>
      )}

      {/* ── Results grouped by date ── */}
      {!searching && results !== null && results.length > 0 && (
        <div style={S.resultsList}>
          {groupedResults.map((group) => (
            <div key={group.label}>
              {/* Date group header */}
              <div style={S.dateGroupHeader}>
                <span style={S.dateGroupLabel}>{group.label}</span>
                <div style={S.dateGroupLine} />
                <span style={S.dateGroupCount}>
                  {group.items.length}\u202f{group.items.length !== 1 ? "emails" : "email"}
                </span>
              </div>
              {/* Cards for this group */}
              {group.items.map(renderCard)}
            </div>
          ))}
        </div>
      )}

    </div>
  );
}
