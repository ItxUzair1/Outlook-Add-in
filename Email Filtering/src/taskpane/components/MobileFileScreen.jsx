/**
 * MobileFileScreen.jsx
 *
 * Mobile "File Email" screen — lets the user pick a filing location and file
 * the currently selected email. Uses the same backend API and email-payload
 * builder as the desktop App.jsx — just a mobile-optimised layout.
 *
 * Flow:
 *  1. On mount, read email metadata from Office context
 *  2. Load locations list from agent (same GET /api/locations)
 *  3. User filters/selects a location
 *  4. User taps "File" → POST /api/file/email (same as desktop)
 */

import * as React from "react";
import {
  getLocations,
  fileEmail,
} from "../services/backendApi";
import { buildCurrentEmailPayload } from "../services/mailboxService";

/* global Office */

const styles = {
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    color: "#1a1a1a",
    background: "#f7f8fa",
  },
  emailCard: {
    background: "#fff",
    borderBottom: "1px solid #e8e8e8",
    padding: "12px 16px",
  },
  emailSubject: {
    fontWeight: 600,
    fontSize: 14,
    color: "#1a1a1a",
    marginBottom: 3,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  emailMeta: {
    fontSize: 12,
    color: "#777",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  searchBar: {
    margin: "10px 16px",
    padding: "10px 12px",
    borderRadius: 8,
    border: "1.5px solid #d0d0d0",
    fontSize: 14,
    outline: "none",
    background: "#fff",
  },
  sectionLabel: {
    padding: "0 16px 6px",
    fontSize: 11,
    fontWeight: 600,
    color: "#999",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  list: {
    flex: 1,
    overflowY: "auto",
    padding: "0 0 80px",
  },
  locationRow: {
    display: "flex",
    alignItems: "center",
    padding: "11px 16px",
    borderBottom: "1px solid #f0f0f0",
    cursor: "pointer",
    background: "#fff",
    transition: "background .12s",
  },
  locationRowSelected: {
    background: "#e3f2fd",
  },
  locationTitle: {
    fontWeight: 500,
    fontSize: 14,
    marginBottom: 2,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  locationPath: {
    fontSize: 11,
    color: "#999",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  dot: {
    width: 8,
    height: 8,
    borderRadius: "50%",
    background: "#0078d4",
    marginRight: 10,
    flexShrink: 0,
    opacity: 0,
    transition: "opacity .15s",
  },
  dotVisible: {
    opacity: 1,
  },
  footer: {
    position: "fixed",
    bottom: 60, // above bottom nav
    left: 0,
    right: 0,
    padding: "10px 16px",
    background: "#fff",
    borderTop: "1px solid #e8e8e8",
    zIndex: 10,
  },
  fileBtn: {
    width: "100%",
    padding: "13px 0",
    borderRadius: 10,
    border: "none",
    background: "#0078d4",
    color: "#fff",
    fontWeight: 700,
    fontSize: 15,
    cursor: "pointer",
    transition: "opacity .15s",
  },
  fileBtnDisabled: {
    background: "#b0b0b0",
    cursor: "default",
  },
  statusBox: {
    margin: "0 16px 10px",
    padding: "10px 12px",
    borderRadius: 8,
    fontSize: 13,
    fontWeight: 500,
  },
  statusOk: { background: "#e6f4ea", color: "#2e7d32" },
  statusError: { background: "#fce8e6", color: "#c62828" },
  statusLoading: { background: "#e3f2fd", color: "#1565c0" },
};

function shortPath(path = "") {
  const parts = path.replace(/\\/g, "/").split("/").filter(Boolean);
  return parts.length > 3 ? `…/${parts.slice(-2).join("/")}` : path;
}

export default function MobileFileScreen() {
  const [emailInfo, setEmailInfo] = React.useState(null);
  const [locations, setLocations] = React.useState([]);
  const [filter, setFilter] = React.useState("");
  const [selectedId, setSelectedId] = React.useState(null);
  const [status, setStatus] = React.useState(null); // null | {type, msg}
  const [filing, setFiling] = React.useState(false);

  // Read current email subject/sender from Office context
  React.useEffect(() => {
    try {
      const item = Office?.context?.mailbox?.item;
      if (item) {
        setEmailInfo({
          subject: item.subject || "(No subject)",
          sender: item.from?.displayName || item.from?.emailAddress || "",
        });
      }
    } catch { /* Office not ready yet */ }
  }, []);

  // Load locations
  React.useEffect(() => {
    getLocations()
      .then((data) => setLocations(Array.isArray(data) ? data : (data?.locations || [])))
      .catch((e) => setStatus({ type: "error", msg: `Failed to load locations: ${e.message}` }));
  }, []);

  const filtered = React.useMemo(() => {
    if (!filter.trim()) return locations;
    const q = filter.toLowerCase();
    return locations.filter(
      (l) =>
        (l.description || "").toLowerCase().includes(q) ||
        (l.path || "").toLowerCase().includes(q)
    );
  }, [locations, filter]);

  const selected = locations.find((l) => l.id === selectedId);

  const handleFile = async () => {
    if (!selected || filing) return;
    setFiling(true);
    setStatus({ type: "loading", msg: "Building email payload…" });
    try {
      const payload = await buildCurrentEmailPayload();
      if (!payload) throw new Error("Could not read email data from Outlook.");

      setStatus({ type: "loading", msg: "Filing email…" });
      await fileEmail({ ...payload, targetPaths: [selected.path] });

      setStatus({ type: "ok", msg: `✅ Filed to: ${selected.description || shortPath(selected.path)}` });
      setSelectedId(null);
      setFilter("");
    } catch (e) {
      setStatus({ type: "error", msg: `❌ Filing failed: ${e.message}` });
    } finally {
      setFiling(false);
    }
  };

  const isDisabled = !selectedId || filing;

  return (
    <div style={styles.container}>
      {/* Current email info */}
      {emailInfo && (
        <div style={styles.emailCard}>
          <div style={styles.emailSubject}>{emailInfo.subject}</div>
          <div style={styles.emailMeta}>From: {emailInfo.sender}</div>
        </div>
      )}

      {/* Status message */}
      {status && (
        <div style={{
          ...styles.statusBox,
          ...(status.type === "ok" ? styles.statusOk
            : status.type === "error" ? styles.statusError
            : styles.statusLoading),
        }}>
          {status.msg}
        </div>
      )}

      {/* Selected location banner */}
      {selected && (
        <div style={{ padding: "4px 16px 0", fontSize: 12, color: "#0078d4", fontWeight: 600 }}>
          Selected: {selected.description || shortPath(selected.path)}
        </div>
      )}

      {/* Filter */}
      <input
        style={styles.searchBar}
        placeholder="Filter locations…"
        value={filter}
        onChange={(e) => setFilter(e.target.value)}
      />

      <div style={styles.sectionLabel}>Choose a location</div>

      {/* Location list */}
      <div style={styles.list}>
        {filtered.map((loc) => {
          const isSelected = loc.id === selectedId;
          return (
            <div
              key={loc.id}
              style={{
                ...styles.locationRow,
                ...(isSelected ? styles.locationRowSelected : {}),
              }}
              onClick={() => setSelectedId(isSelected ? null : loc.id)}
            >
              <div style={{ ...styles.dot, ...(isSelected ? styles.dotVisible : {}) }} />
              <div style={{ minWidth: 0 }}>
                <div style={styles.locationTitle}>
                  {loc.isSuggested ? "⭐ " : ""}{loc.description || shortPath(loc.path)}
                </div>
                <div style={styles.locationPath}>{shortPath(loc.path)}</div>
              </div>
            </div>
          );
        })}
        {filtered.length === 0 && (
          <div style={{ textAlign: "center", color: "#bbb", marginTop: 40, fontSize: 13 }}>
            No locations found.
          </div>
        )}
      </div>

      {/* File button */}
      <div style={styles.footer}>
        <button
          style={{
            ...styles.fileBtn,
            ...(isDisabled ? styles.fileBtnDisabled : {}),
          }}
          disabled={isDisabled}
          onClick={handleFile}
        >
          {filing ? "Filing…" : selectedId ? `File Email` : "Select a location"}
        </button>
      </div>
    </div>
  );
}
