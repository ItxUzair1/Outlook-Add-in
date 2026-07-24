/**
 * MobileLocationsScreen.jsx
 *
 * Displays the list of configured filing locations on mobile.
 * Reuses all existing location API calls — no new backend endpoints.
 */

import * as React from "react";
import {
  ArrowClockwise20Regular,
  Folder20Regular,
} from "@fluentui/react-icons";
import { getLocations, addLocation, deleteLocation, exploreLocation } from "../services/backendApi";

const styles = {
  container: {
    padding: "12px 16px",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    color: "#1a1a1a",
  },
  searchBar: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 8,
    border: "1.5px solid #d0d0d0",
    fontSize: 14,
    boxSizing: "border-box",
    marginBottom: 12,
    outline: "none",
  },
  card: {
    background: "#fff",
    border: "1px solid #e8e8e8",
    borderRadius: 10,
    padding: "12px 14px",
    marginBottom: 8,
    cursor: "pointer",
    boxShadow: "0 1px 3px rgba(0,0,0,.06)",
    transition: "box-shadow .15s",
  },
  cardTitle: {
    fontWeight: 600,
    fontSize: 14,
    color: "#1a1a1a",
    marginBottom: 3,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  cardPath: {
    fontSize: 11,
    color: "#777",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  badge: {
    display: "inline-block",
    fontSize: 10,
    fontWeight: 600,
    padding: "2px 6px",
    borderRadius: 4,
    background: "#e3f2fd",
    color: "#1565c0",
    marginRight: 6,
    textTransform: "uppercase",
  },
  emptyState: {
    textAlign: "center",
    color: "#999",
    marginTop: 60,
    fontSize: 13,
  },
  loadingState: {
    textAlign: "center",
    color: "#999",
    marginTop: 60,
    fontSize: 13,
  },
  errorState: {
    background: "#fce8e6",
    color: "#c62828",
    borderRadius: 8,
    padding: 12,
    marginTop: 12,
    fontSize: 13,
  },
  subfolderItem: {
    padding: "8px 12px",
    borderBottom: "1px solid #f0f0f0",
    fontSize: 13,
    color: "#444",
    display: "flex",
    alignItems: "center",
    gap: 8,
    cursor: "default",
  },
  headerRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: 10,
  },
  locationCount: {
    fontSize: 12,
    fontWeight: 600,
    color: "#aaa",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  refreshBtn: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    padding: "6px 12px",
    borderRadius: 20,
    border: "1.5px solid #d0d0d0",
    background: "#fff",
    color: "#0078d4",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    transition: "border-color .15s, opacity .15s",
    WebkitTapHighlightColor: "transparent",
  },
};

function shortPath(path = "") {
  const parts = path.replace(/\\/g, "/").split("/").filter(Boolean);
  return parts.length > 3 ? `…/${parts.slice(-2).join("/")}` : path;
}

export default function MobileLocationsScreen() {
  const [locations, setLocations] = React.useState([]);
  const [loading, setLoading] = React.useState(true);
  const [refreshing, setRefreshing] = React.useState(false);
  const [error, setError] = React.useState(null);
  const [filter, setFilter] = React.useState("");
  const [expanded, setExpanded] = React.useState(null);
  const [subfolders, setSubfolders] = React.useState({});
  const [subLoading, setSubLoading] = React.useState(null);

  const load = React.useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await getLocations();
      setLocations(Array.isArray(data) ? data : (data?.locations || []));
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  const handleRefresh = React.useCallback(async () => {
    if (refreshing) return;
    setRefreshing(true);
    setError(null);
    try {
      const data = await getLocations();
      setLocations(Array.isArray(data) ? data : (data?.locations || []));
    } catch (e) {
      setError(e.message);
    } finally {
      setRefreshing(false);
    }
  }, [refreshing]);

  React.useEffect(() => { load(); }, [load]);

  const filtered = React.useMemo(() => {
    if (!filter.trim()) return locations;
    const q = filter.toLowerCase();
    return locations.filter(
      (l) =>
        (l.description || "").toLowerCase().includes(q) ||
        (l.path || "").toLowerCase().includes(q)
    );
  }, [locations, filter]);

  const handleCardClick = async (loc) => {
    if (expanded === loc.id) {
      setExpanded(null);
      return;
    }
    setExpanded(loc.id);
    if (subfolders[loc.id]) return; // already loaded
    setSubLoading(loc.id);
    try {
      const data = await exploreLocation(loc.path);
      setSubfolders((prev) => ({ ...prev, [loc.id]: data?.subfolders || [] }));
    } catch {
      setSubfolders((prev) => ({ ...prev, [loc.id]: [] }));
    } finally {
      setSubLoading(null);
    }
  };

  if (loading) {
    return <div style={styles.loadingState}>Loading locations…</div>;
  }

  return (
    <div style={styles.container}>
      {/* Keyframe for spin animation */}
      <style>{`
        @keyframes kmSpin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes kmFadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }
      `}</style>

      {/* Header row: count + refresh button */}
      <div style={styles.headerRow}>
        <span style={styles.locationCount}>
          {filter ? `${filtered.length} of ${locations.length}` : `${locations.length}`} location{locations.length !== 1 ? "s" : ""}
        </span>
        <button
          style={{
            ...styles.refreshBtn,
            opacity: refreshing ? 0.6 : 1,
          }}
          onClick={handleRefresh}
          disabled={refreshing}
          title="Refresh locations list"
        >
          <ArrowClockwise20Regular
            style={{
              animation: refreshing ? "kmSpin 0.8s linear infinite" : "none",
              flexShrink: 0,
            }}
          />
          {refreshing ? "Refreshing…" : "Refresh"}
        </button>
      </div>
      <input
        style={styles.searchBar}
        placeholder="Search locations…"
        value={filter}
        onChange={(e) => setFilter(e.target.value)}
      />

      {error && <div style={styles.errorState}>Error: {error}</div>}

      {!error && filtered.length === 0 && (
        <div style={styles.emptyState}>
          {filter ? "No locations match your search." : "No locations configured yet."}
        </div>
      )}

      {filtered.map((loc) => (
        <div key={loc.id}>
          <div style={styles.card} onClick={() => handleCardClick(loc)}>
            <div style={styles.cardTitle}>
              {loc.isSuggested && <span style={styles.badge}>⭐ Suggested</span>}
              {loc.description || shortPath(loc.path)}
            </div>
            <div style={styles.cardPath}>{shortPath(loc.path)}</div>
          </div>

          {/* Subfolders */}
          {expanded === loc.id && (
            <div style={{
              background: "#fafafa",
              border: "1px solid #e8e8e8",
              borderTop: "none",
              borderRadius: "0 0 10px 10px",
              marginTop: -8,
              marginBottom: 8,
              overflow: "hidden",
            }}>
              {subLoading === loc.id ? (
                <div style={{ padding: "10px 14px", fontSize: 12, color: "#999" }}>
                  Loading subfolders…
                </div>
              ) : (subfolders[loc.id] || []).length === 0 ? (
                <div style={{ padding: "10px 14px", fontSize: 12, color: "#bbb" }}>
                  No subfolders found.
                </div>
              ) : (
                (subfolders[loc.id] || []).map((sf, i) => (
                  <div key={i} style={styles.subfolderItem}>
                    <Folder20Regular style={{ flexShrink: 0, color: "#0078d4", opacity: 0.7 }} />
                    <span>{sf.name || sf}</span>
                  </div>
                ))
              )}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}
