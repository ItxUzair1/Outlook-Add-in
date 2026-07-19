/**
 * MobileFileScreen.jsx
 *
 * Mobile "File Email" screen — lets the user pick one or more filing
 * locations and file the currently selected email.
 *
 * Features (parity with desktop):
 *  - Suggested locations shown at top (starred ⭐)
 *  - Multi-select mode: file to multiple locations at once
 *  - After-filing action: none | add_date | archive | delete | move_filed_items
 *  - Apply "Filed by Koyomail" category after successful filing
 *  - All options are persisted in / read from koyomail_options localStorage
 */

import * as React from "react";
import {
  getLocations,
  fileEmail,
  toggleSuggestion,
} from "../services/backendApi";
import { buildCurrentEmailPayload, addCategoryToCurrentEmail, ensureMasterCategory } from "../services/mailboxService";
import { getGraphToken } from "../utils/authManager";
import { Star16Regular, Star16Filled } from "@fluentui/react-icons";

/* global Office */

// ─── Colours ──────────────────────────────────────────────────────────────────
const BRAND    = "#0078d4";
const BRAND_LT = "#e3f2fd";
const SUCCESS_BG  = "#e6f4ea";
const SUCCESS_FG  = "#2e7d32";
const ERROR_BG    = "#fce8e6";
const ERROR_FG    = "#c62828";
const LOADING_BG  = "#e3f2fd";
const LOADING_FG  = "#1565c0";

// ─── Styles ───────────────────────────────────────────────────────────────────
const S = {
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    color: "#1a1a1a",
    background: "#f7f8fa",
  },

  // Email card at top
  emailCard: {
    background: "#fff",
    borderBottom: "1px solid #e8e8e8",
    padding: "10px 16px",
  },
  emailSubject: {
    fontWeight: 600,
    fontSize: 14,
    color: "#1a1a1a",
    marginBottom: 2,
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

  // Toolbar row (search + multi-select toggle)
  toolbarRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    padding: "8px 16px 0",
  },
  searchBar: {
    flex: 1,
    padding: "9px 12px",
    borderRadius: 8,
    border: "1.5px solid #d0d0d0",
    fontSize: 14,
    outline: "none",
    background: "#fff",
  },
  multiToggle: {
    padding: "9px 12px",
    borderRadius: 8,
    border: "1.5px solid #0078d4",
    background: "#fff",
    color: "#0078d4",
    fontWeight: 600,
    fontSize: 12,
    cursor: "pointer",
    whiteSpace: "nowrap",
    flexShrink: 0,
  },
  multiToggleActive: {
    background: "#0078d4",
    color: "#fff",
  },

  // Status message
  statusBox: {
    margin: "8px 16px 0",
    padding: "10px 12px",
    borderRadius: 8,
    fontSize: 13,
    fontWeight: 500,
  },

  // Section header (Suggested / All Locations)
  sectionLabel: {
    padding: "10px 16px 4px",
    fontSize: 11,
    fontWeight: 700,
    color: "#999",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },

  // Scrollable list
  list: {
    flex: 1,
    overflowY: "auto",
    paddingBottom: 200, // room for footer
  },

  // Location rows
  locationRow: {
    display: "flex",
    alignItems: "center",
    padding: "11px 16px",
    borderBottom: "1px solid #f0f0f0",
    cursor: "pointer",
    background: "#fff",
    transition: "background .1s",
    WebkitTapHighlightColor: "transparent",
  },
  locationRowSelected: {
    background: "#f3f2f1",
  },
  checkBox: {
    width: 18,
    height: 18,
    borderRadius: 4,
    border: "2px solid #c0c0c0",
    marginRight: 12,
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "all .12s",
  },
  checkBoxChecked: {
    background: BRAND,
    border: `2px solid ${BRAND}`,
  },
  checkMark: {
    color: "#fff",
    fontSize: 11,
    lineHeight: 1,
    fontWeight: 700,
  },
  dot: {
    width: 8,
    height: 8,
    borderRadius: "50%",
    background: BRAND,
    marginRight: 12,
    flexShrink: 0,
    opacity: 0,
    transition: "opacity .12s",
  },
  dotVisible: { opacity: 1 },
  locationTitle: {
    fontWeight: 600,
    fontSize: 14,
    color: "#323130",
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

  // ── Post-filing options accordion ──────────────────────────────────────────
  optionsAccordion: {
    margin: "8px 16px 0",
    border: "1.5px solid #e0e0e0",
    borderRadius: 10,
    background: "#fff",
    overflow: "hidden",
  },
  accordionHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 14px",
    cursor: "pointer",
    userSelect: "none",
    WebkitTapHighlightColor: "transparent",
  },
  accordionTitle: {
    fontWeight: 600,
    fontSize: 13,
    color: "#333",
  },
  accordionChevron: {
    fontSize: 12,
    color: "#888",
    transition: "transform .2s",
  },
  accordionBody: {
    borderTop: "1px solid #f0f0f0",
    padding: "12px 14px",
    display: "flex",
    flexDirection: "column",
    gap: 14,
  },

  // Option row inside accordion
  optionRow: {
    display: "flex",
    flexDirection: "column",
    gap: 4,
  },
  optionLabel: {
    fontSize: 12,
    fontWeight: 600,
    color: "#555",
  },
  select: {
    padding: "8px 10px",
    borderRadius: 8,
    border: "1.5px solid #d0d0d0",
    fontSize: 13,
    background: "#fff",
    color: "#1a1a1a",
    outline: "none",
  },

  // Toggle switch (category)
  toggleRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 8,
  },
  toggleLabel: {
    fontSize: 13,
    color: "#333",
    flex: 1,
  },
  toggleTrack: (on) => ({
    width: 40,
    height: 22,
    borderRadius: 11,
    background: on ? BRAND : "#ccc",
    position: "relative",
    flexShrink: 0,
    cursor: "pointer",
    transition: "background .2s",
  }),
  toggleThumb: (on) => ({
    position: "absolute",
    top: 3,
    left: on ? 20 : 3,
    width: 16,
    height: 16,
    borderRadius: "50%",
    background: "#fff",
    transition: "left .2s",
    boxShadow: "0 1px 3px rgba(0,0,0,.25)",
  }),

  // Footer
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
    background: BRAND,
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
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
function shortPath(path = "") {
  const parts = path.replace(/\\/g, "/").split("/").filter(Boolean);
  return parts.length > 3 ? `…/${parts.slice(-2).join("/")}` : path;
}

function loadOpts() {
  try {
    return JSON.parse(localStorage.getItem("koyomail_options") || "{}");
  } catch { return {}; }
}

// ─── Toggle switch component ──────────────────────────────────────────────────
function Toggle({ on, onChange }) {
  return (
    <div
      style={S.toggleTrack(on)}
      onClick={() => onChange(!on)}
      role="switch"
      aria-checked={on}
    >
      <div style={S.toggleThumb(on)} />
    </div>
  );
}

// ─── Main component ───────────────────────────────────────────────────────────
export default function MobileFileScreen() {
  // ── Email info ──
  const [emailInfo, setEmailInfo] = React.useState(null);

  // ── Locations ──
  const [locations, setLocations]   = React.useState([]);
  const [filter, setFilter]         = React.useState("");
  const [multiMode, setMultiMode]   = React.useState(false);
  const [selectedIds, setSelectedIds] = React.useState([]); // always an array

  // ── Status ──
  const [status, setStatus] = React.useState(null); // null | { type, msg }
  const [filing, setFiling] = React.useState(false);

  // ── Post-filing options ──
  const [optionsOpen, setOptionsOpen] = React.useState(false);
  const opts = loadOpts();
  const [afterFiling, setAfterFiling] = React.useState(() => {
    const raw = opts.afterFilingAction;
    return raw === "move_deleted" ? "delete" : (raw || "none");
  });
  const [addCategory, setAddCategory] = React.useState(
    opts.addFiledCategory !== false
  );
  const categoryName = opts.filedCategoryName || "Filed by Koyomail";

  // Read current email subject / sender
  React.useEffect(() => {
    try {
      const item = Office?.context?.mailbox?.item;
      if (item) {
        setEmailInfo({
          subject: item.subject || "(No subject)",
          sender: item.from?.displayName || item.from?.emailAddress || "",
        });
      }
    } catch { /* Office not ready */ }
  }, []);

  const fetchLocations = React.useCallback(() => {
    getLocations()
      .then((data) => setLocations(Array.isArray(data) ? data : (data?.locations || [])))
      .catch((e) => setStatus({ type: "error", msg: `Failed to load locations: ${e.message}` }));
  }, []);

  // Load locations from agent
  React.useEffect(() => {
    fetchLocations();
  }, [fetchLocations]);

  // ── Derived lists ──────────────────────────────────────────────────────────
  const filtered = React.useMemo(() => {
    const q = filter.toLowerCase().trim();
    const base = q
      ? locations.filter(
          (l) =>
            (l.description || "").toLowerCase().includes(q) ||
            (l.path || "").toLowerCase().includes(q)
        )
      : locations;
    return base;
  }, [locations, filter]);

  const suggested = filtered.filter((l) => l.isSuggested);
  const others    = filtered.filter((l) => !l.isSuggested);

  // ── Selection helpers ──────────────────────────────────────────────────────
  const toggleSelect = (id) => {
    if (multiMode) {
      setSelectedIds((prev) =>
        prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
      );
    } else {
      setSelectedIds((prev) => (prev[0] === id ? [] : [id]));
    }
  };

  const toggleMultiMode = () => {
    setMultiMode((m) => !m);
    setSelectedIds([]);
  };

  const isSelected = (id) => selectedIds.includes(id);
  const selectedCount = selectedIds.length;

  const handleToggleFavourite = async (e, id) => {
    e.stopPropagation();
    try {
      await toggleSuggestion(id);
      fetchLocations();
    } catch (err) {
      console.warn("Failed to toggle suggestion", err);
    }
  };

  // ── File action ────────────────────────────────────────────────────────────
  const handleFile = async () => {
    if (selectedCount === 0 || filing) return;
    setFiling(true);
    setStatus({ type: "loading", msg: "Building email payload…" });

    try {
      const payload = await buildCurrentEmailPayload();
      if (!payload) throw new Error("Could not read email data from Outlook.");

      const targetPaths = selectedIds.map(
        (id) => locations.find((l) => l.id === id)?.path
      ).filter(Boolean);

      if (targetPaths.length === 0) throw new Error("No valid locations selected.");

      // ── Acquire Graph token for post-filing actions ────────────────────────
      // On mobile the taskpane is an iframe so Office SSO is skipped.
      // getGraphToken() tries: Tier 0 cache → Tier 2 NAA (Outlook Mobile broker).
      // Without a token the backend silently skips category, add_date, archive etc.
      let graphAccessToken = null;
      let ssoToken = payload.ssoToken || null;
      const needsGraphActions = addCategory ||
        (afterFiling && afterFiling !== "none");

      if (needsGraphActions) {
        setStatus({ type: "loading", msg: "Authenticating with Microsoft…" });
        try {
          const tokenResult = await getGraphToken({
            msalInstance: null,   // NAA/cache paths don't need msalInstance
            interactive: false,
            loginHint: Office?.context?.mailbox?.userProfile?.emailAddress,
          });
          if (tokenResult?.token) {
            if (tokenResult.tier === "sso") {
              // SSO identity token → backend handles OBO exchange
              ssoToken = tokenResult.token;
            } else {
              // NAA/cache → direct Graph access token
              graphAccessToken = tokenResult.token;
            }
            console.log(`[MobileFileScreen] Graph token acquired (tier: ${tokenResult.tier}).`);
          }
        } catch (tokenErr) {
          // Non-fatal: file the email without post-filing Graph actions.
          // Client-side Office.js category call below still runs.
          console.warn("[MobileFileScreen] Could not acquire Graph token:", tokenErr.message);
        }
      }

      setStatus({ type: "loading", msg: "Filing email…" });

      await fileEmail({
        ...payload,
        ssoToken,
        graphAccessToken,
        targetPaths,
        afterFiling,
        addFiledCategory: addCategory,
        filedCategoryName: categoryName,
        // Pass preferences so backend applies them via Graph
        duplicateStrategy: opts.duplicateStrategy || "rename",
        useUtcTime: opts.useUtcTime || false,
        applyReadOnly: opts.applyReadOnly || false,
        emailFont: opts.emailFont || "Times New Roman",
        fontSize: opts.fontSize || "10",
      });

      // Apply "Filed by Koyomail" category via Office.js (client-side)
      // This mirrors App.jsx line ~1915 — works even without Graph token
      if (addCategory) {
        try {
          await ensureMasterCategory(categoryName);
          await addCategoryToCurrentEmail(categoryName);
        } catch (catErr) {
          console.warn("[MobileFileScreen] Client-side category failed:", catErr.message);
          // Not a fatal error — backend may have applied it via Graph
        }
      }

      const locationNames = targetPaths.map(shortPath).join(", ");
      setStatus({
        type: "ok",
        msg: `✅ Filed to: ${locationNames}`,
      });
      setSelectedIds([]);
      setFilter("");
    } catch (e) {
      setStatus({ type: "error", msg: `❌ Filing failed: ${e.message}` });
    } finally {
      setFiling(false);
    }
  };

  const isDisabled = selectedCount === 0 || filing;

  // ── Render helpers ─────────────────────────────────────────────────────────
  const renderRow = (loc) => {
    const sel = isSelected(loc.id);
    return (
      <div
        key={loc.id}
        style={{ ...S.locationRow, ...(sel ? S.locationRowSelected : {}) }}
        onClick={() => toggleSelect(loc.id)}
      >
        {/* Checkbox in multi-mode, dot in single-mode */}
        {multiMode ? (
          <div style={{ ...S.checkBox, ...(sel ? S.checkBoxChecked : {}) }}>
            {sel && <span style={S.checkMark}>✓</span>}
          </div>
        ) : (
          <div style={{ ...S.dot, ...(sel ? S.dotVisible : {}) }} />
        )}
        <div style={{ minWidth: 0, flex: 1 }}>
          <div style={S.locationTitle}>
            {loc.description || shortPath(loc.path)}
          </div>
          <div style={S.locationPath}>
            {loc.collection && loc.collection !== "Private" && (
              <span style={{ fontWeight: 600, marginRight: 4, color: "#323130" }}>[{loc.collection}]</span>
            )}
            {shortPath(loc.path)}
          </div>
        </div>
        <div
          style={{ padding: 8, flexShrink: 0 }}
          onClick={(e) => handleToggleFavourite(e, loc.id)}
        >
          {loc.isSuggested ? (
            <Star16Filled style={{ color: "#ffb900", fontSize: 18 }} />
          ) : (
            <Star16Regular style={{ color: "#c8c6c4", fontSize: 18 }} />
          )}
        </div>
      </div>
    );
  };

  // ── After-filing label (for accordion header summary) ─────────────────────
  const afterFilingLabels = {
    none: "None",
    add_date: "Add filed date to subject",
    archive: "Move to Archive",
    delete: "Move to Deleted Items",
    move_filed_items: "Move to Filed Items folder",
  };

  const optionsSummary = [
    afterFiling !== "none" ? afterFilingLabels[afterFiling] : null,
    addCategory ? `Apply "${categoryName}"` : null,
  ].filter(Boolean).join(" · ") || "None";

  // ── File button label ─────────────────────────────────────────────────────
  const btnLabel = filing
    ? "Filing…"
    : selectedCount > 1
    ? `File to ${selectedCount} locations`
    : selectedCount === 1
    ? "File Email"
    : "Select a location";

  return (
    <div style={S.container}>

      {/* ── Email card ── */}
      {emailInfo && (
        <div style={S.emailCard}>
          <div style={S.emailSubject}>{emailInfo.subject}</div>
          <div style={S.emailMeta}>From: {emailInfo.sender}</div>
        </div>
      )}

      {/* ── Toolbar: search + multi-select toggle ── */}
      <div style={S.toolbarRow}>
        <input
          style={S.searchBar}
          placeholder="Filter locations…"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
        />
        <button
          style={{
            ...S.multiToggle,
            ...(multiMode ? S.multiToggleActive : {}),
          }}
          onClick={toggleMultiMode}
        >
          {multiMode ? "✓ Multi" : "Multi"}
        </button>
      </div>

      {/* ── Status ── */}
      {status && (
        <div style={{
          ...S.statusBox,
          background: status.type === "ok" ? SUCCESS_BG
            : status.type === "error" ? ERROR_BG
            : LOADING_BG,
          color: status.type === "ok" ? SUCCESS_FG
            : status.type === "error" ? ERROR_FG
            : LOADING_FG,
        }}>
          {status.msg}
        </div>
      )}

      {/* ── Post-filing options accordion ── */}
      <div style={S.optionsAccordion}>
        <div style={S.accordionHeader} onClick={() => setOptionsOpen((o) => !o)}>
          <span style={S.accordionTitle}>Post-filing options</span>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
            {!optionsOpen && (
              <span style={{ fontSize: 11, color: "#888", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                {optionsSummary}
              </span>
            )}
            <span style={{ ...S.accordionChevron, transform: optionsOpen ? "rotate(180deg)" : "rotate(0deg)" }}>
              ▾
            </span>
          </span>
        </div>

        {optionsOpen && (
          <div style={S.accordionBody}>

            {/* After-filing action */}
            <div style={S.optionRow}>
              <label style={S.optionLabel}>After filing</label>
              <select
                style={S.select}
                value={afterFiling}
                onChange={(e) => setAfterFiling(e.target.value)}
              >
                <option value="none">Keep in box</option>
                <option value="add_date">Add filed date &amp; time to subject</option>
                <option value="archive">Move to Archive</option>
                <option value="delete">Move to Deleted Items</option>
                <option value="move_filed_items">Move to "Filed Items" folder</option>
              </select>
            </div>

            {/* Apply category toggle */}
            <div style={S.toggleRow}>
              <span style={S.toggleLabel}>
                Apply <strong>"{categoryName}"</strong> category
              </span>
              <Toggle on={addCategory} onChange={setAddCategory} />
            </div>

          </div>
        )}
      </div>

      {/* ── Location list ── */}
      <div style={S.list}>

        {/* Suggested section */}
        {suggested.length > 0 && (
          <>
            <div style={S.sectionLabel}>⭐ Suggested</div>
            {suggested.map(renderRow)}
          </>
        )}

        {/* All locations section */}
        {others.length > 0 && (
          <>
            <div style={S.sectionLabel}>
              {suggested.length > 0 ? "All locations" : "Choose a location"}
            </div>
            {others.map(renderRow)}
          </>
        )}

        {filtered.length === 0 && (
          <div style={{ textAlign: "center", color: "#bbb", marginTop: 48, fontSize: 13 }}>
            No locations found.
          </div>
        )}
      </div>

      {/* ── File button ── */}
      <div style={S.footer}>
        {selectedCount > 0 && !multiMode && (
          <div style={{ fontSize: 12, color: BRAND, fontWeight: 600, marginBottom: 6 }}>
            Selected: {locations.find((l) => l.id === selectedIds[0])?.description
              || shortPath(locations.find((l) => l.id === selectedIds[0])?.path || "")}
          </div>
        )}
        {selectedCount > 0 && multiMode && (
          <div style={{ fontSize: 12, color: BRAND, fontWeight: 600, marginBottom: 6 }}>
            {selectedCount} location{selectedCount > 1 ? "s" : ""} selected
          </div>
        )}
        <button
          style={{ ...S.fileBtn, ...(isDisabled ? S.fileBtnDisabled : {}) }}
          disabled={isDisabled}
          onClick={handleFile}
        >
          {btnLabel}
        </button>
      </div>

    </div>
  );
}
