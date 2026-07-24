/**
 * MobileOptionsScreen.jsx
 *
 * Full settings screen for Koyomail Mobile.
 *
 * Accordion sections:
 *   🔌 Connection    — Agent URL, API Token, Test + Save
 *   📁 Filing        — After-filing action, Attachments, Category, Duplicates, UTC, Reviewed
 *   🔍 Search        — Disable Delete, Disable Transfer
 *   ℹ️  About
 *
 * All settings persist immediately to localStorage (koyomail_options)
 * and sync to the backend via updatePreferences(), matching the desktop pattern.
 */

import * as React from "react";
import {
  PlugConnected24Regular,
  FolderArrowRight24Regular,
  Search24Regular,
  Info24Regular,
  Globe24Regular,
  Checkmark24Regular,
  ChevronDown20Regular,
} from "@fluentui/react-icons";
import { getResolvedBaseUrl, initApiBaseUrl, updatePreferences } from "../services/backendApi";

// ─── Helpers ──────────────────────────────────────────────────────────────────

function loadOpts() {
  try { return JSON.parse(localStorage.getItem("koyomail_options") || "{}"); }
  catch { return {}; }
}

function saveOpt(key, value) {
  try {
    const opts = loadOpts();
    // Normalise legacy key
    const v = key === "afterFilingAction" && value === "move_deleted" ? "delete" : value;
    opts[key] = v;
    localStorage.setItem("koyomail_options", JSON.stringify(opts));
    window.dispatchEvent(new Event("koyomail_options_updated"));
    updatePreferences({ [key]: v }).catch(() => {});
  } catch (e) {
    console.warn("[MobileOptionsScreen] saveOpt failed:", e);
  }
}

// ─── Design tokens ────────────────────────────────────────────────────────────

const BRAND = "#0078d4";
const BRAND_LIGHT = "#e3f2fd";
const BORDER = "#e8e8e8";
const TEXT_MUTED = "#888";
const TEXT_HINT = "#b0b0b0";

// ─── Sub-components ───────────────────────────────────────────────────────────

/**
 * Animated toggle switch.
 */
function Toggle({ on, onChange }) {
  return (
    <div
      role="switch"
      aria-checked={on}
      onClick={() => onChange(!on)}
      style={{
        width: 46,
        height: 26,
        borderRadius: 13,
        background: on ? BRAND : "#d0d0d0",
        position: "relative",
        flexShrink: 0,
        cursor: "pointer",
        transition: "background .2s",
        WebkitTapHighlightColor: "transparent",
      }}
    >
      <div style={{
        position: "absolute",
        top: 3,
        left: on ? 23 : 3,
        width: 20,
        height: 20,
        borderRadius: "50%",
        background: "#fff",
        transition: "left .2s",
        boxShadow: "0 1px 4px rgba(0,0,0,.25)",
      }} />
    </div>
  );
}

/**
 * A single settings row with label, optional description, and a control on the right.
 */
function SettingRow({ label, description, last, children }) {
  return (
    <div style={{
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: 14,
      paddingTop: 12,
      paddingBottom: 12,
      borderBottom: last ? "none" : `1px solid ${BORDER}`,
    }}>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 14, fontWeight: 500, color: "#222", lineHeight: 1.3 }}>
          {label}
        </div>
        {description && (
          <div style={{ fontSize: 11, color: TEXT_MUTED, lineHeight: 1.45, marginTop: 2 }}>
            {description}
          </div>
        )}
      </div>
      <div style={{ flexShrink: 0 }}>
        {children}
      </div>
    </div>
  );
}

/**
 * A select dropdown row with label above.
 */
function SelectRow({ label, description, value, onChange, options, last }) {
  return (
    <div style={{
      paddingTop: 12,
      paddingBottom: 12,
      borderBottom: last ? "none" : `1px solid ${BORDER}`,
    }}>
      <div style={{ fontSize: 14, fontWeight: 500, color: "#222", marginBottom: description ? 3 : 8 }}>
        {label}
      </div>
      {description && (
        <div style={{ fontSize: 11, color: TEXT_MUTED, lineHeight: 1.45, marginBottom: 8 }}>
          {description}
        </div>
      )}
      <select
        style={{
          width: "100%",
          padding: "10px 12px",
          borderRadius: 9,
          border: `1.5px solid ${BORDER}`,
          fontSize: 13,
          background: "#fff",
          color: "#1a1a1a",
          outline: "none",
          boxSizing: "border-box",
          appearance: "auto",
        }}
        value={value}
        onChange={(e) => onChange(e.target.value)}
      >
        {options.map((o) => (
          <option key={o.value} value={o.value}>{o.label}</option>
        ))}
      </select>
    </div>
  );
}

/**
 * Accordion section with icon, title, open/close chevron, and optional badge.
 */
function AccordionSection({ title, icon, badge, badgeColor, isOpen, onToggle, children }) {
  return (
    <div style={{
      background: "#fff",
      borderRadius: 14,
      border: `1px solid ${BORDER}`,
      marginBottom: 12,
      overflow: "hidden",
      boxShadow: "0 1px 4px rgba(0,0,0,.05)",
    }}>
      {/* Section header */}
      <div
        onClick={onToggle}
        style={{
          display: "flex",
          alignItems: "center",
          padding: "14px 16px",
          cursor: "pointer",
          gap: 10,
          WebkitTapHighlightColor: "transparent",
          userSelect: "none",
        }}
      >
        <span style={{
          display: "inline-flex",
          alignItems: "center",
          flexShrink: 0,
          color: "#0078d4",
        }}>
          {icon}
        </span>
        <span style={{ flex: 1, fontWeight: 700, fontSize: 14, color: "#111", letterSpacing: "-0.01em" }}>
          {title}
        </span>
        {badge && (
          <span style={{
            fontSize: 10,
            fontWeight: 700,
            background: badgeColor ? (badgeColor + "22") : BRAND_LIGHT,
            color: badgeColor || BRAND,
            borderRadius: 10,
            padding: "2px 8px",
            letterSpacing: "0.04em",
            border: `1px solid ${badgeColor ? (badgeColor + "44") : "#c3e0f5"}`,
          }}>
            {badge}
          </span>
        )}
        <span style={{
          display: "inline-flex",
          alignItems: "center",
          color: "#bbb",
          transform: isOpen ? "rotate(180deg)" : "rotate(0deg)",
          transition: "transform .22s ease",
        }}>
          <ChevronDown20Regular />
        </span>
      </div>

      {/* Collapsible body */}
      {isOpen && (
        <div style={{
          borderTop: `1px solid #f5f5f5`,
          padding: "0 16px 4px",
          animation: "kmFadeIn .15s ease",
        }}>
          {children}
        </div>
      )}
    </div>
  );
}

// ─── Main component ────────────────────────────────────────────────────────────

export default function MobileOptionsScreen() {
  const saved = loadOpts();

  // ── Section open/close ──────────────────────────────────────────────────────
  const [connectionOpen, setConnectionOpen] = React.useState(true);
  const [filingOpen, setFilingOpen] = React.useState(false);
  const [searchSectionOpen, setSearchSectionOpen] = React.useState(false);
  const [aboutOpen, setAboutOpen] = React.useState(false);

  // ── Connection ──────────────────────────────────────────────────────────────
  const [agentUrl, setAgentUrl] = React.useState(saved.agentUrl || "");
  const [agentToken, setAgentToken] = React.useState(saved.agentToken || "");
  const [testStatus, setTestStatus] = React.useState(null); // null | "checking" | "ok" | "error"
  const [testMsg, setTestMsg] = React.useState("");
  const [saveMsg, setSaveMsg] = React.useState("");

  // ── Filing options ──────────────────────────────────────────────────────────
  const [afterFilingAction, setAfterFilingAction] = React.useState(() => {
    const raw = saved.afterFilingAction;
    return raw === "move_deleted" ? "delete" : (raw || "none");
  });
  const [defaultAttachments, setDefaultAttachments] = React.useState(saved.defaultAttachments || "all");
  const [addFiledCategory, setAddFiledCategory] = React.useState(saved.addFiledCategory !== false);
  const [filedCategoryName, setFiledCategoryName] = React.useState(saved.filedCategoryName || "Filed by Koyomail");
  const [categoryNameDraft, setCategoryNameDraft] = React.useState(saved.filedCategoryName || "Filed by Koyomail");
  const [duplicateStrategy, setDuplicateStrategy] = React.useState(saved.duplicateStrategy || "overwrite");
  const [useUtcTime, setUseUtcTime] = React.useState(!!saved.useUtcTime);
  const [markReviewed, setMarkReviewed] = React.useState(!!saved.markReviewed);

  // ── Search options ──────────────────────────────────────────────────────────
  const [disableDelete, setDisableDelete] = React.useState(!!saved.disableDelete);
  const [disableMoveTo, setDisableMoveTo] = React.useState(!!saved.disableMoveTo);

  // ── Toast ────────────────────────────────────────────────────────────────────
  const [toast, setToast] = React.useState(null);
  const toastTimerRef = React.useRef(null);

  const showToast = React.useCallback((msg, type = "ok") => {
    clearTimeout(toastTimerRef.current);
    setToast({ msg, type });
    toastTimerRef.current = setTimeout(() => setToast(null), 2200);
  }, []);

  React.useEffect(() => () => clearTimeout(toastTimerRef.current), []);

  // ── Generic change helper ───────────────────────────────────────────────────
  const handleChange = React.useCallback((key, value, setter) => {
    setter(value);
    saveOpt(key, value);
    showToast("Saved");
  }, [showToast]);

  // ── Connection handlers ─────────────────────────────────────────────────────
  const handleSaveConnection = () => {
    const opts = loadOpts();
    opts.agentUrl = agentUrl.trim().replace(/\/$/, "");
    opts.agentToken = agentToken.trim();
    localStorage.setItem("koyomail_options", JSON.stringify(opts));
    initApiBaseUrl();
    setSaveMsg("Saved ✓");
    showToast("Connection saved!");
    setTimeout(() => setSaveMsg(""), 3000);
  };

  const handleTest = async () => {
    const url = agentUrl.trim().replace(/\/$/, "");
    if (!url) {
      setTestStatus("error");
      setTestMsg("Please enter an Agent URL first.");
      return;
    }
    setTestStatus("checking");
    setTestMsg("Connecting…");
    try {
      const ctrl = new AbortController();
      const tid = setTimeout(() => ctrl.abort(), 7000);
      const resp = await fetch(`${url}/api/health`, {
        signal: ctrl.signal,
        headers: { "ngrok-skip-browser-warning": "true", Accept: "application/json" },
      });
      clearTimeout(tid);
      if (resp.ok) {
        const data = await resp.json().catch(() => ({}));
        setTestStatus("ok");
        setTestMsg(`✅ Connected — ${data.service || "backend"} is running`);
      } else {
        setTestStatus("error");
        setTestMsg(`❌ HTTP ${resp.status} from agent`);
      }
    } catch (e) {
      setTestStatus("error");
      setTestMsg(`❌ Cannot reach agent: ${e.message}`);
    }
  };

  // Add the missing import for ChevronDown
  // (already imported above via named imports)
  const currentUrl = getResolvedBaseUrl();
  const isLocalhost = currentUrl.includes("localhost");
  const connectionBadge = isLocalhost ? "Local" : "Remote";
  const connectionBadgeColor = isLocalhost ? "#2e7d32" : "#0078d4";

  // ── Styles (inline, scoped) ────────────────────────────────────────────────

  const inputStyle = {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 9,
    border: `1.5px solid ${BORDER}`,
    fontSize: 14,
    boxSizing: "border-box",
    outline: "none",
    fontFamily: "inherit",
    color: "#111",
    background: "#fff",
    transition: "border-color .15s",
  };

  const labelStyle = {
    display: "block",
    marginBottom: 5,
    fontWeight: 600,
    fontSize: 12,
    color: TEXT_MUTED,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  };

  const hintStyle = {
    fontSize: 11,
    color: TEXT_HINT,
    marginTop: 4,
    lineHeight: 1.45,
  };

  const primaryBtn = {
    flex: 1,
    padding: "12px 0",
    borderRadius: 10,
    border: "none",
    background: BRAND,
    color: "#fff",
    fontWeight: 700,
    fontSize: 14,
    cursor: "pointer",
    transition: "opacity .15s",
    WebkitTapHighlightColor: "transparent",
  };

  const outlineBtn = {
    flex: 1,
    padding: "12px 0",
    borderRadius: 10,
    border: `1.5px solid ${BRAND}`,
    background: "#fff",
    color: BRAND,
    fontWeight: 700,
    fontSize: 14,
    cursor: "pointer",
    transition: "opacity .15s",
    WebkitTapHighlightColor: "transparent",
  };

  // ── Render ─────────────────────────────────────────────────────────────────

  return (
    <div style={{
      height: "100%",
      overflowY: "auto",
      padding: "14px 14px 90px",
      fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif",
      boxSizing: "border-box",
      background: "#f7f8fa",
    }}>

      {/* Keyframe animations */}
      <style>{`
        @keyframes kmFadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes kmSlideUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
      `}</style>

      {/* ── Toast notification ── */}
      {toast && (
        <div style={{
          position: "fixed",
          bottom: 76,
          left: "50%",
          transform: "translateX(-50%)",
          background: toast.type === "ok" ? "#1e7e34" : "#c62828",
          color: "#fff",
          fontSize: 13,
          fontWeight: 700,
          padding: "9px 22px",
          borderRadius: 22,
          zIndex: 9999,
          boxShadow: "0 4px 16px rgba(0,0,0,.2)",
          whiteSpace: "nowrap",
          animation: "kmSlideUp .2s ease",
          pointerEvents: "none",
        }}>
          {toast.msg}
        </div>
      )}

      {/* ── Page heading ── */}
      <div style={{ marginBottom: 18 }}>
        <h1 style={{ margin: 0, fontSize: 22, fontWeight: 800, color: "#111", letterSpacing: "-0.02em" }}>
          Settings
        </h1>
        <p style={{ margin: "4px 0 0", fontSize: 12, color: TEXT_MUTED }}>
          Koyomail Mobile preferences
        </p>
      </div>

      {/* ════════════════ CONNECTION SECTION ════════════════ */}
      <AccordionSection
        title="Connection"
        icon={<PlugConnected24Regular />}
        badge={connectionBadge}
        badgeColor={connectionBadgeColor}
        isOpen={connectionOpen}
        onToggle={() => setConnectionOpen((o) => !o)}
      >
        {/* Active endpoint chip */}
        <div style={{ margin: "10px 0 14px" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: TEXT_MUTED, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 5 }}>
            Active endpoint
          </div>
          <div style={{
            padding: "8px 12px",
            background: "#f5f5f5",
            borderRadius: 8,
            fontSize: 12,
            color: "#555",
            wordBreak: "break-all",
            fontFamily: "monospace",
            lineHeight: 1.4,
          }}>
            {currentUrl}
          </div>
        </div>

        {/* Agent URL */}
        <div style={{ marginBottom: 12 }}>
          <label style={labelStyle}>Agent URL</label>
          <input
            style={inputStyle}
            type="url"
            placeholder="https://koyomail.yourcompany.com"
            value={agentUrl}
            onChange={(e) => setAgentUrl(e.target.value)}
            autoCapitalize="none"
            autoCorrect="off"
            spellCheck={false}
          />
          <p style={hintStyle}>Cloudflare Tunnel URL provided by your IT team.</p>
        </div>

        {/* API Token */}
        <div style={{ marginBottom: 14 }}>
          <label style={labelStyle}>API Token</label>
          <input
            style={inputStyle}
            type="password"
            placeholder="Paste the token provided by IT"
            value={agentToken}
            onChange={(e) => setAgentToken(e.target.value)}
            autoCapitalize="none"
            autoCorrect="off"
            spellCheck={false}
          />
          <p style={hintStyle}>Secret token set during agent installation.</p>
        </div>

        {/* Test result */}
        {testStatus && (
          <div style={{
            padding: "10px 12px",
            borderRadius: 9,
            fontSize: 13,
            fontWeight: 500,
            marginBottom: 12,
            lineHeight: 1.4,
            background: testStatus === "ok" ? "#e6f4ea" : testStatus === "error" ? "#fce8e6" : "#f0f0f0",
            color: testStatus === "ok" ? "#2e7d32" : testStatus === "error" ? "#c62828" : "#555",
          }}>
            {testMsg}
          </div>
        )}

        {/* Buttons */}
        <div style={{ display: "flex", gap: 10, marginBottom: 4 }}>
          <button style={outlineBtn} onClick={handleTest}>Test</button>
          <button style={primaryBtn} onClick={handleSaveConnection}>
            {saveMsg || "Save"}
          </button>
        </div>
      </AccordionSection>

      {/* ════════════════ FILING SECTION ════════════════ */}
      <AccordionSection
        title="Filing Options"
        icon={<FolderArrowRight24Regular />}
        isOpen={filingOpen}
        onToggle={() => setFilingOpen((o) => !o)}
      >
        <SelectRow
          label="After filing action"
          description="What happens to the email after it is filed to disk"
          value={afterFilingAction}
          onChange={(v) => handleChange("afterFilingAction", v, setAfterFilingAction)}
          options={[
            { value: "none",              label: "Keep in Inbox" },
            { value: "add_date",          label: "Add filed date & time to subject" },
            { value: "archive",           label: "Move to Archive" },
            { value: "delete",            label: "Move to Deleted Items" },
            { value: "move_filed_items",  label: "Move to \"Filed Items\" folder" },
            { value: "move_filed_folders",label: "Move to Filed sub-folders" },
          ]}
        />

        <SelectRow
          label="Default attachments"
          description="What to include when filing an email"
          value={defaultAttachments}
          onChange={(v) => handleChange("defaultAttachments", v, setDefaultAttachments)}
          options={[
            { value: "all",         label: "File message with attachments" },
            { value: "message",     label: "File message only" },
            { value: "attachments", label: "File attachments only" },
          ]}
        />

        <SelectRow
          label="Duplicate handling"
          description="When a file with the same name already exists in the folder"
          value={duplicateStrategy}
          onChange={(v) => handleChange("duplicateStrategy", v, setDuplicateStrategy)}
          options={[
            { value: "overwrite", label: "Overwrite existing file" },
            { value: "skip",      label: "Skip (keep existing)" },
            { value: "rename",    label: "Rename (add number suffix)" },
          ]}
        />

        <SettingRow
          label="Apply filed category"
          description="Assign an Outlook category to visually mark filed emails"
        >
          <Toggle on={addFiledCategory} onChange={(v) => handleChange("addFiledCategory", v, setAddFiledCategory)} />
        </SettingRow>

        {/* Category name input — only visible when category is on */}
        {addFiledCategory && (
          <div style={{
            paddingTop: 10,
            paddingBottom: 12,
            borderBottom: `1px solid ${BORDER}`,
            animation: "kmFadeIn .15s ease",
          }}>
            <label style={labelStyle}>Category name</label>
            <div style={{ display: "flex", gap: 8 }}>
              <input
                style={{ ...inputStyle, flex: 1 }}
                value={categoryNameDraft}
                onChange={(e) => setCategoryNameDraft(e.target.value)}
                placeholder="Filed by Koyomail"
              />
              <button
                style={{
                  padding: "10px 14px",
                  borderRadius: 9,
                  border: "none",
                  background: BRAND,
                  color: "#fff",
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: "pointer",
                  flexShrink: 0,
                  WebkitTapHighlightColor: "transparent",
                }}
                onClick={() => {
                  setFiledCategoryName(categoryNameDraft);
                  handleChange("filedCategoryName", categoryNameDraft, setFiledCategoryName);
                }}
              >
                Save
              </button>
            </div>
          </div>
        )}

        <SettingRow
          label="Use UTC filing time"
          description="Save timestamps in UTC instead of local time zone"
        >
          <Toggle on={useUtcTime} onChange={(v) => handleChange("useUtcTime", v, setUseUtcTime)} />
        </SettingRow>

        <SettingRow
          label="Mark subject as reviewed"
          description="Appends a reviewed indicator to the email subject after filing"
          last
        >
          <Toggle on={markReviewed} onChange={(v) => handleChange("markReviewed", v, setMarkReviewed)} />
        </SettingRow>
      </AccordionSection>

      {/* ════════════════ SEARCH SECTION ════════════════ */}
      <AccordionSection
        title="Search Options"
        icon={<Search24Regular />}
        isOpen={searchSectionOpen}
        onToggle={() => setSearchSectionOpen((o) => !o)}
      >
        <SettingRow
          label="Disable the Delete option"
          description="Hides the Delete button in search results (useful for shared installs)"
        >
          <Toggle on={disableDelete} onChange={(v) => handleChange("disableDelete", v, setDisableDelete)} />
        </SettingRow>

        <SettingRow
          label="Disable the Transfer option"
          description="Hides the Move / Transfer button in search results"
          last
        >
          <Toggle on={disableMoveTo} onChange={(v) => handleChange("disableMoveTo", v, setDisableMoveTo)} />
        </SettingRow>
      </AccordionSection>

      {/* ════════════════ ABOUT SECTION ════════════════ */}
      <AccordionSection
        title="About"
        icon={<Info24Regular />}
        isOpen={aboutOpen}
        onToggle={() => setAboutOpen((o) => !o)}
      >
        <p style={{ fontSize: 13, color: "#666", lineHeight: 1.65, margin: "10px 0 4px" }}>
          <strong>Koyomail Mobile</strong> connects to an on-site agent running on your
          company's server. The agent provides secure access to your email archive on
          network drives without copying any data to the cloud.
        </p>
        <p style={{ fontSize: 12, color: TEXT_HINT, margin: "8px 0 8px", lineHeight: 1.5 }}>
          All settings are stored locally on this device and synced securely to the
          on-site agent when a connection is available.
        </p>
      </AccordionSection>

    </div>
  );
}
