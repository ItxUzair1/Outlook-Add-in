/**
 * MobileOptionsScreen.jsx
 *
 * Lets the user configure the on-site agent connection details (URL + token).
 * Only shown in the mobile taskpane — not used on desktop.
 */

import * as React from "react";
import { getResolvedBaseUrl, initApiBaseUrl } from "../services/backendApi";

const styles = {
  container: {
    padding: "16px",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    color: "#1a1a1a",
  },
  section: {
    marginBottom: 24,
  },
  sectionTitle: {
    fontWeight: 600,
    fontSize: 13,
    color: "#555",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    marginBottom: 12,
  },
  label: {
    display: "block",
    marginBottom: 4,
    fontWeight: 500,
    fontSize: 13,
  },
  input: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 8,
    border: "1.5px solid #d0d0d0",
    fontSize: 14,
    boxSizing: "border-box",
    marginBottom: 12,
    outline: "none",
    transition: "border-color .15s",
  },
  row: {
    display: "flex",
    gap: 10,
    marginTop: 4,
  },
  btn: {
    flex: 1,
    padding: "11px 0",
    borderRadius: 8,
    border: "none",
    cursor: "pointer",
    fontWeight: 600,
    fontSize: 14,
    transition: "opacity .15s",
  },
  btnPrimary: {
    background: "#0078d4",
    color: "#fff",
  },
  btnOutline: {
    background: "#fff",
    color: "#0078d4",
    border: "1.5px solid #0078d4",
  },
  statusBox: {
    marginTop: 10,
    padding: "10px 12px",
    borderRadius: 8,
    fontSize: 13,
    fontWeight: 500,
  },
  statusOk: {
    background: "#e6f4ea",
    color: "#2e7d32",
  },
  statusError: {
    background: "#fce8e6",
    color: "#c62828",
  },
  hint: {
    fontSize: 12,
    color: "#777",
    marginTop: -8,
    marginBottom: 12,
  },
};

function loadSavedOpts() {
  try {
    return JSON.parse(localStorage.getItem("koyomail_options") || "{}");
  } catch { return {}; }
}

export default function MobileOptionsScreen() {
  const saved = loadSavedOpts();
  const [agentUrl, setAgentUrl] = React.useState(saved.agentUrl || "");
  const [agentToken, setAgentToken] = React.useState(saved.agentToken || "");
  const [testStatus, setTestStatus] = React.useState(null); // null | "ok" | "error" | "checking"
  const [testMsg, setTestMsg] = React.useState("");
  const [saveMsg, setSaveMsg] = React.useState("");

  const handleSave = () => {
    try {
      const opts = JSON.parse(localStorage.getItem("koyomail_options") || "{}");
      opts.agentUrl = agentUrl.trim().replace(/\/$/, "");
      opts.agentToken = agentToken.trim();
      localStorage.setItem("koyomail_options", JSON.stringify(opts));
      initApiBaseUrl();
      setSaveMsg("Settings saved!");
      setTimeout(() => setSaveMsg(""), 3000);
    } catch (e) {
      setSaveMsg("Failed to save to local storage.");
    }
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
      const tid = setTimeout(() => ctrl.abort(), 6000);
      const resp = await fetch(`${url}/api/health`, {
        signal: ctrl.signal,
        headers: {
          "ngrok-skip-browser-warning": "true",
          "Accept": "application/json"
        }
      });
      clearTimeout(tid);
      if (resp.ok) {
        const data = await resp.json().catch(() => ({}));
        setTestStatus("ok");
        setTestMsg(`✅ Connected — ${data.service || "backend"} is running`);
      } else {
        setTestStatus("error");
        setTestMsg(`❌ Agent responded with status ${resp.status}`);
      }
    } catch (e) {
      setTestStatus("error");
      setTestMsg(`❌ Cannot reach agent: ${e.message}`);
    }
  };

  const currentUrl = getResolvedBaseUrl();

  return (
    <div style={styles.container}>

      {/* Current connection */}
      <div style={styles.section}>
        <div style={styles.sectionTitle}>Current Connection</div>
        <div style={{
          padding: "10px 12px",
          background: "#f5f5f5",
          borderRadius: 8,
          fontSize: 12,
          color: "#555",
          wordBreak: "break-all",
        }}>
          {currentUrl}
        </div>
      </div>

      {/* Agent configuration */}
      <div style={styles.section}>
        <div style={styles.sectionTitle}>Agent Configuration</div>

        <label style={styles.label}>Agent URL</label>
        <input
          style={styles.input}
          type="url"
          placeholder="https://koyomail.yourcompany.com"
          value={agentUrl}
          onChange={(e) => setAgentUrl(e.target.value)}
          autoCapitalize="none"
          autoCorrect="off"
          spellCheck={false}
        />
        <p style={styles.hint}>The Cloudflare Tunnel URL provided by your IT team.</p>

        <label style={styles.label}>API Token</label>
        <input
          style={styles.input}
          type="password"
          placeholder="Paste the token provided by IT"
          value={agentToken}
          onChange={(e) => setAgentToken(e.target.value)}
          autoCapitalize="none"
          autoCorrect="off"
          spellCheck={false}
        />
        <p style={styles.hint}>The secret token set during agent installation.</p>

        {testStatus && (
          <div style={{
            ...styles.statusBox,
            ...(testStatus === "ok" ? styles.statusOk
              : testStatus === "error" ? styles.statusError
              : { background: "#f0f0f0", color: "#555" }),
          }}>
            {testMsg}
          </div>
        )}

        <div style={styles.row}>
          <button
            style={{ ...styles.btn, ...styles.btnOutline }}
            onClick={handleTest}
          >
            Test Connection
          </button>
          <button
            style={{ ...styles.btn, ...styles.btnPrimary }}
            onClick={handleSave}
          >
            {saveMsg ? "Saved ✓" : "Save"}
          </button>
        </div>
      </div>

      {/* Info */}
      <div style={styles.section}>
        <div style={styles.sectionTitle}>About</div>
        <p style={{ fontSize: 12, color: "#777", lineHeight: 1.5 }}>
          Koyomail Mobile connects to an on-site agent running on your company&apos;s
          server. The agent provides secure access to your email archive network drives
          without copying any data to the cloud.
        </p>
      </div>
    </div>
  );
}
