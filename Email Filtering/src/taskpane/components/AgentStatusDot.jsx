/**
 * AgentStatusDot.jsx
 *
 * A small coloured indicator shown in the mobile header.
 * Pings GET /api/health on the configured agent URL every 60 s.
 *   🟢 green  — agent reachable
 *   🔴 red    — agent unreachable or not configured
 *   ⚪ grey   — currently checking
 */

import * as React from "react";
import { getResolvedBaseUrl } from "../services/backendApi";

const DOT = {
  display: "inline-block",
  width: 10,
  height: 10,
  borderRadius: "50%",
  marginLeft: 6,
  flexShrink: 0,
  cursor: "pointer",
};

const COLORS = {
  checking: "#9e9e9e",
  ok: "#2e7d32",
  error: "#c62828",
};

const LABELS = {
  checking: "Checking agent…",
  ok: "Agent connected",
  error: "Agent not reachable",
};

export default function AgentStatusDot() {
  const [status, setStatus] = React.useState("checking");
  const [showTip, setShowTip] = React.useState(false);

  const check = React.useCallback(async () => {
    setStatus("checking");
    try {
      const baseUrl = getResolvedBaseUrl();
      if (baseUrl.includes("localhost")) {
        // Running as a local backend — not an agent; show ok silently
        setStatus("ok");
        return;
      }
      const ctrl = new AbortController();
      const tid = setTimeout(() => ctrl.abort(), 5000);
      const resp = await fetch(`${baseUrl}/api/health`, { signal: ctrl.signal });
      clearTimeout(tid);
      setStatus(resp.ok ? "ok" : "error");
    } catch {
      setStatus("error");
    }
  }, []);

  React.useEffect(() => {
    check();
    const interval = setInterval(check, 60_000);
    return () => clearInterval(interval);
  }, [check]);

  return (
    <span style={{ position: "relative", display: "inline-flex", alignItems: "center" }}>
      <span
        role="img"
        aria-label={LABELS[status]}
        style={{ ...DOT, background: COLORS[status] }}
        onClick={() => setShowTip((v) => !v)}
      />
      {showTip && (
        <span
          style={{
            position: "absolute",
            top: 18,
            right: 0,
            background: "#1e1e1e",
            color: "#fff",
            fontSize: 11,
            padding: "4px 8px",
            borderRadius: 4,
            whiteSpace: "nowrap",
            zIndex: 999,
            boxShadow: "0 2px 8px rgba(0,0,0,.35)",
          }}
        >
          {LABELS[status]} — {getResolvedBaseUrl()}
        </span>
      )}
    </span>
  );
}
