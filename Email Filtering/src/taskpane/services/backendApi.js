import { toErrorMessage } from "../utils/errorUtils.js";

// ─── Backend URL Resolution ──────────────────────────────────────────────────
//
// Priority order:
//   1. Build-time env var  (local dev / CI)
//   2. Saved agentUrl from localStorage koyomail_options  (mobile / remote desktop)
//   3. https://localhost:4000  (default — desktop workstation)

function _getSavedAgentUrl() {
  try {
    const opts = JSON.parse(localStorage.getItem("koyomail_options") || "{}");
    return opts.agentUrl ? opts.agentUrl.replace(/\/$/, "") : null;
  } catch { return null; }
}

function _getSavedAgentToken() {
  try {
    const opts = JSON.parse(localStorage.getItem("koyomail_options") || "{}");
    return opts.agentToken || null;
  } catch { return null; }
}

// Runtime-resolved base URL — updated once by initApiBaseUrl() at app startup
let _resolvedBaseUrl = process.env.API_BASE_URL || "https://localhost:4000";

/**
 * Attempt to reach the local backend.
 * Returns "local" if localhost:4000 responds within 2 s, otherwise "remote".
 */
export async function detectBackendMode() {
  if (process.env.API_BASE_URL) return "local"; // dev override — always local
  try {
    const controller = new AbortController();
    const tid = setTimeout(() => controller.abort(), 2000);
    const resp = await fetch("https://localhost:4000/api/health", {
      signal: controller.signal,
    });
    clearTimeout(tid);
    if (resp.ok) return "local";
  } catch { /* not reachable */ }
  return "remote";
}

/**
 * Call once at app startup (in Office.onReady) before the first render.
 * Sets the module-level resolved URL and returns it.
 */
export async function initApiBaseUrl() {
  if (process.env.API_BASE_URL) {
    _resolvedBaseUrl = process.env.API_BASE_URL;
    console.log(`[backendApi] Mode: dev-override → ${_resolvedBaseUrl}`);
    return _resolvedBaseUrl;
  }
  const savedUrl = _getSavedAgentUrl();
  if (savedUrl) {
    _resolvedBaseUrl = savedUrl;
    console.log(`[backendApi] Mode: saved-agent → ${_resolvedBaseUrl}`);
    return _resolvedBaseUrl;
  }
  const mode = await detectBackendMode();
  _resolvedBaseUrl = "https://localhost:4000";
  console.log(`[backendApi] Mode: ${mode} → ${_resolvedBaseUrl}`);
  return _resolvedBaseUrl;
}

/** Returns the currently resolved base URL. */
export function getResolvedBaseUrl() {
  return _resolvedBaseUrl;
}

/**
 * Kept for backward compatibility — modules that imported API_BASE_URL as a
 * constant still work. For runtime value always use getResolvedBaseUrl().
 */
export const API_BASE_URL = _resolvedBaseUrl;

// ─── HTTP helper ─────────────────────────────────────────────────────────────

/**
 * remoteLog — fire-and-forget logger that sends auth diagnostics to the
 * backend terminal. Use this when DevTools is unavailable (e.g. New Outlook).
 *
 * @param {"info"|"warn"|"error"|"ok"} level
 * @param {string} message
 * @param {object} [data]   optional extra data to print (JSON)
 */
export function remoteLog(level, message, data) {
  try {
    fetch(`${getResolvedBaseUrl()}/api/debug/auth-log`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ level, message, data }),
    }).catch(() => {}); // swallow network errors silently
  } catch {
    // Never let logging break auth flow
  }
}

async function request(path, options = {}) {
  const baseUrl = getResolvedBaseUrl();
  const isRemote = !baseUrl.includes("localhost");

  const headers = {
    "Content-Type": "application/json",
    "ngrok-skip-browser-warning": "true",
    ...(options.headers || {}),
  };

  // Inject agent API token for remote (mobile / agent) connections only
  if (isRemote) {
    const token = _getSavedAgentToken();
    if (token) {
      headers["x-koyomail-token"] = token;
    }
  }

  const response = await fetch(`${baseUrl}${path}`, {
    headers,
    ...options,
  });

  if (response.status === 204) {
    return null;
  }

  const raw = await response.text();
  let data = null;
  if (raw) {
    try {
      data = JSON.parse(raw);
    } catch {
      data = raw;
    }
  }

  if (!response.ok) {
    const fallback = `Request failed (${response.status})`;
    throw new Error(toErrorMessage(data, fallback));
  }

  return data;
}

// ─── API functions (unchanged signatures) ────────────────────────────────────

export function getLocations(options = {}) {
  const params = new URLSearchParams();
  params.set("_t", Date.now());
  if (options.sender) {
    params.set("sender", options.sender);
  }
  return request(`/api/locations?${params.toString()}`);
}

export function getSenderHistory(sender) {
  const params = new URLSearchParams();
  params.set("_t", Date.now());
  if (sender) {
    params.set("sender", sender);
  }
  return request(`/api/locations/sender-history?${params.toString()}`);
}

export function addLocation(payload) {
  return request("/api/locations", {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export function updateLocation(id, payload) {
  return request(`/api/locations/${id}`, {
    method: "PUT",
    body: JSON.stringify(payload),
  });
}

export function deleteLocation(id) {
  return request(`/api/locations/${id}`, {
    method: "DELETE",
  });
}

export function fileEmail(payload, options = {}) {
  return request("/api/file/email", {
    method: "POST",
    body: JSON.stringify(payload),
    ...options,
  });
}

export function createDraftEmail(payload, options = {}) {
  return request("/api/file/draft", {
    method: "POST",
    body: JSON.stringify(payload),
    ...options,
  });
}

export function applyPostFilingActions(payload, options = {}) {
  return request("/api/file/post-filing", {
    method: "POST",
    body: JSON.stringify(payload),
    ...options,
  });
}

export async function getConnectivityStatus() {
  try {
    const data = await request(`/api/locations/status?_t=${Date.now()}`);
    return data || {};
  } catch (error) {
    console.error("Connectivity check failed:", error);
    return {};
  }
}

export async function checkPathsConnectivity(paths) {
  try {
    const data = await request(`/api/locations/status/check`, {
      method: "POST",
      body: JSON.stringify({ paths }),
    });
    return data || {};
  } catch (error) {
    console.error("Paths connectivity check failed:", error);
    return {};
  }
}

export function exploreLocation(path) {
  return request("/api/locations/explore", {
    method: "POST",
    body: JSON.stringify({ path }),
  });
}

export function removeSuggestion(id, sender) {
  const params = new URLSearchParams();
  if (sender) {
    params.set("sender", sender);
  }
  return request(`/api/locations/${id}/remove-suggestion?${params.toString()}`, {
    method: "POST",
  });
}

export function toggleSuggestion(id, sender) {
  const params = new URLSearchParams();
  if (sender) {
    params.set("sender", sender);
  }
  return request(`/api/locations/${id}/toggle-suggestion?${params.toString()}`, {
    method: "POST",
  });
}

export function markLocationUnused(id) {
  return request(`/api/locations/${id}/mark-unused`, {
    method: "POST",
  });
}

export function searchEmails(params) {
  return request(`/api/search?${params.toString()}`);
}

export function getSearchPreview(params) {
  return request(`/api/search/preview?${params.toString()}`);
}

/**
 * Build a download URL that works with window.open (no headers possible),
 * appending the agent token as a query param for remote connections.
 */
export function buildDownloadUrl(params) {
  const baseUrl = getResolvedBaseUrl();
  if (!baseUrl.includes("localhost")) {
    const token = _getSavedAgentToken();
    if (token) {
      params.set("_token", token);
    }
  }
  return `${baseUrl}/api/search/download?${params.toString()}`;
}

export function getPreferences() {
  return request(`/api/preferences?_t=${Date.now()}`);
}

export function updatePreferences(payload) {
  return request("/api/preferences", {
    method: "PUT",
    body: JSON.stringify(payload),
  });
}
