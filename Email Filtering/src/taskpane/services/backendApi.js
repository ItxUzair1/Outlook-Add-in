import { toErrorMessage } from "../utils/errorUtils.js";

export const API_BASE_URL = process.env.API_BASE_URL || "https://localhost:4000";
const BASE_URL = API_BASE_URL;

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
    fetch(`${BASE_URL}/api/debug/auth-log`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ level, message, data }),
    }).catch(() => {}); // swallow network errors silently
  } catch {
    // Never let logging break auth flow
  }
}

async function request(path, options = {}) {
  const response = await fetch(`${BASE_URL}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
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

export function getLocations(options = {}) {
  const params = new URLSearchParams();
  params.set("_t", Date.now());
  if (options.sender) {
    params.set("sender", options.sender);
  }
  return request(`/api/locations?${params.toString()}`);
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

export async function getConnectivityStatus() {
  try {
    // Backend returns an object: { [id]: boolean }
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

export function removeSuggestion(id) {
  return request(`/api/locations/${id}/remove-suggestion`, {
    method: "POST",
  });
}

export function toggleSuggestion(id) {
  return request(`/api/locations/${id}/toggle-suggestion`, {
    method: "POST",
  });
}

export function markLocationUnused(id) {
  return request(`/api/locations/${id}/mark-unused`, {
    method: "POST",
  });
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


