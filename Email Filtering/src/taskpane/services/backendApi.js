import { toErrorMessage } from "../utils/errorUtils.js";

const BASE_URL = "http://localhost:4000";

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

export function getLocations() {
  return request("/api/locations");
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

export function fileEmail(payload) {
  return request("/api/file/email", {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export async function getConnectivityStatus() {
  try {
    // Backend returns an object: { [id]: boolean }
    const data = await request("/api/locations/status");
    return data || {};
  } catch (error) {
    console.error("Connectivity check failed:", error);
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

export function testGraphApi(ssoToken) {
  return request("/api/file/test-graph", {
    method: "POST",
    body: JSON.stringify({ ssoToken }),
  });
}
