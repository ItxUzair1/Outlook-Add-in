/**
 * authManager.js — Unified three-tier authentication for Mail Manager
 *
 * TIER 1: Office SSO   — Office.auth.getAccessToken()
 *         Works silently in Classic Outlook (enterprise/M365).
 *         Zero UX, no prompts.
 *
 * TIER 2: NAA          — createNestablePublicClientApplication (MSAL v3+)
 *         Works in New Outlook. The broker intercepts acquireTokenPopup
 *         and handles it at the OS level — no iframe, no Office Dialog API.
 *
 * TIER 3: MSAL Redirect — instance.loginRedirect / acquireTokenRedirect
 *         Classic Outlook fallback. Opens sign-in inside the taskpane window.
 *         This is the existing working flow — preserved exactly as-is.
 */

/* global Office */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { msalNaaConfig, naaLoginRequest, loginRequest, TASKPANE_REDIRECT_URI } from "../authConfig";
import { remoteLog } from "../services/backendApi";

// ─── Singleton NAA client ───────────────────────────────────────────────────
let _naaPca = null;
let _naaInitialized = false;

/**
 * Lazily initialise the NAA MSAL client.
 * Returns null if NestedAppAuth is not supported on this host.
 */
async function getNaaClient() {
  if (_naaInitialized) return _naaPca;
  _naaInitialized = true;

  try {
    if (typeof Office === "undefined") {
      console.log("[authManager] Office not available — skipping NAA.");
      return null;
    }

    // Check both 1.1 and 1.0 — New Outlook may only report 1.0 support
    const naa11 = Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.1");
    const naa10 = Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.0");
    const supportedHint = naa11 || naa10;

    const naaStatus = `NAA 1.1=${naa11}, NAA 1.0=${naa10}`;
    console.log(`[authManager] NAA requirement check: ${naaStatus}`);
    remoteLog("info", `NAA requirement check: ${naaStatus} (proceeding to attempt init regardless of hint)`, {
      naa11,
      naa10,
      host: Office?.context?.diagnostics?.host,
      platform: Office?.context?.diagnostics?.platform,
      inIframe: typeof window !== "undefined" ? window.self !== window.top : "unknown",
    });

    console.log("[authManager] Attempting to initialise nestable MSAL client...");
    remoteLog("info", "Attempting createNestablePublicClientApplication()");
    
    // We try to initialize even if `isSetSupported` is false.
    // Why? Because XML sideloading in New Outlook sometimes falsely reports false 
    // for the requirement set, but the underlying host platform STILL supports NAA.
    // If the host truly doesn't support it, MSAL throws an explicit error we catch.
    _naaPca = await createNestablePublicClientApplication(msalNaaConfig);
    
    remoteLog("ok", "NAA client initialised successfully");
    return _naaPca;
  } catch (err) {
    if (err.message?.includes("NestedAppAuth") || err.message?.includes("NotSupported")) {
      console.log("[authManager] NAA NOT SUPPORTED on this host — skipping Tier 2.");
      remoteLog("warn", `NAA Initialization rejected by host: ${err.message}. Will skip Tier 2 and use Tier 3 redirect.`);
    } else {
      console.warn("[authManager] Failed to initialise NAA client:", err.message);
      remoteLog("error", `NAA client INIT FAILED: ${err.message}`);
    }
    return null;
  }
}

// ─── Token cache (in-memory + localStorage) ─────────────────────────────────
const TOKEN_CACHE_KEY = "mailManagerGraphTokenV1";

function readCachedToken() {
  try {
    const raw = localStorage.getItem(TOKEN_CACHE_KEY);
    if (!raw) return null;
    const { accessToken, expiresOn } = JSON.parse(raw);
    // Expire 2 minutes early to avoid edge-case expiry during a request
    if (!accessToken || !expiresOn || Date.now() >= expiresOn - 120_000) {
      localStorage.removeItem(TOKEN_CACHE_KEY);
      return null;
    }
    return accessToken;
  } catch {
    localStorage.removeItem(TOKEN_CACHE_KEY);
    return null;
  }
}

function cacheToken(accessToken, expiresOn) {
  if (!accessToken) return;
  const fallback = Date.now() + 45 * 60 * 1000;
  const ts = expiresOn
    ? Number(expiresOn instanceof Date ? expiresOn.getTime() : expiresOn)
    : fallback;
  localStorage.setItem(
    TOKEN_CACHE_KEY,
    JSON.stringify({ accessToken, expiresOn: Number.isFinite(ts) ? ts : fallback })
  );
}

// ─── Main export ─────────────────────────────────────────────────────────────

/**
 * getGraphToken({ msalInstance, interactive, loginHint })
 *
 * @param {object}  opts
 * @param {object}  opts.msalInstance  — The classic PublicClientApplication from index.jsx
 * @param {boolean} [opts.interactive] — If true, will fall through to interactive sign-in (Tier 3 redirect)
 * @param {string}  [opts.loginHint]   — Email address hint for NAA broker
 *
 * @returns {{ token: string, tier: string }}
 *   token — the access token string
 *   tier  — "cache" | "sso" | "naa-silent" | "naa-interactive" | "msal-silent" | "msal-redirect"
 */
export async function getGraphToken({ msalInstance, interactive = false, loginHint } = {}) {
  // ── Check in-memory / localStorage cache first ────────────────────────────
  const cached = readCachedToken();
  if (cached) {
    console.log("[authManager] ✅ Tier 0 — returning cached token.");
    remoteLog("ok", "Tier 0: Token served from cache");
    return { token: cached, tier: "cache" };
  }

  remoteLog("info", "Auth flow started", { interactive, hasLoginHint: !!loginHint });

  // ── TIER 1: Office SSO ────────────────────────────────────────────────────
  remoteLog("info", "Tier 1: Attempting Office SSO (getAccessToken)...");
  try {
    if (typeof Office !== "undefined" && Office?.auth?.getAccessToken) {
      const ssoToken = await Office.auth.getAccessToken({
        allowSignInPrompt: false,
        allowConsentPrompt: false,
        forMSGraphAccess: true,
      });
      if (ssoToken) {
        console.log("[authManager] ✅ Tier 1 — SSO token acquired.");
        remoteLog("ok", "Tier 1: SSO token acquired ✅");
        cacheToken(ssoToken, Date.now() + 55 * 60 * 1000);
        return { token: ssoToken, tier: "sso" };
      }
    }
  } catch (ssoErr) {
    const code = ssoErr?.code ?? ssoErr?.errorCode ?? "";
    console.warn(`[authManager] Tier 1 SSO failed (code ${code}):`, ssoErr.message ?? ssoErr);
    remoteLog("warn", `Tier 1: SSO FAILED — code=${code} message=${ssoErr.message ?? ssoErr}`);
  }

  // ── TIER 2: NAA (New Outlook) ─────────────────────────────────────────────
  const naaPca = await getNaaClient();
  remoteLog("info", `Tier 2: NAA client ${naaPca ? "READY" : "NOT AVAILABLE — skipping to Tier 3"}`);
  if (naaPca) {
    // 2a. Silent — use fully-qualified Graph scope URIs required by NAA broker
    try {
      const hint =
        loginHint ||
        Office?.context?.mailbox?.userProfile?.emailAddress ||
        undefined;
      const accounts = naaPca.getAllAccounts();
      const account = accounts[0] ?? undefined;

      const silentResult = await naaPca.acquireTokenSilent({
        ...naaLoginRequest,
        account,
        loginHint: hint,
      });

      if (silentResult?.accessToken) {
        console.log("[authManager] ✅ Tier 2a — NAA silent token acquired.");
        remoteLog("ok", "Tier 2a: NAA silent token acquired ✅");
        cacheToken(silentResult.accessToken, silentResult.expiresOn);
        return { token: silentResult.accessToken, tier: "naa-silent" };
      }
    } catch (naasilentErr) {
      console.warn("[authManager] Tier 2a NAA silent failed:", naasilentErr.errorCode ?? naasilentErr.message);
      remoteLog("warn", `Tier 2a: NAA silent FAILED — ${naasilentErr.errorCode ?? naasilentErr.message}`);
    }

    // 2b. Interactive (broker intercepts this — no actual popup in New Outlook)
    if (interactive) {
      try {
        const hint =
          loginHint ||
          Office?.context?.mailbox?.userProfile?.emailAddress ||
          undefined;

        console.log("[authManager] Tier 2b — NAA interactive (broker will handle)...");
        remoteLog("info", "Tier 2b: NAA acquireTokenPopup — broker intercepting...");
        const popupResult = await naaPca.acquireTokenPopup({
          ...naaLoginRequest,
          loginHint: hint,
        });

        if (popupResult?.accessToken) {
          console.log("[authManager] ✅ Tier 2b — NAA interactive token acquired.");
          remoteLog("ok", "Tier 2b: NAA interactive token acquired ✅");
          cacheToken(popupResult.accessToken, popupResult.expiresOn);
          return { token: popupResult.accessToken, tier: "naa-interactive" };
        }
      } catch (naaPopupErr) {
        const naaErrCode = naaPopupErr.errorCode ?? naaPopupErr.message ?? "unknown";
        console.warn("[authManager] Tier 2b NAA interactive failed:", naaErrCode);
        remoteLog("error", `Tier 2b: NAA interactive FAILED — ${naaErrCode}`, {
          errorCode: naaPopupErr.errorCode,
          subError: naaPopupErr.subError,
          message: naaPopupErr.message,
          stack: naaPopupErr.stack?.split("\n")[0],
        });

        // ❌ DO NOT fall through to Tier 3 when NAA host is detected.
        // In New Outlook, loginRedirect() is always blocked (redirect_in_iframe).
        // Re-throw so the caller can surface the error to the user.
        throw new Error(
          `NAA sign-in failed (${naaErrCode}). ` +
          "Please ensure you are signed into Outlook with a work or school account and try again."
        );
      }
    }

    // NAA host detected but no token acquired and not interactive — signal gracefully
    if (!interactive) {
      throw new Error("NAA silent auth failed. Interactive sign-in required.");
    }
  }

  // ── TIER 3: MSAL with classic PublicClientApplication ────────────────────
  // This is the existing working flow — silent first, then in-window redirect.
  if (!msalInstance) {
    throw new Error("No authentication method succeeded and no MSAL instance provided.");
  }

  // 3a. Silent
  {
    const active = msalInstance.getActiveAccount();
    const allAccounts = msalInstance.getAllAccounts();
    const account = active ?? allAccounts[0] ?? null;

    if (account) {
      try {
        const silentResult = await msalInstance.acquireTokenSilent({
          ...loginRequest,
          account,
        });
        if (silentResult?.accessToken) {
          console.log("[authManager] ✅ Tier 3a — MSAL silent token acquired.");
          cacheToken(silentResult.accessToken, silentResult.expiresOn);
          return { token: silentResult.accessToken, tier: "msal-silent" };
        }
      } catch (msalSilentErr) {
        console.warn("[authManager] Tier 3a MSAL silent failed:", msalSilentErr.errorCode ?? msalSilentErr.message);
      }
    }
  }

  // 3b. Interactive redirect (in-window — the existing working sign-in flow for Classic Outlook)
  if (interactive) {
    // ❌ SAFETY GUARD: Never redirect inside an iframe.
    // New Outlook runs the taskpane inside an iframe — loginRedirect() is always
    // blocked there (MSAL throws redirect_in_iframe before even trying).
    // If we are in an iframe, NAA should have handled this — surface a clear error.
    const inIframe = typeof window !== "undefined" && window.self !== window.top;
    if (inIframe) {
      console.error(
        "[authManager] ❌ Prevented redirect_in_iframe. " +
        "NAA was not detected on this host but we are inside an iframe. " +
        `NAA 1.1: ${Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.1")}, ` +
        `NAA 1.0: ${Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.0")}`
      );
      remoteLog("error", "Tier 3b: BLOCKED — redirect_in_iframe detected. NAA was NOT initialised but host is an iframe.", {
        naa11: Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.1"),
        naa10: Office?.context?.requirements?.isSetSupported("NestedAppAuth", "1.0"),
        host: Office?.context?.diagnostics?.host,
        platform: Office?.context?.diagnostics?.platform,
      });
      throw new Error(
        "Sign-in is not available inside the New Outlook taskpane via redirect. " +
        "Please ensure you are signed into Outlook with a work or school Microsoft 365 account."
      );
    }

    const active = msalInstance.getActiveAccount();
    const allAccounts = msalInstance.getAllAccounts();
    const account = active ?? allAccounts[0] ?? null;

    console.log("[authManager] Tier 3b — MSAL interactive redirect (in-window sign-in for Classic Outlook)...");

    if (account) {
      await msalInstance.acquireTokenRedirect({
        ...loginRequest,
        account,
        redirectUri: TASKPANE_REDIRECT_URI,
      });
    } else {
      await msalInstance.loginRedirect({
        ...loginRequest,
        redirectUri: TASKPANE_REDIRECT_URI,
      });
    }

    // Redirect takes over navigation — this line is never reached
    throw new Error("Redirecting to Microsoft sign-in...");
  }

  throw new Error("Authentication required. Call getGraphToken({ interactive: true }) to sign in.");
}

/**
 * clearCachedToken — call on sign-out or token invalidation.
 */
export function clearCachedToken() {
  localStorage.removeItem(TOKEN_CACHE_KEY);
  _naaPca = null;
  _naaInitialized = false;
}
