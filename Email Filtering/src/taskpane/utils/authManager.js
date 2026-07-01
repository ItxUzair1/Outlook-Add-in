/**
 * authManager.js — Unified three-tier authentication for Koyomail
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
import { openAuthDialogAndGetToken } from "./dialogAuth.js";

const SSO_TIMEOUT_MS = 8000;
const NAA_INIT_TIMEOUT_MS = 10000;

function withTimeout(promise, timeoutMs, timeoutMessage) {
  let timer;
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(timeoutMessage)), timeoutMs);
    }),
  ]).finally(() => {
    if (timer) clearTimeout(timer);
  });
}

function getAuthRedirectDialogUrl() {
  if (typeof window !== "undefined" && window.location?.origin) {
    return `${window.location.origin}/auth-redirect.html`;
  }
  return "/auth-redirect.html";
}

/** New Outlook and filing dialogs run inside an iframe where Office SSO is slow/unreliable. */
export function isOutlookIframeHost() {
  return typeof window !== "undefined" && window.self !== window.top;
}

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
    });

    const inIframe = typeof window !== "undefined" ? window.self !== window.top : false;
    
    // We try to initialize even if `isSetSupported` is false.
    // Why? Because XML sideloading in New Outlook sometimes falsely reports false 
    // for the requirement set. However, New Outlook ALWAYS runs the taskpane in an iframe.
    // If not in an iframe and not explicitly supported, we are in Classic Desktop.
    if (!supportedHint && !inIframe) {
        console.log("[authManager] Classic Outlook detected (no iframe, no hint). Skipping NAA to prevent WAM popup errors.");
        remoteLog("info", "Skipping NAA — Classic Outlook detected.");
        return null;
    }

    _naaPca = await withTimeout(
      createNestablePublicClientApplication(msalNaaConfig),
      NAA_INIT_TIMEOUT_MS,
      "NAA client initialization timed out"
    );
    
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
const TOKEN_CACHE_KEY = "koyomailGraphTokenV2";

function readCachedToken() {
  try {
    const raw = localStorage.getItem(TOKEN_CACHE_KEY);
    if (!raw) return null;
    const { accessToken, expiresOn, tier } = JSON.parse(raw);
    // Expire 2 minutes early to avoid edge-case expiry during a request
    if (!accessToken || !expiresOn || Date.now() >= expiresOn - 120_000) {
      localStorage.removeItem(TOKEN_CACHE_KEY);
      return null;
    }
    return { token: accessToken, tier: tier || "cache" };
  } catch {
    localStorage.removeItem(TOKEN_CACHE_KEY);
    return null;
  }
}

function cacheToken(accessToken, expiresOn, tier = "unknown") {
  if (!accessToken) return;
  // SSO identity tokens expire in ~1 hour; cap their cache at 50 minutes to
  // prevent stale SSO tokens being served long after they have expired.
  // NAA/MSAL Graph tokens can be cached for up to 1 hour (per MSAL policy).
  const SSO_MAX_TTL = 50 * 60 * 1000; // 50 minutes
  const DEFAULT_TTL = 60 * 60 * 1000; // 1 hour
  const fallback = Date.now() + (tier === "sso" ? SSO_MAX_TTL : DEFAULT_TTL);
  const ts = expiresOn
    ? Number(expiresOn instanceof Date ? expiresOn.getTime() : expiresOn)
    : fallback;
  // For SSO tokens, never cache beyond 50 minutes regardless of the supplied expiry
  const effectiveExpiry = tier === "sso" ? Math.min(ts, Date.now() + SSO_MAX_TTL) : ts;
  localStorage.setItem(
    TOKEN_CACHE_KEY,
    JSON.stringify({ accessToken, expiresOn: Number.isFinite(effectiveExpiry) ? effectiveExpiry : fallback, tier })
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
    console.log(`[authManager] ✅ Tier 0 — returning cached token (original tier: ${cached.tier}).`);
    remoteLog("ok", `Tier 0: Token served from cache (original tier: ${cached.tier})`);
    return { token: cached.token, tier: cached.tier };
  }

  remoteLog("info", "Auth flow started", { interactive, hasLoginHint: !!loginHint });

  const inIframeHost = isOutlookIframeHost();
  const isDialog = typeof Office !== "undefined" && Office?.context?.ui?.messageParent;

  // ── TIER 1: Office SSO (Classic desktop only — unreliable in iframe / New Outlook) ──
  if (!inIframeHost && !isDialog) {
  remoteLog("info", "Tier 1: Attempting Office SSO (getAccessToken)...");
  try {
    if (typeof Office !== "undefined" && Office?.auth?.getAccessToken) {
      const requestSsoToken = (options) =>
        withTimeout(
          new Promise((resolve, reject) => {
            Office.auth.getAccessToken(options, (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
              } else {
                const code = result.error?.code ?? "Unknown";
                const msg = result.error?.message ?? "No error message provided by Office";
                reject(new Error(`SSO Token Failed: ${msg} (Code: ${code})`));
              }
            });
          }),
          SSO_TIMEOUT_MS,
          "SSO Token Timeout"
        );

      let ssoToken = null;
      try {
        ssoToken = await requestSsoToken({
          allowSignInPrompt: interactive,
          allowConsentPrompt: interactive,
          forMSGraphAccess: true,
        });
      } catch (primarySsoErr) {
        const msg = String(primarySsoErr?.message || "").toLowerCase();
        const shouldRetryWithoutGraphHint =
          msg.includes("code: 7000") ||
          msg.includes("permission denied") ||
          msg.includes("sufficient permissions");

        if (shouldRetryWithoutGraphHint) {
          console.warn("[authManager] SSO with forMSGraphAccess failed; retrying without forMSGraphAccess.");
          ssoToken = await requestSsoToken({
            allowSignInPrompt: interactive,
            allowConsentPrompt: interactive,
          });
        } else {
          throw primarySsoErr;
        }
      }

      if (ssoToken) {
        console.log("[authManager] ✅ Tier 1 — SSO token acquired.");
        remoteLog("ok", "Tier 1: SSO token acquired ✅");
        cacheToken(ssoToken, null, "sso"); // SSO tokens: 50-minute cap applied inside cacheToken
        return { token: ssoToken, tier: "sso" };
      }
    }
  } catch (ssoErr) {
    const code = ssoErr?.code ?? ssoErr?.errorCode ?? "";
    console.warn(`[authManager] Tier 1 SSO failed (code ${code}):`, ssoErr.message ?? ssoErr);
    remoteLog("warn", `Tier 1: SSO FAILED — code=${code} message=${ssoErr.message ?? ssoErr}`);
  }
  } else {
    console.log("[authManager] Tier 1 SSO skipped — iframe host or dialog. Using NAA/MSAL.");
    remoteLog("info", "Tier 1: SSO skipped for iframe host or dialog");
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

      const silentResult = await withTimeout(
        naaPca.acquireTokenSilent({
          ...naaLoginRequest,
          account,
          loginHint: hint,
        }),
        15000,
        "NAA silent token request timed out"
      );

      if (silentResult?.accessToken) {
        console.log("[authManager] ✅ Tier 2a — NAA silent token acquired.");
        remoteLog("ok", "Tier 2a: NAA silent token acquired ✅");
        cacheToken(silentResult.accessToken, silentResult.expiresOn, "naa-silent");
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
        const popupResult = await withTimeout(
          naaPca.acquireTokenPopup({
            ...naaLoginRequest,
            loginHint: hint,
          }),
          120000,
          "NAA interactive sign-in timed out"
        );

        if (popupResult?.accessToken) {
          console.log("[authManager] ✅ Tier 2b — NAA interactive token acquired.");
          remoteLog("ok", "Tier 2b: NAA interactive token acquired ✅");
          cacheToken(popupResult.accessToken, popupResult.expiresOn, "naa-interactive");
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

        const inIframeHost = typeof window !== "undefined" && window.self !== window.top;
        if (!inIframeHost) {
          throw new Error(
            `NAA sign-in failed (${naaErrCode}). ` +
            "Please ensure you are signed into Outlook with a work or school account and try again."
          );
        }

        console.warn("[authManager] NAA interactive failed in iframe — falling through to Office auth dialog.");
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

  const inIframe = typeof window !== "undefined" && window.self !== window.top;

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
          loginHint: loginHint || account.username,
          forceRefresh: false,
        });
        if (silentResult?.accessToken) {
          console.log("[authManager] ✅ Tier 3a — MSAL silent token acquired.");
          cacheToken(silentResult.accessToken, silentResult.expiresOn, "msal-silent");
          return { token: silentResult.accessToken, tier: "msal-silent" };
        }
      } catch (msalSilentErr) {
        console.warn("[authManager] Tier 3a MSAL silent failed:", msalSilentErr.errorCode ?? msalSilentErr.message);
      }
    }
  }

  // 3b. Interactive sign-in
  if (interactive) {
    // Filing/taskpane dialogs run in an iframe — MSAL redirect is blocked there.
    // Use a top-level Office auth dialog (displayInIframe: false) instead.
    if (inIframe) {
      if (Office?.context?.ui?.displayDialogAsync) {
        try {
          console.log("[authManager] Tier 3b — opening Office auth dialog (iframe host)...");
          remoteLog("info", "Tier 3b: Opening Office auth dialog for iframe host");
          const dialogResult = await openAuthDialogAndGetToken(getAuthRedirectDialogUrl());
          if (dialogResult?.accessToken) {
            if (dialogResult.account && msalInstance?.setActiveAccount) {
              msalInstance.setActiveAccount(dialogResult.account);
            }
            cacheToken(dialogResult.accessToken, dialogResult.expiresOn, "msal-dialog");
            return { token: dialogResult.accessToken, tier: "msal-dialog" };
          }
        } catch (dialogErr) {
          console.warn("[authManager] Office auth dialog failed:", dialogErr.message);
          remoteLog("error", `Tier 3b: Office auth dialog FAILED — ${dialogErr.message}`);
          throw new Error(
            dialogErr.message ||
            "Sign-in dialog could not be completed. Please try again."
          );
        }
      }

      console.error(
        "[authManager] Interactive sign-in blocked in iframe and Office Dialog API is unavailable."
      );
      throw new Error(
        "Sign-in is not available in this Outlook window. " +
        "Please ensure you are signed into Outlook with a work or school Microsoft 365 account and try again."
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
