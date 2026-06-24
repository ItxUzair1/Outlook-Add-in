// Prevent webpack-dev-server error overlay from capturing minor / third-party unhandled errors in local testing
if (typeof window !== "undefined") {
  window.addEventListener("unhandledrejection", (event) => {
    console.warn("[Global Unhandled Rejection Intercepted]:", event.reason);
    if (event) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  }, true);

  window.addEventListener("error", (event) => {
    console.warn("[Global Error Intercepted]:", event.message);
    if (event) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  }, true);
}

import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { PublicClientApplication, EventType } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";

/* global document, Office, module, require */

const title = "Koyomail";
let mode = null;
if (typeof window !== "undefined") {
  mode = new URLSearchParams(window.location.search).get("mode");
}
try {
  if (!mode && typeof localStorage !== "undefined") {
    mode = localStorage.getItem("koyomail_dialog_mode");
  }
  if (typeof localStorage !== "undefined") {
    localStorage.removeItem("koyomail_dialog_mode");
  }
} catch (e) {
  console.warn("localStorage access restricted:", e);
}
mode = mode || "file";

/**
 * MSAL should be instantiated outside of the component tree to prevent it from being re-instantiated on re-renders.
 * For more, visit: https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-react/docs/getting-started.md
 */
const msalInstance = new PublicClientApplication(msalConfig);

const ACTIVE_ACCOUNT_KEY = "koyomail_activeAccountId";

msalInstance.addEventCallback((event) => {
  if (event.eventType === EventType.LOGIN_SUCCESS && event.payload.account) {
    const account = event.payload.account;
    msalInstance.setActiveAccount(account);
    // Persist for Classic Outlook WebView restarts
    try { localStorage.setItem(ACTIVE_ACCOUNT_KEY, account.homeAccountId); } catch { /* ignore */ }
  }
});

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(async () => {
  try {
    await msalInstance.initialize();

    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult?.account) {
      msalInstance.setActiveAccount(redirectResult.account);
      try { localStorage.setItem(ACTIVE_ACCOUNT_KEY, redirectResult.account.homeAccountId); } catch { /* ignore */ }
    }
    
    // Recover the previously-active account after Classic Outlook WebView restarts
    if (!msalInstance.getActiveAccount() && msalInstance.getAllAccounts().length > 0) {
      const savedId = localStorage.getItem(ACTIVE_ACCOUNT_KEY);
      const match = savedId
        ? msalInstance.getAllAccounts().find(a => a.homeAccountId === savedId)
        : null;
      msalInstance.setActiveAccount(match || msalInstance.getAllAccounts()[0]);
    }

    root?.render(
      <MsalProvider instance={msalInstance}>
        <FluentProvider theme={webLightTheme}>
          <App title={title} initialMode={mode} />
        </FluentProvider>
      </MsalProvider>
    );
  } catch (error) {
    console.error("MSAL Initialization failed:", error);
    
    // Create a mock MSAL instance to prevent rendering crashes
    const mockMsalInstance = {
      initialize: () => Promise.resolve(),
      handleRedirectPromise: () => Promise.resolve(null),
      getAllAccounts: () => [],
      getActiveAccount: () => null,
      setActiveAccount: () => {},
      addEventCallback: () => {},
      acquireTokenSilent: () => Promise.reject(new Error("MSAL not initialized")),
      acquireTokenRedirect: () => Promise.reject(new Error("MSAL not initialized")),
      loginRedirect: () => Promise.reject(new Error("MSAL not initialized")),
    };

    root?.render(
       <MsalProvider instance={mockMsalInstance}>
         <FluentProvider theme={webLightTheme}>
           <App title={title} initialMode={mode} msalError={error?.message || String(error)} />
         </FluentProvider>
       </MsalProvider>
    );
  }
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(
      <MsalProvider instance={msalInstance}>
        <FluentProvider theme={webLightTheme}>
          <NextApp title={title} initialMode={mode} />
        </FluentProvider>
      </MsalProvider>
    );
  });
}
