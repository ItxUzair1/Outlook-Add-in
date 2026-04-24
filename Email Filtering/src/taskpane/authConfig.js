/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { LogLevel } from "@azure/msal-browser";

const getOrigin = () => {
    if (typeof window !== "undefined" && window.location?.origin) {
        return window.location.origin;
    }
    return "https://localhost:3000";
};

export const TASKPANE_REDIRECT_URI = `${getOrigin()}/taskpane.html`;
export const DIALOG_REDIRECT_URI = `${getOrigin()}/auth-redirect.html`;

export const getRedirectFallbackUris = () => {
    const origin = getOrigin();
    return [
        `${origin}/taskpane.html`,
        origin,
        `${origin}/auth-redirect.html`,
    ];
};

/**
 * Configuration object to be passed to MSAL instance on creation. 
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md 
 */
export const msalConfig = {
    auth: {
        clientId: "3860f34f-e563-42e6-a9d6-7022d0cd5632",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: TASKPANE_REDIRECT_URI,
        navigateToLoginRequestUrl: false, // Prevent top-frame navigation issues
    },
    cache: {
        cacheLocation: "localStorage", // localStorage is often more reliable than sessionStorage in taskpanes
        storeAuthStateInCookie: true, // Crucial for IE11/Edge Legacy and some Desktop environments
    },
    system: {
        loadFrameTimeout: 30000, // Slow WebView hosts often need more time.
        iframeHashTimeout: 30000,
        windowHashTimeout: 60000,
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) {
                    return;
                }
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        console.warn(message);
                        return;
                    default:
                        return;
                }
            }
        }
    }
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
/**
 * Scopes that match the Azure App Registration delegated permissions.
 * All scopes below are granted for Default Directory (confirmed in Azure portal).
 */
export const loginRequest = {
    scopes: ["User.Read", "Mail.Read", "Mail.ReadWrite", "Mail.Send", "email", "offline_access"]
};

/**
 * NAA-specific scope request — uses fully-qualified Graph resource URIs.
 * The NAA broker requires scopes in this format:
 *   https://graph.microsoft.com/  (resource) + scope
 * Using short-form scopes (User.Read) may fail with the broker.
 * All scopes confirmed granted in Azure portal.
 */
export const naaLoginRequest = {
    scopes: [
        "https://graph.microsoft.com/User.Read",
        "https://graph.microsoft.com/Mail.Read",
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/Mail.Send",
        "email",
    ]
};

/**
 * Add here the endpoints and scopes for "Accessing a Web API" on which you would like to obtain an access token.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md 
 */
export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me"
};

/**
 * Config for the NAA (Nested App Authentication) MSAL instance.
 * Used ONLY by authManager.js Tier 2.
 * Uses createNestablePublicClientApplication — NOT PublicClientApplication.
 *
 * Key differences from msalConfig:
 *  - No redirectUri (broker handles auth at OS level, no redirect needed)
 *  - sessionStorage (NAA tokens are session-scoped)
 *  - storeAuthStateInCookie: false (not needed for NAA)
 */
export const msalNaaConfig = {
    auth: {
        clientId: "3860f34f-e563-42e6-a9d6-7022d0cd5632",
        authority: "https://login.microsoftonline.com/common",
        // ✅ Do NOT set redirectUri — the broker (Outlook) handles auth natively
        supportsNestedAppAuth: true,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) return;
                switch (level) {
                    case LogLevel.Error:   console.error(`[MSAL-NAA] ${message}`); return;
                    case LogLevel.Warning: console.warn(`[MSAL-NAA] ${message}`);  return;
                    case LogLevel.Info:    console.info(`[MSAL-NAA] ${message}`);  return;
                    case LogLevel.Verbose: console.debug(`[MSAL-NAA] ${message}`); return;
                    default: return;
                }
            },
        },
    },
};
