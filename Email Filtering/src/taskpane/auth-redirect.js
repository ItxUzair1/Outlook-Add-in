import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest, DIALOG_REDIRECT_URI } from "./authConfig";

/* global Office */

/**
 * This page is loaded in the Office dialog. It handles the MSAL flow.
 */
Office.onReady(async (info) => {
    if (info.host) {
        console.log("[auth-redirect] Office ready, initializing MSAL...");
        const msalInstance = new PublicClientApplication(msalConfig);

        const postToParent = (payload) => {
            try {
                Office.context.ui.messageParent(JSON.stringify(payload));
            } catch (messageErr) {
                console.error("[auth-redirect] Failed to message parent:", messageErr);
            }
        };

        const redirectRequest = {
            ...loginRequest,
            redirectUri: DIALOG_REDIRECT_URI,
            prompt: "select_account",
        };
        
        try {
            await msalInstance.initialize();
            
            // Check if we are returning from a redirect
            const result = await msalInstance.handleRedirectPromise();
            
            if (result?.accessToken) {
                console.log("[auth-redirect] Access token acquired, messaging parent...");
                postToParent({ 
                    type: "token", 
                    accessToken: result.accessToken, 
                    account: result.account,
                    expiresOn: result.expiresOn ? new Date(result.expiresOn).getTime() : null,
                });
            } else {
                const accounts = msalInstance.getAllAccounts();
                const account = accounts[0] || null;

                // If an account already exists, request token via redirect. Otherwise run interactive login.
                if (account) {
                    msalInstance.setActiveAccount(account);
                    console.log("[auth-redirect] Account found, requesting token by redirect...");
                    await msalInstance.acquireTokenRedirect({
                        ...redirectRequest,
                        account,
                    });
                } else {
                    console.log("[auth-redirect] No account found, starting login redirect...");
                    await msalInstance.loginRedirect(redirectRequest);
                }
            }
        } catch (error) {
            console.error("[auth-redirect] Error in MSAL flow:", error);
            postToParent({ 
                type: "error", 
                message: error.message 
            });
        }
    }
});
