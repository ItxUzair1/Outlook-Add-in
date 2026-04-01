import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import { config } from "../config/index.js";

let cca = null;
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

function getMsalClient() {
  if (cca) return cca;

  if (!config.azureClientId || !config.azureClientSecret) {
    console.warn("[graphService] Microsoft Graph API credentials missing. Graph features will be disabled until Sunday.");
    return null;
  }

  const msalConfig = {
    auth: {
      clientId: config.azureClientId,
      authority: `https://login.microsoftonline.com/${config.azureTenantId || "common"}`,
      clientSecret: config.azureClientSecret,
    }
  };

  cca = new msal.ConfidentialClientApplication(msalConfig);
  return cca;
}

/**
 * Exchanges an Office SSO Identity Token for a Microsoft Graph Access Token
 * using the On-Behalf-Of (OBO) flow.
 */
async function getGraphToken(ssoToken) {
  if (!config.azureClientId || !config.azureClientSecret) {
    throw new Error("Azure credentials missing in backend .env");
  }

  const client = getMsalClient();
  if (!client) {
    throw new Error("Microsoft Graph Service is not configured. Please add the AZURE credentials to your .env file.");
  }

  const oboRequest = {
    oboAssertion: ssoToken,
    scopes: config.graphScopes,
  };

  try {
    console.log("\n================ SSO TOKEN DEBUGGING ================");
    console.log(`[graphService] Received SSO Token from frontend (${ssoToken.length} chars)`);
    const parts = ssoToken.split(".");
    if (parts.length === 3) {
      try {
        const payload = JSON.parse(Buffer.from(parts[1], "base64").toString("utf-8"));
        console.log(`[graphService] Token Tenant ID (tid): ${payload.tid}`);
        console.log(`[graphService] Token Audience  (aud): ${payload.aud}`);
        console.log(`[graphService] Token Issuer    (iss): ${payload.iss}`);
        console.log(`[graphService] User Principal  (upn): ${payload.upn || payload.preferred_username || "N/A"}`);
      } catch (e) {
        console.log("[graphService] Could not decode token payload JSON.");
      }
    } else {
      console.log("[graphService] Token does not appear to be a standard 3-part JWT.");
    }
    console.log(`[graphService] Using MSAL Authority: https://login.microsoftonline.com/${config.azureTenantId || "common"}`);
    
    console.log("[graphService] Attempting Microsoft Graph OBO token exchange...");
    const response = await client.acquireTokenOnBehalfOf(oboRequest);
    console.log("[graphService] Token exchange successful!");
    console.log("=====================================================\n");
    const token = typeof response?.accessToken === "string" ? response.accessToken.trim() : "";
    if (!token) {
      throw new Error("Graph Token Exchange failed: access token is empty.");
    }
    return token;
  } catch (error) {
    console.error("\n================ SSO TOKEN EXCHANGE FAILED ================");
    console.error("[graphService] Token exchange failed:", error.name);
    console.error("[graphService] Message:", error.message);
    if (error.subError) console.error("[graphService] SubError Code:", error.subError);
    console.error("===========================================================\n");
    throw new Error(`Graph Token Exchange failed: ${error.message}`);
  }
}

async function resolveGraphAccessToken(authToken, options = {}) {
  const { isAccessToken = false } = options;
  const normalizedToken = typeof authToken === "string" ? authToken.trim() : authToken;

  if (!normalizedToken) {
    console.warn("[graphService] No authentication token provided in resolveGraphAccessToken.");
    throw new Error("No authentication token was provided for Microsoft Graph.");
  }

  if (isAccessToken) {
    const lower = String(normalizedToken).toLowerCase();
    if (lower === "null" || lower === "undefined" || lower === "[object object]" || lower === "") {
      console.error("[graphService] Direct Graph access token is invalid or empty string.");
      throw new Error("Direct Graph access token is invalid.");
    }
    return normalizedToken;
  }

  console.log(`[graphService] Resolving SSO token (${normalizedToken.length} chars) via OBO flow...`);
  return getGraphToken(normalizedToken);
}

function normalizeItemId(itemId) {
  if (!itemId) return "";
  const strId = String(itemId).trim();
  // If it's already an encoded URL fragment or has no special chars, just use it
  if (strId.includes("%")) return strId;
  return encodeURIComponent(strId);
}

async function runGraphRequest(token, path, options = {}) {
  const url = `${GRAPH_BASE_URL}${path}`;
  
  // Aggressively clean the token: strip all non-printable ASCII and control characters.
  // Standard trim() only handles whitespace, but invisible chars can break headers.
  const rawToken = typeof token === "string" ? token : String(token || "");
  const cleanedToken = rawToken.replace(/[^\x21-\x7E]/g, "").trim();
  
  if (!cleanedToken) {
    const errorMsg = `Graph API request aborted (Rigid-check): Access token is missing or effectively empty. (Original Type: ${typeof token}, Raw Length: ${rawToken.length})`;
    console.error(`[graphService] ${errorMsg}`);
    throw new Error(errorMsg);
  }

  // Debugging invisible characters: log first 5 hex codes if needed.
  const hexDebug = Array.from(cleanedToken.slice(0, 5)).map(c => c.charCodeAt(0).toString(16)).join(" ");
  console.log(`[graphService] Requesting: ${path} (Len: ${cleanedToken.length}, Hex: ${hexDebug}...)`);
  const mergedHeaders = {
    ...(options.headers || {}),
    "Authorization": `Bearer ${cleanedToken}`,
  };
  
  const response = await fetch(url, {
    ...options,
    headers: mergedHeaders,
  });

  if (!response.ok) {
    const err = await response.text();
    console.error(`[graphService] [GS-FAIL-02] Graph API error (${path}):`, response.status, err);
    throw new Error(`Graph API error [GS-FAIL-02] (${path}): ${response.status} - ${err}`);
  }

  return response;
}

async function fetchAttachmentContent(token, itemId, attachmentId) {
  const path = `/me/messages/${normalizeItemId(itemId)}/attachments/${normalizeItemId(attachmentId)}`;
  const response = await runGraphRequest(token, path);
  const attachment = await response.json();
  return attachment.contentBytes || "";
}

/**
 * Fetches the full message metadata and content from Microsoft Graph.
 */
export async function fetchEmailMessage(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(token, `/me/messages/${normalizeItemId(itemId)}`);
  return await response.json();
}

/**
 * Fetches the raw MIME stream of a message for format-preserving conversion.
 */
export async function fetchMimeMessage(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(token, `/me/messages/${normalizeItemId(itemId)}/$value`);
  const buffer = await response.buffer();
  return buffer.toString("base64");
}

/**
 * Fetches all attachments for a specific message.
 */
export async function fetchAttachments(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(token, `/me/messages/${normalizeItemId(itemId)}/attachments`);
  const data = await response.json();
  const attachments = data.value || [];

  console.log(`[graphService] Found ${attachments.length} attachments for message ${itemId}`);

  const enriched = await Promise.all(
    attachments.map(async (att) => {
      let base64Content = att.contentBytes || "";

      if (!base64Content && att.id && att["@odata.type"] === "#microsoft.graph.fileAttachment") {
        try {
          console.log(`[graphService] Fetching content for attachment: ${att.name} (${att.id})`);
          base64Content = await fetchAttachmentContent(token, itemId, att.id);
        } catch (error) {
          console.warn(`[graphService] Failed to fetch content for attachment ${att.name || att.id}:`, error.message);
        }
      }

      return {
        id: att.id,
        name: att.name,
        contentType: att.contentType,
        size: att.size,
        base64Content,
        isInline: att.isInline,
        contentId: att.contentId,
        contentLocation: att.contentLocation,
      };
    })
  );

  return enriched.filter((att) => att.base64Content);
}

/**
 * Moves an email to a target folder (e.g. 'archive' or 'deleteditems').
 */
export async function moveEmail(authToken, itemId, destinationId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(token, `/me/messages/${normalizeItemId(itemId)}/move`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ destinationId })
  });

  return await response.json();
}

export async function markEmailReviewed(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  await runGraphRequest(token, `/me/messages/${normalizeItemId(itemId)}`, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ isRead: true })
  });
  return { success: true };
}

export async function sendFilingLinkEmail(authToken, payload, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const meResponse = await runGraphRequest(token, "/me");
  const me = await meResponse.json();
  const recipient = me.mail || me.userPrincipalName;

  if (!recipient) {
    throw new Error("Could not resolve mailbox recipient for filing link email.");
  }

  const filedEntries = Array.isArray(payload?.filedEntries) ? payload.filedEntries : [];
  const subject = String(payload?.originalSubject || "No Subject");
  const filedAt = payload?.filedAt ? new Date(payload.filedAt).toLocaleString() : new Date().toLocaleString();
  const comment = String(payload?.comment || "").trim();

  const lines = [
    "A message was filed by Mail Manager.",
    "",
    `Original Subject: ${subject}`,
    `Filed At: ${filedAt}`,
  ];

  if (comment) {
    lines.push(`Comment: ${comment}`);
  }

  lines.push("", "Filed Location(s):");
  if (filedEntries.length > 0) {
    for (const entry of filedEntries) {
      lines.push(`- ${entry}`);
    }
  } else {
    lines.push("- (No path available)");
  }

  await runGraphRequest(token, "/me/sendMail", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      message: {
        subject: `Filed link: ${subject}`,
        body: {
          contentType: "Text",
          content: lines.join("\n")
        },
        toRecipients: [
          {
            emailAddress: { address: recipient }
          }
        ]
      },
      saveToSentItems: true
    })
  });

  return { success: true, recipient };
}

export async function archiveEmail(authToken, itemId, options = {}) {
  return moveEmail(authToken, itemId, "archive", options);
}

export async function deleteEmail(authToken, itemId, options = {}) {
  // Business rule: "delete" in this add-in means move to Deleted Items, not hard-delete.
  return moveEmail(authToken, itemId, "deleteditems", options);
}


