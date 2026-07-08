import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import { config } from "../config/index.js";

let cca = null;
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

class AsyncQueue {
  constructor(concurrency) {
    this.concurrency = concurrency;
    this.running = 0;
    this.queue = [];
  }
  async enqueue(task) {
    if (this.running >= this.concurrency) {
      await new Promise(resolve => this.queue.push(resolve));
    }
    this.running++;
    try {
      return await task();
    } finally {
      this.running--;
      if (this.queue.length > 0) {
        const next = this.queue.shift();
        next();
      }
    }
  }
}

// Cap concurrent Microsoft Graph outbound requests to 4 globally.
// This prevents bursts that trigger 15-second HTTP 429 rate limit penalties.
const graphQueue = new AsyncQueue(4);

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

// Cache OBO exchanges — without this, every Graph call re-exchanges the SSO token (~1-3s each).
const oboTokenCache = new Map();
const OBO_CACHE_TTL_MS = 50 * 60 * 1000;
const masterCategoryListCache = new Map();
const MASTER_CATEGORY_CACHE_TTL_MS = 5 * 60 * 1000;

function getTokenCacheKey(token) {
  return String(token || "").slice(0, 96);
}

function readOboCache(ssoToken) {
  const key = getTokenCacheKey(ssoToken);
  const entry = oboTokenCache.get(key);
  if (!entry || Date.now() >= entry.expiresAt) {
    if (entry) oboTokenCache.delete(key);
    return null;
  }
  return entry.token;
}

function writeOboCache(ssoToken, accessToken) {
  oboTokenCache.set(getTokenCacheKey(ssoToken), {
    token: accessToken,
    expiresAt: Date.now() + OBO_CACHE_TTL_MS,
  });
}

function isTransientNetworkError(err) {
  const msg = String(err?.message || err?.cause?.message || err || "").toLowerCase();
  return (
    msg.includes("econnreset") ||
    msg.includes("etimedout") ||
    msg.includes("eai_again") ||
    msg.includes("socket hang up") ||
    msg.includes("network request failed")
  );
}

function isAccessDeniedError(err) {
  const msg = String(err?.message || err || "").toLowerCase();
  return msg.includes("403") || msg.includes("erroraccessdenied") || msg.includes("access is denied");
}

function getMailboxPrefix(options = {}) {
  const mailbox = options.delegateMailbox || options.sharedMailbox;
  if (mailbox) {
    return `/users/${encodeURIComponent(mailbox.trim())}`;
  }
  return "/me";
}

async function ensureMasterCategoryOnGraph(token, categoryName, color = "preset19", options = {}) {
  const mailboxSuffix = options.delegateMailbox || options.sharedMailbox || "";
  const cacheKey = getTokenCacheKey(token) + ":" + mailboxSuffix;
  const cached = masterCategoryListCache.get(cacheKey);
  if (cached?.accessDenied && Date.now() < cached.expiresAt) {
    return;
  }

  const prefix = getMailboxPrefix(options);

  try {
    let masterCategories = null;
    if (cached && Date.now() < cached.expiresAt && !cached.accessDenied) {
      masterCategories = cached.value;
    } else {
      const catResp = await runGraphRequest(token, `${prefix}/outlook/masterCategories`);
      const catData = await catResp.json();
      masterCategories = Array.isArray(catData?.value) ? catData.value : [];
      masterCategoryListCache.set(cacheKey, {
        value: masterCategories,
        expiresAt: Date.now() + MASTER_CATEGORY_CACHE_TTL_MS,
      });
    }

    const existingCat = masterCategories.find((c) => c.displayName === categoryName);
    if (!existingCat) {
      await runGraphRequest(token, `${prefix}/outlook/masterCategories`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ displayName: categoryName, color }),
      });
      masterCategories.push({ displayName: categoryName, color });
      masterCategoryListCache.set(cacheKey, {
        value: masterCategories,
        expiresAt: Date.now() + MASTER_CATEGORY_CACHE_TTL_MS,
      });
    } else if (existingCat.color !== color && existingCat.id) {
      await runGraphRequest(token, `${prefix}/outlook/masterCategories/${existingCat.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ color }),
      });
    }
  } catch (err) {
    if (isAccessDeniedError(err)) {
      console.warn("[graphService] Master category list unavailable (403) — skipping. Will still apply category on the message.");
      masterCategoryListCache.set(cacheKey, {
        value: [],
        accessDenied: true,
        expiresAt: Date.now() + MASTER_CATEGORY_CACHE_TTL_MS,
      });
      return;
    }
    throw err;
  }
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

export async function resolveGraphAccessToken(authToken, options = {}) {
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
  const cached = readOboCache(normalizedToken);
  if (cached) return cached;
  const accessToken = await getGraphToken(normalizedToken);
  writeOboCache(normalizedToken, accessToken);
  return accessToken;
}

function normalizeItemId(itemId) {
  if (!itemId) return "";
  const strId = String(itemId).trim();
  // Replace forward slashes with hyphens to prevent Microsoft Graph's routing engine
  // from incorrectly parsing percent-encoded slashes (%2F) as segment separators.
  const safeId = strId.replace(/\//g, "-");
  if (safeId.includes("%")) return safeId;
  return encodeURIComponent(safeId);
}

async function runGraphRequest(token, path, options = {}, retryCount = 0) {
  const MAX_RETRIES = 3;
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
  const mergedHeaders = {
    ...(options.headers || {}),
    "Authorization": `Bearer ${cleanedToken}`,
  };

  const doFetch = async () => graphQueue.enqueue(() => fetch(url, {
    ...options,
    headers: mergedHeaders,
  }));

  let response;
  try {
    response = await doFetch();
  } catch (fetchErr) {
    if (retryCount < MAX_RETRIES && isTransientNetworkError(fetchErr)) {
      const waitMs = (retryCount + 1) * 1000;
      console.warn(`[graphService] Transient network error on ${path}: ${fetchErr.message}. Retrying in ${waitMs}ms (${retryCount + 1}/${MAX_RETRIES})...`);
      await new Promise((r) => setTimeout(r, waitMs));
      return runGraphRequest(token, path, options, retryCount + 1);
    }
    throw fetchErr;
  }

  // Handle 429 Too Many Requests (Microsoft Graph rate limiting).
  // Read the Retry-After header (in seconds) and wait before retrying.
  if (response.status === 429) {
    if (retryCount >= MAX_RETRIES) {
      const err = await response.text();
      console.error(`[graphService] [GS-FAIL-429] Rate limit exceeded after ${MAX_RETRIES} retries (${path}):`, err);
      throw new Error(`Graph API rate limit exceeded [GS-FAIL-429] (${path}): 429 - ${err}`);
    }
    const retryAfterHeader = response.headers.get("Retry-After");
    const retryAfterSec = parseInt(retryAfterHeader || "5", 10);
    const waitMs = (retryAfterSec + 1) * 1000;
    console.warn(`[graphService] Rate limited by Microsoft Graph (429). Retry-After: ${retryAfterSec}s. Waiting ${waitMs}ms before retry ${retryCount + 1}/${MAX_RETRIES}... (${path})`);
    await new Promise(r => setTimeout(r, waitMs));
    return runGraphRequest(token, path, options, retryCount + 1);
  }

  if (!response.ok) {
    const err = await response.text();
    console.error(`[graphService] [GS-FAIL-02] Graph API error (${path}):`, response.status, err);
    throw new Error(`Graph API error [GS-FAIL-02] (${path}): ${response.status} - ${err}`);
  }

  return response;
}

async function fetchAttachmentContent(token, itemId, attachmentId, options = {}) {
  const path = `${getMailboxPrefix(options)}/messages/${normalizeItemId(itemId)}/attachments/${normalizeItemId(attachmentId)}`;
  const response = await runGraphRequest(token, path);
  const attachment = await response.json();
  return attachment.contentBytes || "";
}

/**
 * Fetches the full message metadata and content from Microsoft Graph.
 */
export async function fetchEmailMessage(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const select = options.select;
  const prefix = getMailboxPrefix(options);
  const path = select
    ? `${prefix}/messages/${normalizeItemId(itemId)}?$select=${encodeURIComponent(select)}`
    : `${prefix}/messages/${normalizeItemId(itemId)}`;
  const response = await runGraphRequest(token, path);
  return await response.json();
}

/**
 * Lightweight Graph lookup used to verify an item ID before post-filing actions.
 */
export async function verifyGraphMessageId(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(
    token,
    `${getMailboxPrefix(options)}/messages/${normalizeItemId(itemId)}?$select=id,subject,hasAttachments`
  );
  return await response.json();
}

/**
 * Converts Exchange (EWS) item IDs to REST IDs understood by Microsoft Graph.
 */
export async function translateExchangeIds(authToken, inputIds, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(token, `${getMailboxPrefix(options)}/translateExchangeIds`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      inputIds: inputIds,
      sourceIdType: "ewsId",
      targetIdType: "restId",
    }),
  });
  const data = await response.json();
  const first = Array.isArray(data?.value) ? data.value[0] : null;
  return first?.targetId || null;
}

/**
 * Fetches the raw MIME stream of a message for format-preserving conversion.
 */
export async function fetchMimeMessage(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(token, `${getMailboxPrefix(options)}/messages/${normalizeItemId(itemId)}/$value`);
  const buffer = await response.buffer();
  return buffer.toString("base64");
}

/**
 * Fetches all attachments for a specific message.
 */
export async function fetchAttachments(authToken, itemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const prefix = getMailboxPrefix(options);
  const response = await runGraphRequest(token, `${prefix}/messages/${normalizeItemId(itemId)}/attachments`);
  const data = await response.json();
  const attachments = data.value || [];

  console.log(`[graphService] Found ${attachments.length} attachments for message ${itemId}`);

  const enriched = await Promise.all(
    attachments.map(async (att) => {
      let base64Content = att.contentBytes || "";

      if (!base64Content && att.id && att["@odata.type"] === "#microsoft.graph.fileAttachment") {
        try {
          console.log(`[graphService] Fetching content for attachment: ${att.name} (${att.id})`);
          base64Content = await fetchAttachmentContent(token, itemId, att.id, options);
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
  const response = await runGraphRequest(token, `${getMailboxPrefix(options)}/messages/${normalizeItemId(itemId)}/move`, {
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
  await runGraphRequest(token, `${getMailboxPrefix(options)}/messages/${normalizeItemId(itemId)}`, {
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
  const prefix = getMailboxPrefix(options);
  const meResponse = await runGraphRequest(token, prefix);
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
    "A message was filed by Koyomail.",
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

  await runGraphRequest(token, `${prefix}/sendMail`, {
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

/**
 * Creates a DRAFT email in the user's mailbox containing the filing links.
 * The draft appears in Outlook's Drafts folder, ready for the user to add
 * recipients, edit, and send manually.
 */
export async function createDraftLinkEmail(authToken, payload, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);

  // Escape HTML special characters to prevent XSS in email body
  const escapeHtml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  const filedEntries = Array.isArray(payload?.filedEntries) ? payload.filedEntries : [];
  const subject = String(payload?.originalSubject || "No Subject");
  const comment = String(payload?.comment || "").trim();
  const fontFamily = payload?.fontFamily || "Segoe UI";
  const fontSize = payload?.fontSize || "11pt";

  // Build clickable file:/// links for each filed path
  const linkItems = filedEntries.map(entry => {
    let fileUrl = entry;
    if (entry.startsWith("\\\\")) {
        // UNC path: e.g. \\localhost\C$ -> file://localhost/C$
        fileUrl = `file://${entry.substring(2).replace(/\\/g, "/")}`;
    } else {
        // Local path: e.g. C:\folder -> file:///C:/folder
        fileUrl = `file:///${entry.replace(/\\/g, "/")}`;
    }
    return `<li style="margin-bottom: 6px;"><a href="${escapeHtml(fileUrl)}" style="color: #0078d4; text-decoration: none;">${escapeHtml(entry)}</a></li>`;
  }).join("");

  const commentBlock = comment
    ? `<p style="margin: 8px 0;"><strong>Comment:</strong> ${escapeHtml(comment)}</p>`
    : "";

  const htmlBody = `
    <div style="font-family: '${escapeHtml(fontFamily)}', sans-serif; font-size: ${escapeHtml(fontSize)}; color: #323130;">
      <p>The following email has been filed to a shared location:</p>
      <p><strong>Subject:</strong> ${escapeHtml(subject)}</p>
      ${commentBlock}
      <p><strong>Filed Location(s):</strong></p>
      <ul style="list-style: none; padding-left: 0;">
        ${linkItems}
      </ul>
      <p style="color: #888; font-size: 9pt; margin-top: 16px;"><em>Generated by Koyomail</em></p>
    </div>
  `;

  const response = await runGraphRequest(token, `${getMailboxPrefix(options)}/messages`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      subject: `Filed Link: ${subject}`,
      body: {
        contentType: "HTML",
        content: htmlBody
      },
      // No toRecipients — user fills them in
      toRecipients: [],
      isDraft: true
    })
  });

  const draft = await response.json();
  console.log(`[graphService] Draft email created: ${draft.id}`);
  return { success: true, draftId: draft.id, webLink: draft.webLink };
}

export async function archiveEmail(authToken, itemId, options = {}) {
  return moveEmail(authToken, itemId, "archive", options);
}

export async function deleteEmail(authToken, itemId, options = {}) {
  // Business rule: "delete" in this add-in means move to Deleted Items, not hard-delete.
  return moveEmail(authToken, itemId, "deleteditems", options);
}

/**
 * Adds a category label to an email via Microsoft Graph.
 */
export async function addCategoryToEmail(authToken, itemId, categoryName, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const { skipMasterCategoryEnsure = false } = options;

  if (!skipMasterCategoryEnsure) {
    try {
      await ensureMasterCategoryOnGraph(token, categoryName, "preset19", options);
    } catch (err) {
      console.warn("[graphService] Failed to ensure master category:", err.message);
    }
  }

  const prefix = getMailboxPrefix(options);
  const getResp = await runGraphRequest(token, `${prefix}/messages/${normalizeItemId(itemId)}?$select=categories`);
  const msgData = await getResp.json();
  const existing = Array.isArray(msgData.categories) ? msgData.categories : [];
  if (existing.includes(categoryName)) {
    return { success: true, alreadyPresent: true };
  }
  const response = await runGraphRequest(token, `${prefix}/messages/${normalizeItemId(itemId)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ categories: [...existing, categoryName] }),
  });
  const updatedData = await response.json();
  return { success: true, newId: updatedData.id || itemId };
}

/**
 * Applies all post-filing Graph actions in the minimum number of API calls.
 * Typical flow: 1 GET + 1 PATCH + (optional) 1 move = 2-3 calls instead of 8-12.
 */
export async function applyPostFilingBatch(authToken, itemId, actions, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const graphOpts = { isAccessToken: true };
  const {
    markReviewed = false,
    addFiledCategory = false,
    filedCategoryName = "Filed",
    assistantCategories = "",
    afterFiling = "none",
    useUtc = false,
    filedFolderPrefix = "*",
    targetFolderName = "Filed",
    deleteEmptyFolders = false,
    fallbackSubject = "",
    skipMasterCategoryEnsure = false,
  } = actions;

  const extraCats = String(assistantCategories || "").split(",").map((c) => c.trim()).filter(Boolean);
  const needsCategory = addFiledCategory || extraCats.length > 0;
  const needsSubjectChange = markReviewed || afterFiling === "add_date";
  const needsMove = afterFiling && afterFiling !== "none" && afterFiling !== "add_date";

  let resolvedItemId = itemId;
  let msgData = { subject: fallbackSubject, categories: [], isRead: false };

  if (needsCategory || needsSubjectChange) {
    const response = await runGraphRequest(
      token,
      `/me/messages/${normalizeItemId(resolvedItemId)}?$select=id,subject,categories,isRead`
    );
    msgData = await response.json();
    if (msgData?.id) resolvedItemId = msgData.id;
  }

  const patch = {};
  if (markReviewed && !msgData.isRead) {
    patch.isRead = true;
  }

  let subject = msgData.subject || fallbackSubject || "";
  if (markReviewed && !subject.startsWith("[Reviewed]")) {
    subject = `[Reviewed] ${subject}`;
  }
  if (afterFiling === "add_date") {
    const dateStr = useUtc
      ? `${new Date().toISOString().replace("T", " ").substring(0, 19)} UTC`
      : new Date().toLocaleString();
    const prefix = `[Filed ${dateStr}] `;
    if (!subject.startsWith(prefix.trim())) {
      subject = `${prefix}${subject}`;
    }
  }
  if (subject !== msgData.subject) {
    patch.subject = subject;
  }

  if (needsCategory) {
    const cats = Array.isArray(msgData.categories) ? [...msgData.categories] : [];
    if (addFiledCategory && !cats.includes(filedCategoryName)) {
      if (!skipMasterCategoryEnsure) {
        try {
          await ensureMasterCategoryOnGraph(token, filedCategoryName, "preset19");
        } catch (err) {
          console.warn(`[graphService] Master category ensure failed for "${filedCategoryName}":`, err.message);
        }
      }
      cats.push(filedCategoryName);
    }
    for (const cat of extraCats) {
      if (!cats.includes(cat)) {
        if (!skipMasterCategoryEnsure) {
          try {
            await ensureMasterCategoryOnGraph(token, cat);
          } catch (err) {
            console.warn(`[graphService] Master category ensure failed for "${cat}":`, err.message);
          }
        }
        cats.push(cat);
      }
    }
    if (JSON.stringify(cats) !== JSON.stringify(msgData.categories || [])) {
      patch.categories = cats;
    }
  }

  if (Object.keys(patch).length > 0) {
    await runGraphRequest(token, `/me/messages/${normalizeItemId(resolvedItemId)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(patch),
    });
  }

  if (needsMove) {
    if (afterFiling === "delete" || afterFiling === "move_deleted") {
      await moveEmail(token, resolvedItemId, "deleteditems", graphOpts);
    } else if (afterFiling === "archive") {
      await moveEmail(token, resolvedItemId, "archive", graphOpts);
    } else if (afterFiling === "move_filed_items") {
      const folderId = await getOrCreateMailFolder(token, "inbox", "Filed Items", graphOpts);
      await moveEmail(token, resolvedItemId, folderId, graphOpts);
    } else if (afterFiling === "move_filed_folders") {
      const folderName = `${filedFolderPrefix} ${targetFolderName}`.trim();
      const folderId = await getOrCreateMailFolder(token, "inbox", folderName, graphOpts);
      await moveEmail(token, resolvedItemId, folderId, graphOpts);
      if (deleteEmptyFolders) {
        await cleanupEmptyFolders(token, "inbox", filedFolderPrefix, graphOpts);
      }
    }
  }

  return { success: true, itemId: resolvedItemId };
}

/**
 * Updates the subject of an email (e.g. to prepend filed date).
 */
export async function updateEmailSubject(authToken, itemId, newSubject, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const response = await runGraphRequest(token, `${getMailboxPrefix(options)}/messages/${normalizeItemId(itemId)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ subject: newSubject })
  });
  const updatedData = await response.json();
  return { success: true, newId: updatedData.id || itemId };
}

/**
 * Creates a sub-folder under a parent folder (e.g. Inbox).
 * Returns the folder id. If the folder already exists, returns its id.
 */
export async function getOrCreateMailFolder(authToken, parentFolderId, folderName, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const prefix = getMailboxPrefix(options);
  // Check if folder already exists
  const listResp = await runGraphRequest(token, `${prefix}/mailFolders/${parentFolderId}/childFolders?$filter=displayName eq '${folderName.replace(/'/g, "''")}'`);
  const listData = await listResp.json();
  if (listData.value && listData.value.length > 0) {
    return listData.value[0].id;
  }
  // Create the folder
  const createResp = await runGraphRequest(token, `${prefix}/mailFolders/${parentFolderId}/childFolders`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: folderName })
  });
  const created = await createResp.json();
  return created.id;
}

/**
 * Scans the user's child folders under a given parent (e.g., 'inbox') and deletes
 * any that match a specific prefix and have no items.
 */
export async function cleanupEmptyFolders(authToken, parentFolderId, prefix, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  const mboxPrefix = getMailboxPrefix(options);
  try {
    const listResp = await runGraphRequest(token, `${mboxPrefix}/mailFolders/${parentFolderId}/childFolders`);
    const listData = await listResp.json();
    if (!listData.value) return;

    for (const folder of listData.value) {
      if (prefix && folder.displayName.startsWith(prefix) && folder.totalItemCount === 0) {
        try {
          await runGraphRequest(token, `${mboxPrefix}/mailFolders/${folder.id}`, { method: 'DELETE' });
        } catch (err) {
          console.warn(`[graphService] Error deleting empty folder ${folder.displayName}:`, err.message);
        }
      }
    }
  } catch (error) {
    console.error("[graphService] Empty folder cleanup failed:", error.message);
  }
}

/**
 * Paginate/Fetch the parent message of a given thread (conversationId) using Microsoft Graph.
 * It searches the conversation and finds the most recent message that is NOT the current message.
 */
export async function fetchParentMessageInThread(authToken, conversationId, currentItemId, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);
  // Fetch up to 5 messages in this conversation ordered by received date
  const path = `${getMailboxPrefix(options)}/messages?$filter=conversationId eq '${normalizeItemId(conversationId)}'&$orderby=receivedDateTime desc&$top=5`;
  const response = await runGraphRequest(token, path);
  const data = await response.json();
  
  if (!data || !data.value) return null;
  
  // Find the first message that is not the current message
  const parentMsg = data.value.find(m => m.id !== currentItemId);
  return parentMsg || null;
}

/**
 * Searches the user's Sent Items folder for a message matching the given subject.
 * Returns the first matching message or null.
 * Used to apply categories/subject updates to On-Send emails after they are sent.
 *
 * @param {string} authToken  - SSO or access token
 * @param {string} subject    - Exact subject string to search for
 * @param {object} [options]  - { isAccessToken }
 * @returns {object|null}     - The Graph message object or null
 */
export async function searchSentMessage(authToken, subject, options = {}) {
  const token = await resolveGraphAccessToken(authToken, options);

  // Escape single quotes for OData filter
  const escapedSubject = subject.replace(/'/g, "''");

  // NOTE: Do NOT add $orderby here. Microsoft Graph returns 400 InefficientFilter
  // when $filter on a non-indexed property (subject) is combined with $orderby
  // inside a specific mailFolder endpoint. The filter alone is sufficient.
  const path = `${getMailboxPrefix(options)}/mailFolders/sentitems/messages?$filter=subject eq '${escapedSubject}'&$top=5&$select=id,subject,sentDateTime,categories`;

  try {
    const response = await runGraphRequest(token, path);
    const data = await response.json();
    if (!data || !Array.isArray(data.value) || data.value.length === 0) {
      return null;
    }
    // Sort client-side by sentDateTime descending to get the newest match
    const sorted = data.value.sort((a, b) =>
      new Date(b.sentDateTime || 0) - new Date(a.sentDateTime || 0)
    );
    return sorted[0];
  } catch (err) {
    // Rethrow so the caller (fileService On-Send task) can apply the
    // direct-token vs OBO fallback logic correctly.
    console.warn(`[graphService] searchSentMessage failed for subject "${subject}":`, err.message);
    throw err;
  }
}
