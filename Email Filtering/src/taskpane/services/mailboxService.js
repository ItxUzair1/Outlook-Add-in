/* global Office */

import { isOutlookIframeHost } from "../utils/authManager.js";
import { SUCCESS_CATEGORY_COLOR } from "../utils/filingCategoryUtils.js";

/**
 * Gets the Identity Token (SSO Token) from Office.js.
 * This token is used by the backend to perform On-Behalf-Of actions.
 */
export async function getSsoToken() {
  const timeoutPromise = new Promise((_, reject) =>
    setTimeout(() => reject(new Error("SSO Token Timeout")), 8000)
  );

  const tokenPromise = (async () => {
    const requestToken = (options) => new Promise((resolve, reject) => {
      Office.auth.getAccessToken(options, (result) => {
        try {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            const code = result.error ? result.error.code : "Unknown";
            const msg = result.error ? result.error.message : "No error message provided by Office";
            reject(new Error(`SSO Token Failed: ${msg} (Code: ${code})`));
          }
        } catch (e) {
          reject(new Error(`Critical failure in SSO callback: ${e.message}`));
        }
      });
    });

    if (!Office?.auth?.getAccessToken) {
      throw new Error("Office SSO Auth not supported in this environment.");
    }

    try {
      return await requestToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    } catch (primaryErr) {
      const msg = String(primaryErr?.message || "").toLowerCase();
      const shouldRetryWithoutGraphHint =
        msg.includes("code: 7000") ||
        msg.includes("permission denied") ||
        msg.includes("sufficient permissions");

      if (!shouldRetryWithoutGraphHint) {
        throw primaryErr;
      }

      console.warn("[mailboxService] SSO with forMSGraphAccess failed; retrying without forMSGraphAccess.");
      return await requestToken({ allowSignInPrompt: true, allowConsentPrompt: true });
    }
  })();

  return Promise.race([tokenPromise, timeoutPromise]);
}

function getAsync(executor) {
  return new Promise((resolve, reject) => {
    executor((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message || "Office async operation failed"));
      }
    });
  });
}

async function getBodyPreview(item) {
  try {
    if (!item?.body?.getAsync) {
      return "";
    }
    const value = await getAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
    return String(value || "").slice(0, 4000);
  } catch (err) {
    console.warn("[mailboxService] getBodyPreview failed:", err.message);
    return "";
  }
}

/**
 * Gets ONLY the names, IDs, and sizes of attachments (without the heavy content).
 * This is much faster and can be done synchronously.
 */
async function getAttachmentMetadata(item) {
  try {
    if (!item?.getAttachmentsAsync) {
      return Array.isArray(item?.attachments) ? item.attachments.map(att => ({
        id: att.id || att.name,
        name: att.name,
        size: att.size || 0,
        isMetadataOnly: true
      })) : [];
    }

    const attachments = await getAsync((cb) => item.getAttachmentsAsync(cb));
    return (attachments || []).map(att => ({
      id: att.id,
      name: att.name,
      size: att.size,
      contentType: att.contentType,
      isInline: att.isInline,
      isMetadataOnly: true
    }));
  } catch (error) {
    console.error("[mailboxService] getAttachmentMetadata error:", error);
    return [];
  }
}

async function getAttachments(item) {
  try {
    if (!item?.getAttachmentsAsync || !item?.getAttachmentContentAsync) {
      if (Array.isArray(item?.attachments) && item.attachments.length > 0) {
          return item.attachments.map(att => ({
          id: att.id || att.name,
          name: att.name,
          isInline: !!att.isInline,
          contentType: att.contentType || "application/octet-stream",
          base64Content: att.content || "",
        }));
      }
      return [];
    }

    const attachments = await getAsync((cb) => item.getAttachmentsAsync(cb));

    const attachmentPromises = (attachments || []).map(async (att) => {
      try {
        const content = await getAsync((cb) => item.getAttachmentContentAsync(att.id, cb));
        
        if (content && content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          const base64 = content.content || "";
          if (base64.length > 0) {
            return {
              id: att.id,
              name: att.name,
              isInline: !!att.isInline,
              contentType: att.contentType || "application/octet-stream",
              base64Content: base64,
            };
          }
        }
      } catch (err) {
        console.warn(`[mailboxService] Error getting content for ${att.name}:`, err);
      }
      return null;
    });

    const output = (await Promise.all(attachmentPromises)).filter(Boolean);

    return output;
  } catch (error) {
    console.error("[mailboxService] getAttachments error:", error);
    return [];
  }
}

async function getComposeProperty(prop) {
  if (!prop) return null;
  if (typeof prop === "string" || Array.isArray(prop)) return prop;
  if (prop.getAsync) {
    try {
      return await getAsync(cb => prop.getAsync(cb));
    } catch { return null; }
  }
  return prop;
}

function toAddressList(input) {
  if (!Array.isArray(input)) return [];
  return input.map((x) => x?.emailAddress || x?.displayName || "").filter(Boolean);
}

export function toGraphItemId(itemId) {
  try {
    const mailbox = Office?.context?.mailbox;
    if (mailbox?.convertToRestId && Office?.MailboxEnums?.RestVersion?.v2_0) {
      return mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    }
  } catch (error) {
    console.warn("[mailboxService] convertToRestId failed:", error);
  }
  return itemId;
}

/**
 * Gathers ONLY basic metadata (Subject, ID) that can be retrieved instantly.
 */
export async function buildEmailMetadata() {
  let item = Office.context?.mailbox?.item;
  if (!item) {
    for (let i = 0; i < 20; i++) {
      await new Promise(resolve => setTimeout(resolve, 100));
      item = Office.context?.mailbox?.item;
      if (item) break;
    }
  }

  if (typeof Office === "undefined" || !item) {
    return null;
  }
  const attMetadata = await getAttachmentMetadata(item);
  const bodyPreview = await getBodyPreview(item);
  
  const rawSubject = await getComposeProperty(item.subject);
  const rawTo = await getComposeProperty(item.to);
  const rawCc = await getComposeProperty(item.cc);
  const rawFrom = await getComposeProperty(item.from);

  const userProfile = Office.context?.mailbox?.userProfile;
  const userDisplayName = userProfile?.displayName || "";
  const userEmail = userProfile?.emailAddress || "";

  let resolvedSenderName = rawFrom?.displayName || item.from?.displayName || "";
  if ((!resolvedSenderName || resolvedSenderName.includes("@")) && userDisplayName) {
    resolvedSenderName = userDisplayName;
  }

  const resolvedSender = rawFrom?.emailAddress || item.from?.emailAddress || userEmail || "";

  return {
    itemId: toGraphItemId(item.itemId),
    internetMessageId: item.internetMessageId || item.itemId || "",
    conversationId: item.conversationId || "",
    subject: typeof rawSubject === "string" ? rawSubject : "No Subject",
    sender: resolvedSender,
    senderName: resolvedSenderName,
    to: toAddressList(rawTo || item.to),
    cc: toAddressList(rawCc || item.cc),
    sentAt: item.dateTimeCreated || new Date().toISOString(),
    attachments: attMetadata,
    hasAttachments: attMetadata.length > 0,
    bodyPreview: String(bodyPreview || ""),
    body: "",
    isPartial: true
  };
}

export async function buildCurrentEmailPayload(options = {}) {
  const forceRefresh = !!options.forceRefresh;
  const allowCachedFallback = options.allowCachedFallback !== false;
  const isOnSend = !!options.isOnSend;
  let cachedPayload = null;

  const cached = localStorage.getItem("currentEmailPayload");
  if (cached) {
    try {
      const { payload, timestamp } = JSON.parse(cached);
      if ((Date.now() - timestamp) < 300000) {
        cachedPayload = payload;

        // Normal callers can use cache directly. Forced refresh is used by commands
        // to avoid being stuck with a previously cached partial payload.
        if (!forceRefresh || !payload.isPartial) {
          if (!payload.isPartial) localStorage.removeItem("currentEmailPayload");
          return payload;
        }
      }
    } catch (err) {
      localStorage.removeItem("currentEmailPayload");
    }
  }

  let item = Office.context?.mailbox?.item;
  if (!item) {
    for (let i = 0; i < 20; i++) {
      await new Promise(resolve => setTimeout(resolve, 100));
      item = Office.context?.mailbox?.item;
      if (item) break;
    }
  }

  if (typeof Office === "undefined" || !Office.context?.mailbox || !item) {
    const mode = new URLSearchParams(window.location.search).get("mode");
    if (mode === "help") return null;

    // Dialog contexts may not have direct mailbox item access.
    if (cachedPayload && allowCachedFallback) {
      return cachedPayload;
    }

    throw new Error("No mailbox item is currently selected.");
  }

  let ssoToken = null;
  let ssoTokenError = null;
  // Office SSO is unreliable in New Outlook / filing-dialog iframes — skip to avoid timeouts.
  if (!isOutlookIframeHost()) {
    try {
      ssoToken = await getSsoToken();
    } catch (err) {
      ssoTokenError = err.message;
      console.warn("[mailboxService] SSO Token unavailable:", err.message);
    }
  }

  const graphItemId = toGraphItemId(item?.itemId);
  
  const attMetadata = await getAttachmentMetadata(item);

  let attachments = [];
  if (!ssoToken || !graphItemId || isOnSend) {
    attachments = await getAttachments(item);
  } else {
    attachments = (attMetadata || []).map(att => ({ ...att, isMetadataOnly: true }));
  }

  let bodyPreview = "";
  try { bodyPreview = await getBodyPreview(item); } catch (e) {}
  
  let bodyHtml = "";
  try {
    bodyHtml = await getAsync((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
  } catch (e) {
    console.warn("[mailboxService] Failed to get HTML body, using preview.");
  }

  const rawSubject = await getComposeProperty(item.subject);
  const rawTo = await getComposeProperty(item.to);
  const rawCc = await getComposeProperty(item.cc);
  const rawFrom = await getComposeProperty(item.from);

  const userProfile = Office.context?.mailbox?.userProfile;
  const userDisplayName = userProfile?.displayName || "";
  const userEmail = userProfile?.emailAddress || "";

  let resolvedSenderName = rawFrom?.displayName || item.from?.displayName || "";
  if ((!resolvedSenderName || resolvedSenderName.includes("@")) && userDisplayName) {
    resolvedSenderName = userDisplayName;
  }

  const resolvedSender = rawFrom?.emailAddress || item.from?.emailAddress || userEmail || "";

  return {
    itemId: graphItemId,
    internetMessageId: item.internetMessageId || item.itemId || "",
    conversationId: item.conversationId || "",
    subject: typeof rawSubject === "string" ? rawSubject : "No Subject",
    sender: resolvedSender,
    senderName: resolvedSenderName,
    to: toAddressList(rawTo || item.to),
    cc: toAddressList(rawCc || item.cc),
    sentAt: item.dateTimeCreated || new Date().toISOString(),
    bodyPreview: String(bodyPreview || ""),
    body: String(bodyHtml || bodyPreview || ""),
    isHtml: !!bodyHtml,
    attachments,
    hasAttachments: attMetadata.length > 0,
    ssoToken,
    ssoTokenError,
    isPartial: false
  };
}

export async function ensureMasterCategory(categoryName, color) {
  if (!Office.context?.mailbox?.masterCategories) return false;
  const targetColor = color || (Office.MailboxEnums && Office.MailboxEnums.CategoryColor ? Office.MailboxEnums.CategoryColor.Preset19 : SUCCESS_CATEGORY_COLOR);
  return new Promise((resolve) => {
    Office.context.mailbox.masterCategories.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        const existing = res.value.find(c => c.displayName === categoryName);
        if (existing) {
          // If the category exists but its color does not match the target color, recreate it to update the color
          if (existing.color !== targetColor) {
            console.log(`[mailboxService] Category exists but color is ${existing.color}, updating to ${targetColor}`);
            Office.context.mailbox.masterCategories.removeAsync([categoryName], (remRes) => {
              if (remRes.status === Office.AsyncResultStatus.Succeeded) {
                Office.context.mailbox.masterCategories.addAsync([{ displayName: categoryName, color: targetColor }], (addRes) => {
                  resolve(addRes.status === Office.AsyncResultStatus.Succeeded);
                });
              } else {
                resolve(false);
              }
            });
          } else {
            resolve(true);
          }
        } else {
          Office.context.mailbox.masterCategories.addAsync([{ displayName: categoryName, color: targetColor }], (addRes) => {
            resolve(addRes.status === Office.AsyncResultStatus.Succeeded);
          });
        }
      } else {
        resolve(false);
      }
    });
  });
}

/**
 * Adds a category to the currently selected email directly on the client.
 */
export async function addCategoryToCurrentEmail(categoryName, color = SUCCESS_CATEGORY_COLOR) {
  const item = Office.context?.mailbox?.item;
  if (!item || !item.categories) return false;
  
  await ensureMasterCategory(categoryName, color);
  
  return new Promise((resolve) => {
    item.categories.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        const existing = res.value;
        if (!existing.includes(categoryName)) {
          item.categories.addAsync([categoryName], (addRes) => {
             resolve(addRes.status === Office.AsyncResultStatus.Succeeded);
          });
        } else {
          resolve(true);
        }
      } else {
        resolve(false);
      }
    });
  });
}
