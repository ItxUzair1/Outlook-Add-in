/* global Office */

/**
 * Gets the Identity Token (SSO Token) from Office.js.
 * This token is used by the backend to perform On-Behalf-Of actions.
 */
export async function getSsoToken() {
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

  return new Promise((resolve, reject) => {
    if (!Office?.auth?.getAccessToken) {
      reject(new Error("Office SSO Auth not supported in this environment."));
      return;
    }

    requestToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true })
      .then(resolve)
      .catch((primaryErr) => {
        const msg = String(primaryErr?.message || "").toLowerCase();
        const shouldRetryWithoutGraphHint =
          msg.includes("code: 7000") ||
          msg.includes("permission denied") ||
          msg.includes("sufficient permissions");

        if (!shouldRetryWithoutGraphHint) {
          reject(primaryErr);
          return;
        }

        console.warn("[mailboxService] SSO with forMSGraphAccess failed; retrying without forMSGraphAccess.");
        requestToken({ allowSignInPrompt: true, allowConsentPrompt: true })
          .then(resolve)
          .catch((fallbackErr) => {
            reject(new Error(`${fallbackErr.message}. Hint: verify IdentityAPI requirement in manifest and Outlook account permissions.`));
          });
      });
  });
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
          base64Content: att.content || "",
        }));
      }
      return [];
    }

    const attachments = await getAsync((cb) => item.getAttachmentsAsync(cb));
    const output = [];

    for (const att of attachments || []) {
      try {
        const content = await getAsync((cb) => item.getAttachmentContentAsync(att.id, cb));
        
        if (content && content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          const base64 = content.content || "";
          if (base64.length > 0) {
            output.push({
              id: att.id,
              name: att.name,
              base64Content: base64,
            });
          }
        }
      } catch (err) {
        console.warn(`[mailboxService] Error getting content for ${att.name}:`, err);
      }
    }

    return output;
  } catch (error) {
    console.error("[mailboxService] getAttachments error:", error);
    return [];
  }
}

function toAddressList(input) {
  if (!Array.isArray(input)) return [];
  return input.map((x) => x?.emailAddress || x?.displayName || "").filter(Boolean);
}

function toGraphItemId(itemId) {
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
  if (typeof Office === "undefined" || !Office.context?.mailbox?.item) {
    return null;
  }
  const item = Office.context.mailbox.item;
  const attMetadata = await getAttachmentMetadata(item);
  const bodyPreview = await getBodyPreview(item);
  
  return {
    itemId: toGraphItemId(item.itemId),
    internetMessageId: item.internetMessageId || item.itemId || "",
    subject: item.subject || "No Subject",
    sender: item.from?.emailAddress || item.from?.displayName || "",
    to: toAddressList(item.to),
    cc: toAddressList(item.cc),
    sentAt: item.dateTimeCreated || new Date().toISOString(),
    attachments: attMetadata,
    bodyPreview: String(bodyPreview || ""),
    body: "",
    isPartial: true
  };
}

export async function buildCurrentEmailPayload(options = {}) {
  const forceRefresh = !!options.forceRefresh;
  const allowCachedFallback = options.allowCachedFallback !== false;
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
  try {
    ssoToken = await getSsoToken();
  } catch (err) {
    ssoTokenError = err.message;
    console.warn("[mailboxService] SSO Token unavailable:", err.message);
  }

  const graphItemId = toGraphItemId(item?.itemId);
  
  let attachments = [];
  if (!ssoToken || !graphItemId) {
    attachments = await getAttachments(item);
  }

  let bodyPreview = "";
  try { bodyPreview = await getBodyPreview(item); } catch (e) {}
  
  let bodyHtml = "";
  try {
    bodyHtml = await getAsync((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
  } catch (e) {
    console.warn("[mailboxService] Failed to get HTML body, using preview.");
  }

  return {
    itemId: graphItemId,
    internetMessageId: item.internetMessageId || item.itemId || "",
    subject: item.subject || "No Subject",
    sender: item.from?.emailAddress || item.from?.displayName || "",
    to: toAddressList(item.to),
    cc: toAddressList(item.cc),
    sentAt: item.dateTimeCreated || new Date().toISOString(),
    bodyPreview: String(bodyPreview || ""),
    body: String(bodyHtml || bodyPreview || ""),
    isHtml: !!bodyHtml,
    attachments,
    ssoToken,
    ssoTokenError,
    isPartial: false
  };
}
