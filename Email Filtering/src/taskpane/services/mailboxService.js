/* global Office */

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
  if (!item?.body?.getAsync) {
    return "";
  }

  const value = await getAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
  return String(value || "").slice(0, 4000);
}

async function getAttachments(item) {
  try {
    if (!item?.getAttachmentsAsync || !item?.getAttachmentContentAsync) {
      // Fallback for some older Win32 versions where attachments might be a property
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
  if (!Array.isArray(input)) {
    return [];
  }

  return input.map((x) => x?.emailAddress || x?.displayName || "").filter(Boolean);
}

export async function buildCurrentEmailPayload() {
  // Check for cached payload from parent (useful for Dialog mode)
  const cached = localStorage.getItem("currentEmailPayload");
  if (cached) {
    try {
      const { payload, timestamp } = JSON.parse(cached);
      // Increase timeout to 5 minutes (300,000 ms) as users may take time to select a location
      const isFresh = (Date.now() - timestamp) < 300000;
      
      if (isFresh) {
        console.log("[mailboxService] Using cached email payload from parent context.");
        // Clear cache after successful read to prevent using stale data for next filing
        localStorage.removeItem("currentEmailPayload");
        return payload;
      } else {
        console.warn("[mailboxService] Cached payload is stale, falling back to Office.js.");
        localStorage.removeItem("currentEmailPayload");
      }
    } catch (err) {
      console.warn("[mailboxService] Error parsing cached payload:", err);
      localStorage.removeItem("currentEmailPayload");
    }
  } else {
    console.log("[mailboxService] No cached payload found, using Office.js.");
  }

  if (typeof Office === "undefined" || !Office.context?.mailbox?.item) {
    throw new Error("Office.js is not initialized yet or mailbox item is not available.");
  }

  const item = Office.context.mailbox.item;

  const attachments = await getAttachments(item);
  const bodyPreview = await getBodyPreview(item);
  const bodyHtml = await getAsync((cb) => item.body.getAsync(Office.CoercionType.Html, cb));

  const sender = item.from?.emailAddress || item.from?.displayName || "";
  const to = toAddressList(item.to);
  const cc = toAddressList(item.cc);

  return {
    internetMessageId: item.internetMessageId || item.itemId || "",
    subject: item.subject || "No Subject",
    sender,
    to,
    cc,
    sentAt: item.dateTimeCreated || new Date().toISOString(),
    bodyPreview,
    body: bodyHtml,
    isHtml: true,
    attachments,
  };
}
