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
  if (!item?.getAttachmentsAsync || !item?.getAttachmentContentAsync) {
    return [];
  }

  const attachments = await getAsync((cb) => item.getAttachmentsAsync(cb));
  const output = [];

  for (const att of attachments || []) {
    try {
      const content = await getAsync((cb) => item.getAttachmentContentAsync(att.id, cb));
      if (content && content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
        output.push({
          id: att.id,
          name: att.name,
          base64Content: content.content,
        });
      }
    } catch {
      // Some attachment types are not retrievable through this API.
    }
  }

  return output;
}

function toAddressList(input) {
  if (!Array.isArray(input)) {
    return [];
  }

  return input.map((x) => x?.emailAddress || x?.displayName || "").filter(Boolean);
}

export async function buildCurrentEmailPayload() {
  const item = Office.context.mailbox.item;
  if (!item) {
    throw new Error("No selected email item found in Outlook context.");
  }

  const attachments = await getAttachments(item);
  const bodyPreview = await getBodyPreview(item);

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
    attachments,
  };
}
