/* global Office */
import { toRestItemId, toEwsItemId } from "./itemIdUtils.js";
import { addCategoryToCurrentEmail } from "../services/mailboxService.js";

/**
 * Reports an error to the dialog/taskpane via localStorage.
 * @param {string} message The error message to report.
 */
export function reportActionError(message) {
  console.error("Reporting Action Error:", message);
  try {
    localStorage.setItem("koyomailActionError", JSON.stringify({
      message,
      timestamp: Date.now()
    }));
  } catch (e) {
    console.warn("localStorage write failed in reportActionError", e);
  }
}

export function buildDiagnostics(itemId) {
  const diagnostics = Office?.context?.diagnostics;
  const hostName = diagnostics?.hostName || "n/a";
  const hostVersion = diagnostics?.hostVersion || "n/a";
  const req15 = Office?.context?.requirements?.isSetSupported
    ? Office.context.requirements.isSetSupported("Mailbox", "1.5")
    : "n/a";
  return ` (Host: ${hostName}, V: ${hostVersion}, ID: ${String(itemId || "").substring(0, 8)}..., Req1.5: ${req15})`;
}

export function formatAfterFilingApiError(err, actionLabel, itemId) {
  const raw = err?.message || "Unknown error";
  const lower = raw.toLowerCase();
  const blocked = lower.includes("ews & rest blocked") ||
    lower.includes("rest token failed") ||
    lower.includes("makeewsrequestasync") ||
    lower.includes("callback token") ||
    lower.includes("exchange server returned an error");

  if (blocked) {
    return `${actionLabel} could not be completed automatically in this Outlook host. Email was filed successfully. Please move it manually.${buildDiagnostics(itemId)}`;
  }

  return `${actionLabel} failed: ${raw}${buildDiagnostics(itemId)}`;
}

export function isGraphPostFilingDeferralError(message) {
  if (!message) return false;
  const lower = String(message).toLowerCase();
  return (
    lower.includes("post-filing actions skipped") ||
    lower.includes("could not verify email id") ||
    lower.includes("graph authentication or email id unavailable")
  );
}

function archiveItemLocally(item, itemId) {
  return new Promise((resolve, reject) => {
    if (item?.archiveAsync && (!itemId || item.itemId === itemId)) {
      item.archiveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          moveItemViaEws(itemId || item.itemId, "archive").then(resolve).catch(reject);
        }
      });
      return;
    }
    moveItemViaEws(itemId, "archive").then(resolve).catch(reject);
  });
}

/**
 * Runs post-filing actions via Office.js / EWS when backend Graph actions failed.
 * Classic Outlook supports these host APIs even when Microsoft Graph enrichment fails.
 */
export async function runClientPostFilingFallback({
  itemId,
  afterFiling = "none",
  markReviewed = false,
  addFiledCategory = false,
  filedCategoryName = "Filed by Koyomail",
}) {
  const completed = [];
  const item = Office?.context?.mailbox?.item;
  const effectiveItemId = item?.itemId || itemId;

  if (addFiledCategory) {
    const categoryOk = await addCategoryToCurrentEmail(filedCategoryName);
    if (categoryOk) completed.push("category");
  }

  if (afterFiling && afterFiling !== "none" && afterFiling !== "add_date" && effectiveItemId) {
    if (afterFiling === "delete" || afterFiling === "move_deleted") {
      await deleteItemViaEws(effectiveItemId);
      completed.push("afterFiling");
    } else if (afterFiling === "archive") {
      await archiveItemLocally(item, effectiveItemId);
      completed.push("afterFiling");
    }
  }

  if (markReviewed) {
    // Read-mode items do not expose a universal mark-reviewed API in Classic Outlook.
    console.warn("[afterFilingUtils] Mark-as-reviewed is not supported via client fallback in this host.");
  }

  return {
    recovered: completed.length > 0,
    completed,
  };
}

/**
 * Asks the parent Outlook command surface to run post-filing actions (dialog context).
 */
export async function requestParentPostFilingFallback({
  itemId,
  afterFiling = "none",
  addFiledCategory = false,
  filedCategoryName = "Filed by Koyomail",
}) {
  if (!Office?.context?.ui?.messageParent) {
    return { recovered: false, completed: [] };
  }

  Office.context.ui.messageParent(JSON.stringify({
    action: "postFilingFallback",
    itemId,
    afterFiling,
    addFiledCategory,
    filedCategoryName,
  }));

  for (let secondsPassed = 0; secondsPassed < 10; secondsPassed += 1) {
    await new Promise((resolve) => setTimeout(resolve, 1000));
    const storedError = localStorage.getItem("koyomailActionError");
    if (storedError) {
      const { message: parentError } = JSON.parse(storedError);
      localStorage.removeItem("koyomailActionError");
      throw new Error(parentError);
    }
  }

  return { recovered: true, completed: ["parent"] };
}

/**
 * Attempts to recover post-filing actions after a backend Graph failure.
 */
export async function recoverPostFilingAfterGraphFailure({
  postFilingError,
  itemId,
  afterFiling = "none",
  markReviewed = false,
  addFiledCategory = false,
  filedCategoryName = "Filed by Koyomail",
}) {
  if (!isGraphPostFilingDeferralError(postFilingError)) {
    return { recovered: false, completed: [] };
  }

  const hasLocalItem = !!Office?.context?.mailbox?.item;
  if (hasLocalItem) {
    return runClientPostFilingFallback({
      itemId,
      afterFiling,
      markReviewed,
      addFiledCategory,
      filedCategoryName,
    });
  }

  if (Office?.context?.ui?.messageParent) {
    return requestParentPostFilingFallback({
      itemId,
      afterFiling,
      addFiledCategory,
      filedCategoryName,
    });
  }

  return { recovered: false, completed: [] };
}

/**
 * Deletes an item using EWS (Exchange Web Services).
 * @param {string} itemId The EWS ItemId of the email.
 * @returns {Promise<void>}
 */
export function deleteItemViaEws(itemId) {
  return moveItemViaEws(itemId, "deleteditems");
}

/**
 * Moves an item to a distinguished folder using EWS.
 * @param {string} itemId The EWS ItemId.
 * @param {string} folderId The distinguished folder ID (e.g. 'deleteditems', 'archive').
 * @param {boolean} [useHeader=true] Whether to include the Exchange RequestServerVersion soap header.
 * @returns {Promise<void>}
 */
export function moveItemViaEws(itemId, folderId, useHeader = true) {
  return new Promise((resolve, reject) => {
    const ewsItemId = toEwsItemId(itemId);

    // Simple XML escape
    const escapedId = ewsItemId.replace(/&/g, '&amp;')
                           .replace(/</g, '&lt;')
                           .replace(/>/g, '&gt;')
                           .replace(/"/g, '&quot;')
                           .replace(/'/g, '&apos;');

    let body = "";
    if (folderId === "deleteditems") {
      body = '<m:DeleteItem DeleteType="MoveToDeletedItems">' +
               '<m:ItemIds><t:ItemId Id="' + escapedId + '" /></m:ItemIds>' +
             '</m:DeleteItem>';
    } else {
      body = '<m:MoveItem>' +
               '<m:ToFolderId><t:DistinguishedFolderId Id="' + folderId + '" /></m:ToFolderId>' +
               '<m:ItemIds><t:ItemId Id="' + escapedId + '" /></m:ItemIds>' +
             '</m:MoveItem>';
    }

    const header = useHeader ? '<soap:Header><t:RequestServerVersion Version="Exchange2010" /></soap:Header>' : '<soap:Header />';
    const ewsRequest = 
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
                     'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
                     'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
                     'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
         header +
         '<soap:Body>' + body + '</soap:Body>' +
      '</soap:Envelope>';

    Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
      const responseXml = result.value;
      const diag = buildDiagnostics(itemId);

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        if (typeof responseXml === 'string' && (responseXml.includes("ResponseCode>NoError</") || responseXml.includes('ResponseClass="Success"'))) {
          resolve();
        } else if (useHeader && (!responseXml || responseXml.trim() === "")) {
          console.log("EWS empty with header, retrying without header...");
          moveItemViaEws(itemId, folderId, false).then(resolve).catch(reject);
        } else if (!useHeader && (!responseXml || responseXml.trim() === "")) {
          console.log("EWS failed. Trying REST fallback...");
          tryPostFilingActionViaRest(itemId, folderId).then(resolve).catch((restErr) => {
            reject(new Error("EWS & REST blocked. " + restErr.message + diag));
          });
        } else {
          reject(new Error("EWS SOAP Error: " + (typeof responseXml === 'string' ? responseXml.substring(0, 100) : "Unknown EWS response")));
        }
      } else {
        const detail = result.error ? `${result.error.name}: ${result.error.message}` : "Request failed";
        reject(new Error(detail + diag));
      }
    });
  });
}

export function tryPostFilingActionViaRest(itemId, folderId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const accessToken = result.value;
        const ewsUrl = Office.context.mailbox.ewsUrl;
        const restItemId = toRestItemId(itemId);
        const restUrl = ewsUrl.toLowerCase().includes("outlook.office365.com") || ewsUrl.toLowerCase().includes("outlook.office.com") 
          ? "https://outlook.office.com/api/v2.0" 
          : ewsUrl.replace("/ews/exchange.asmx", "/api/v2.0");

        const actionUrl = folderId === "deleteditems" 
          ? `${restUrl}/me/messages/${restItemId}`
          : `${restUrl}/me/messages/${restItemId}/move`;

        const method = folderId === "deleteditems" ? "DELETE" : "POST";
        const body = folderId === "deleteditems" ? null : JSON.stringify({ "DestinationId": folderId === "archive" ? "archive" : folderId });

        fetch(actionUrl, {
          method: method,
          headers: {
            "Authorization": "Bearer " + accessToken,
            "Content-Type": "application/json"
          },
          body: body
        }).then(response => {
          if (response.ok) {
            resolve();
          } else {
            response.text().then(txt => reject(new Error(`REST ${response.status}: ${txt.substring(0,50)}`)));
          }
        }).catch(err => reject(new Error("REST Fetch failed: " + err.message)));
      } else {
        const tokenError = result.error?.message || "Token unavailable in this Outlook host";
        reject(new Error("REST token failed: " + tokenError));
      }
    });
  });
}
