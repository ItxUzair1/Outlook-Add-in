/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { buildCurrentEmailPayload } from "../taskpane/services/mailboxService";
import { toRestItemId, toEwsItemId } from "../taskpane/utils/itemIdUtils.js";

Office.onReady(() => {
  // Update heartbeat to let dialog know the background context is alive
  localStorage.setItem("koyomailCommandsHeartbeat", Date.now());
  setInterval(() => {
    localStorage.setItem("koyomailCommandsHeartbeat", Date.now());
  }, 1000);
});

/**
 * Reports an error to the dialog via localStorage.
 * @param {string} message The error message to report.
 */
function reportActionError(message) {
  console.error("Reporting Action Error:", message);
  localStorage.setItem("koyomailActionError", JSON.stringify({
    message,
    timestamp: Date.now()
  }));
}

function buildDiagnostics(itemId) {
  const diagnostics = Office?.context?.diagnostics;
  const hostName = diagnostics?.hostName || "n/a";
  const hostVersion = diagnostics?.hostVersion || "n/a";
  const req15 = Office?.context?.requirements?.isSetSupported
    ? Office.context.requirements.isSetSupported("Mailbox", "1.5")
    : "n/a";
  return ` (Host: ${hostName}, V: ${hostVersion}, ID: ${String(itemId || "").substring(0, 8)}..., Req1.5: ${req15})`;
}

function formatAfterFilingApiError(err, actionLabel, itemId) {
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

/**
 * Deletes an item using EWS (Exchange Web Services).
 * @param {string} itemId The EWS ItemId of the email.
 * @returns {Promise<void>}
 */
function deleteItemViaEws(itemId) {
  return moveItemViaEws(itemId, "deleteditems");
}

/**
 * Moves an item to a distinguished folder using EWS.
 * @param {string} itemId The EWS ItemId.
 * @param {string} folderId The distinguished folder ID (e.g. 'deleteditems', 'archive').
 * @returns {Promise<void>}
 */
function moveItemViaEws(itemId, folderId, useHeader = true) {
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
          // If empty with header, try one last time WITHOUT header
          console.log("EWS empty with header, retrying without header...");
          moveItemViaEws(itemId, folderId, false).then(resolve).catch(reject);
        } else if (!useHeader && (!responseXml || responseXml.trim() === "")) {
          // Both EWS attempts failed with empty. Try REST fallback.
          console.log("EWS failed. Trying REST fallback...");
          tryPostFilingActionViaRest(itemId, folderId).then(resolve).catch((restErr) => {
            reject(new Error("EWS & REST blocked. " + restErr.message + diag));
          });
        } else {
          // ... rest of error logic
          reject(new Error("EWS SOAP Error: " + responseXml.substring(0,50)));
        }
      } else {
        const detail = result.error ? `${result.error.name}: ${result.error.message}` : "Request failed";
        reject(new Error(detail + diag));
      }
    });
  });
}

function tryPostFilingActionViaRest(itemId, folderId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const accessToken = result.value;
        const ewsUrl = Office.context.mailbox.ewsUrl;
        const restItemId = toRestItemId(itemId);
        // Construct REST URL from EWS URL
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


function showMilestoneNotification(event, featureName, isStatusUpdate = false) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: isStatusUpdate 
      ? featureName 
      : `${featureName} is available as a button in Milestone 2 but full functionality is planned for the next milestone.`,
    icon: "Icon.80",
    persistent: false,
  };

  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    isStatusUpdate ? "StatusNotification" : `${featureName}MilestoneNotification`,
    message
  );

  if (event && event.completed) {
    event.completed();
  }
}

function searchAction(event) {
  showMilestoneNotification(event, "Search");
}

let dialog;

async function openFilingDialogAction(event) {
  // Clear any existing stale payload
  localStorage.removeItem("currentEmailPayload");

  // Step 1: Gather fast metadata and cache it immediately
  // This ensures the dialog opens with the Subject/Sender filled even if attachments take time.
  const { buildEmailMetadata, buildCurrentEmailPayload } = require("../taskpane/services/mailboxService");
  const metadata = await buildEmailMetadata();
  if (metadata) {
    localStorage.setItem("currentEmailPayload", JSON.stringify({
      payload: metadata,
      timestamp: Date.now()
    }));
    console.log("[commands] Fast metadata cached.");
  }

  // Step 2: Open Dialog immediately

  // Step 3: Start heavy enrichment in the background (Body, Attachments, SSO)
  // We do NOT await this before opening the dialog, so the dialog remains responsive.
  (async () => {
    try {
      console.log("[commands] Starting background enrichment...");
      let fullPayload = null;
      for (let attempt = 1; attempt <= 10; attempt += 1) {
        try {
          fullPayload = await buildCurrentEmailPayload({ forceRefresh: true, allowCachedFallback: false });
          if (fullPayload && !fullPayload.isPartial) {
            break;
          }
        } catch (err) {
          // Keep retrying while mailbox item initializes in command context.
        }
        await new Promise((resolve) => setTimeout(resolve, 300));
      }

      if (!fullPayload || fullPayload.isPartial) {
        throw new Error("Background enrichment could not retrieve full payload in command context.");
      }

      localStorage.setItem("currentEmailPayload", JSON.stringify({
        payload: fullPayload,
        timestamp: Date.now()
      }));
      console.log("[commands] Background enrichment complete (Body & Attachments cached).");
    } catch (error) {
      console.warn("[commands] Background enrichment failed:", error.message);
    }
  })();

  console.log("[commands] Preparing to open dialog...");

  // Use the origin of the current command to derive the dialog URL
  const dialogUrl = `${window.location.origin}/taskpane.html?mode=file`;

  // displayInIframe is needed for some environments, but 80% width/height gives a good desktop size
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 70, width: 70, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Dialog failed to open: " + asyncResult.error.message);
        if (event && event.completed) {
          event.completed();
        }
        return;
      }
      
      dialog = asyncResult.value;

      // Handle events from the dialog (e.g., manual closure)
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.log("Dialog event received:", arg.error);
        // This includes the user clicking the 'X' button
        if (event && event.completed) {
          event.completed();
        }
      });

      // Handle messages from the dialog (e.g., filing actions)
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") {
          dialog.close();
          if (event && event.completed) event.completed();
          return;
        }

        try {
          const data = JSON.parse(arg.message);
          if (data.action === "afterFiling") {
            const item = Office.context.mailbox.item;
            if (!item) {
              const errMsg = "AfterFiling: No mailbox item found in parent context.";
              reportActionError(errMsg);
              dialog.close();
              if (event && event.completed) event.completed();
              return;
            }

            if (data.value === "delete") {
              // Avoid removeAsync here: some Outlook hosts can treat it as hard-delete.
              // Use EWS MoveToDeletedItems first (with REST fallback inside moveItemViaEws).
              deleteItemViaEws(item.itemId)
                .then(() => {
                  dialog.close();
                  if (event && event.completed) event.completed();
                })
                .catch((err) => {
                  reportActionError(formatAfterFilingApiError(err, "Delete", item.itemId));
                  if (event && event.completed) event.completed();
                });
            } else if (data.value === "archive") {
              if (item.archiveAsync) {
                item.archiveAsync((result) => {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Email archived via archiveAsync.");
                    dialog.close();
                    if (event && event.completed) event.completed();
                  } else {
                    console.warn("archiveAsync failed, trying EWS fallback: " + result.error.message);
                    moveItemViaEws(item.itemId, "archive")
                      .then(() => {
                        dialog.close();
                        if (event && event.completed) event.completed();
                      })
                      .catch((err) => {
                        reportActionError(formatAfterFilingApiError(err, "Archive", item.itemId));
                        if (event && event.completed) event.completed();
                      });
                  }
                });
              } else {
                console.log("archiveAsync not found, using EWS fallback directly.");
                moveItemViaEws(item.itemId, "archive")
                  .then(() => {
                    dialog.close();
                    if (event && event.completed) event.completed();
                  })
                  .catch((err) => {
                    reportActionError(formatAfterFilingApiError(err, "Archive", item.itemId));
                    if (event && event.completed) event.completed();
                  });
              }
            }
          }
        } catch (e) {
          reportActionError("Error processing dialog message: " + e.message);
          if (event && event.completed) event.completed();
          console.error(e);
        }
      });

    }
  );
}

function suggestedAction(event) {
  showMilestoneNotification(event, "Suggested locations");
}



function commentsAction(event) {
  showMilestoneNotification(event, "Comments");
}

function optionsAction(event) {
  showMilestoneNotification(event, "Options");
}

function helpAction(event) {
  const dialogUrl = `${window.location.origin}/taskpane.html?mode=help`;
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 75, width: 75, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Help dialog failed to open: " + asyncResult.error.message);
        if (event && event.completed) event.completed();
        return;
      }
      
      const helpDialog = asyncResult.value;
      helpDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") {
          helpDialog.close();
          if (event && event.completed) event.completed();
        }
      });

      helpDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (event && event.completed) event.completed();
      });
    }
  );
}

Office.actions.associate("searchAction", searchAction);
Office.actions.associate("optionsAction", optionsAction);
Office.actions.associate("suggestedAction", suggestedAction);
Office.actions.associate("commentsAction", commentsAction);
Office.actions.associate("helpAction", helpAction);
Office.actions.associate("openFilingDialogAction", openFilingDialogAction);
