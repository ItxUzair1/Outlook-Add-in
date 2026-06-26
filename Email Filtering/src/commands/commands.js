/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { toRestItemId } from "../taskpane/utils/itemIdUtils.js";
import { getLocations, fileEmail, remoteLog } from "../taskpane/services/backendApi";
import {
  reportActionError,
  formatAfterFilingApiError,
  deleteItemViaEws,
  moveItemViaEws,
  addCategoryViaEws,
  executeAfterFilingMoveByItemId,
} from "../taskpane/utils/afterFilingUtils.js";
import { buildEmailMetadata, buildCurrentEmailPayload, addCategoryToCurrentEmail } from "../taskpane/services/mailboxService";
import { enqueueFilingJob } from "../taskpane/services/filingQueue.js";

function handleOpenDialogRequest() {
  try {
    const req = localStorage.getItem("koyomailOpenDialogRequest");
    if (req) {
      localStorage.removeItem("koyomailOpenDialogRequest");
      localStorage.setItem("koyomail_dialog_mode", "file_dialog");
      const dialogUrl = `${window.location.origin}/taskpane.html?mode=file_dialog`;
      openDialogWithHandlers(dialogUrl, null);
    }
  } catch (e) {
    console.warn("Error handling open dialog request in commands.js:", e);
  }
}

Office.onReady(() => {
  // Update heartbeat to let dialog know the background context is alive
  localStorage.setItem("koyomailCommandsHeartbeat", Date.now());
  
  // Instant trigger via storage listener
  window.addEventListener("storage", (e) => {
    if (e.key === "koyomailOpenDialogRequest" && e.newValue) {
      handleOpenDialogRequest();
    }
  });

  setInterval(() => {
    localStorage.setItem("koyomailCommandsHeartbeat", Date.now());
    handleOpenDialogRequest();
  }, 1000);
});


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
  try { localStorage.setItem("koyomail_dialog_mode", "search"); } catch(e) {}
  const dialogUrl = `${window.location.origin}/taskpane.html?mode=search`;

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 75, width: 80, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Search dialog failed to open: " + asyncResult.error.message);
        if (event && event.completed) event.completed();
        return;
      }

      const searchDialog = asyncResult.value;
      searchDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") {
          searchDialog.close();
          if (event && event.completed) event.completed();
        }
      });
      searchDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        if (event && event.completed) event.completed();
      });
    }
  );
}

let dialog;

async function openFilingDialogAction(event) {
  // Clear any existing stale payload
  localStorage.removeItem("currentEmailPayload");
  localStorage.removeItem("multiEmailPayload");

  await handleSingleSelectFiling(event);
}

async function handleSingleSelectFiling(event) {
  console.log("[commands] Handling single-select filing");
  
  const metadata = await buildEmailMetadata();
  if (metadata) {
    localStorage.setItem("currentEmailPayload", JSON.stringify({
      payload: metadata,
      timestamp: Date.now()
    }));
    console.log("[commands] Fast metadata cached.");
  }

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
        } catch (err) {}
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

  try { localStorage.setItem("koyomail_dialog_mode", "file"); } catch(e) {}
  const dialogUrl = `${window.location.origin}/taskpane.html?mode=file`;
  openDialogWithHandlers(dialogUrl, event);
}

function openDialogWithHandlers(dialogUrl, event) {
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 75, width: 75, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Dialog failed to open: " + asyncResult.error.message);
        if (event && event.completed) {
          event.completed();
        }
        return;
      }
      
      dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.log("Dialog event received:", arg.error);
        if (event && event.completed) event.completed();
      });

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") {
          dialog.close();
          if (event && event.completed) event.completed();
          return;
        }

        if (arg.message.startsWith("backgroundFile:")) {
          dialog.close();
          if (event && event.completed) event.completed();
          try {
            const { payload, meta } = JSON.parse(arg.message.substring(15));
            enqueueFilingJob({ payload, meta });
          } catch (err) {
            console.error("[commands] backgroundFile failed:", err);
          }
          return;
        }

        try {
          const data = JSON.parse(arg.message);
          if (data.action === "afterFiling") {
            // For multi-select, data.itemId will be provided. For single-select, fallback to item.itemId
            const item = Office.context.mailbox.item;
            const targetItemId = data.itemId || (item ? item.itemId : null);

            if (!targetItemId) {
              const errMsg = "AfterFiling: No itemId provided or mailbox item found.";
              reportActionError(errMsg);
              dialog.close();
              if (event && event.completed) event.completed();
              return;
            }

            if (data.value === "delete") {
              deleteItemViaEws(targetItemId)
                .then(() => {
                  // Only close dialog if we're done (the UI will send a close event separately if needed, 
                  // or we just let it run. Wait, if it's multi-select, we don't want to close the dialog 
                  // on the FIRST afterFiling event. Let's let the UI handle closing!)
                  // Actually, we shouldn't close the dialog here if the UI sends "afterFiling" for multi-select.
                  // For now, let's let the UI send the "close" message when all are done.
                })
                .catch((err) => {
                  reportActionError(formatAfterFilingApiError(err, "Delete", targetItemId));
                });
            } else if (data.value === "archive") {
              // For single select we tried archiveAsync, but for multi-select we must use EWS.
              // Just use EWS for both to simplify, or check if we have the item.
              if (item && item.itemId === targetItemId && item.archiveAsync) {
                item.archiveAsync((result) => {
                  if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    moveItemViaEws(targetItemId, "archive").catch(err => reportActionError(formatAfterFilingApiError(err, "Archive", targetItemId)));
                  }
                });
              } else {
                moveItemViaEws(targetItemId, "archive").catch(err => reportActionError(formatAfterFilingApiError(err, "Archive", targetItemId)));
              }
            }
          } else if (data.action === "postFilingFallback") {
            const item = Office.context.mailbox.item;
            const targetItemId = data.itemId || (item ? item.itemId : null);

            const runFallback = async () => {
              if (data.addFiledCategory && targetItemId) {
                const categoryName = data.filedCategoryName || "Filed by Koyomail";
                await addCategoryViaEws(targetItemId, categoryName);
              }

              if (!targetItemId || !data.afterFiling || data.afterFiling === "none" || data.afterFiling === "add_date") {
                return;
              }

              await executeAfterFilingMoveByItemId(targetItemId, data.afterFiling, {
                targetFolderName: "Filed",
                filedFolderPrefix: "*",
              });
            };

            runFallback().catch((err) => {
              reportActionError(formatAfterFilingApiError(err, "Post-filing action", targetItemId));
            });
          }
        } catch (e) {
          reportActionError("Error processing dialog message: " + e.message);
          console.error(e);
        }
      });
    }
  );
}



function optionsAction(event) {
  try { localStorage.setItem("koyomail_dialog_mode", "options"); } catch(e) {}
  const dialogUrl = `${window.location.origin}/taskpane.html?mode=options`;
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 75, width: 75, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Options dialog failed to open: " + asyncResult.error.message);
        if (event && event.completed) event.completed();
        return;
      }
      
      const optionsDialog = asyncResult.value;
      optionsDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") {
          optionsDialog.close();
          if (event && event.completed) event.completed();
        }
      });

      optionsDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (event && event.completed) event.completed();
      });
    }
  );
}

function helpAction(event) {
  try { localStorage.setItem("koyomail_dialog_mode", "help"); } catch(e) {}
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

async function showStatusNotification(message, event, isSuccess = false, isError = false, allowSend = true, delayMs = 0) {
  const type = isError 
    ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage 
    : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
    
  const msgObj = {
    type: type,
    message: message
  };

  if (!isError) {
    msgObj.icon = "Icon.80";
    msgObj.persistent = false;
  }

  let item = Office.context.mailbox.item;
  if (!item) {
    for (let i = 0; i < 20; i++) {
      await new Promise(resolve => setTimeout(resolve, 100));
      item = Office.context.mailbox.item;
      if (item) break;
    }
  }

  if (item) {
    item.notificationMessages.replaceAsync(
      "StatusNotification",
      msgObj,
      async (result) => {
        if (delayMs > 0) {
          await new Promise(resolve => setTimeout(resolve, delayMs));
        }
        if (event && event.completed) {
          event.completed({ allowEvent: allowSend });
        }
      }
    );
  } else {
    console.warn("Failed to show notification because item is undefined:", message);
    if (delayMs > 0) {
      await new Promise(resolve => setTimeout(resolve, delayMs));
    }
    if (event && event.completed) {
      event.completed({ allowEvent: allowSend });
    }
  }
}

function collectionsAction(event) {
  try { localStorage.setItem("koyomail_dialog_mode", "locations"); } catch(e) {}
  const dialogUrl = `${window.location.origin}/taskpane.html?mode=locations`;
  
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 75, width: 75, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Collections dialog failed to open: " + asyncResult.error.message);
        if (event && event.completed) event.completed();
        return;
      }
      
      const collectionsDialog = asyncResult.value;
      collectionsDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") {
          collectionsDialog.close();
          if (event && event.completed) event.completed();
        }
      });

      collectionsDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (event && event.completed) event.completed();
      });
    }
  );
}

Office.actions.associate("searchAction", searchAction);
Office.actions.associate("optionsAction", optionsAction);

Office.actions.associate("helpAction", helpAction);
Office.actions.associate("openFilingDialogAction", openFilingDialogAction);
Office.actions.associate("collectionsAction", collectionsAction);

let pendingOnSendEvent = null;

function onMessageSendHandler(event) {
  pendingOnSendEvent = event;
  
  try {
    let baseUrl = window.location.origin;
    if (!baseUrl) {
      baseUrl = window.location.protocol + "//" + window.location.host;
    }
    try { localStorage.setItem("koyomail_dialog_mode", "onsend"); } catch(e) {}
    const dialogUrl = `${baseUrl}/taskpane.html?mode=onsend`;
    
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 75, width: 75, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("On-Send dialog failed to open: " + asyncResult.error.message);
        if (pendingOnSendEvent) {
          pendingOnSendEvent.completed({ allowEvent: true });
          pendingOnSendEvent = null;
        }
        return;
      }
      
      const onSendDialog = asyncResult.value;
      onSendDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "allowSend") {
          onSendDialog.close();
          if (pendingOnSendEvent) {
            pendingOnSendEvent.completed({ allowEvent: true });
            pendingOnSendEvent = null;
          }
        } else if (arg.message === "cancelSend") {
          onSendDialog.close();
          if (pendingOnSendEvent) {
            pendingOnSendEvent.completed({ allowEvent: false });
            pendingOnSendEvent = null;
          }
        } else if (arg.message.startsWith("fileEmail:")) {
          onSendDialog.close();
          try {
            const data = JSON.parse(arg.message.substring(10));
            buildCurrentEmailPayload().then(payload => {
              if (payload) {
                // Build the full payload forwarding the SSO token from the On-Send dialog.
                // We intentionally skip client-side item.categories / item.subject calls here.
                // During the ItemSend event New Outlook freezes the compose item and ALL such
                // calls fail with Error Code 5000. The backend will instead apply the category
                // to the Sent Items copy via Microsoft Graph after a short delay.
                
                // Read koyoOptions so the backend knows what categories/actions to apply.
                let koyoOpts = {};
                try {
                  const optsStr = localStorage.getItem("koyomail_options") || localStorage.getItem("koyoOptions");
                  koyoOpts = optsStr ? JSON.parse(optsStr) : {};
                } catch (e) { /* ignore */ }

                const finalPayload = {
                  ...payload,
                  targetPaths: data.paths,
                  subject: data.subject || payload.subject,
                  comment: data.comment || "",
                  attachmentsOption: data.attachmentsOption || "all",
                  markReviewed: data.markReviewed || false,
                  sendLink: data.sendLink || false,
                  isOnSend: true,
                  ssoToken: data.ssoToken || payload.ssoToken || null,
                  // Prioritize options sent from the front-end dialog, fall back to local storage
                  addFiledCategory: data.addFiledCategory !== undefined ? data.addFiledCategory : (koyoOpts.addFiledCategory !== false),
                  filedCategoryName: data.filedCategoryName || koyoOpts.filedCategoryName || "Filed by mailmanager (koyomail)",
                  afterFiling: data.afterFiling || koyoOpts.afterFilingAction || "none",
                  useUtcTime: data.useUtcTime !== undefined ? data.useUtcTime : !!koyoOpts.useUtcTime,
                  assistantCategories: data.assistantCategories || koyoOpts.assistantCategories || ""
                };

                remoteLog("info", `[commands] On-Send filing payload ready. ssoToken present: ${!!finalPayload.ssoToken}, subject: "${finalPayload.subject}", addFiledCategory: ${finalPayload.addFiledCategory}, catName: "${finalPayload.filedCategoryName}"`);

                fileEmail(finalPayload).then((response) => {
                  const isFullySkipped = response && response.results && response.results.length > 0 && response.results.every(r => r.status === "skipped");
                  const isPartiallySkipped = response && response.results && response.results.some(r => r.status === "skipped") && response.results.some(r => r.status !== "skipped");

                  if (isFullySkipped) {
                    const skippedActionsMsg = (finalPayload.afterFiling !== "none" || finalPayload.markReviewed) ? " (Post-filing actions skipped)." : "";
                    const msg = `This email is already filed.${skippedActionsMsg}`;
                    showStatusNotification(msg, pendingOnSendEvent, false, false, true, 3000);
                    pendingOnSendEvent = null;
                  } else if (isPartiallySkipped) {
                    const msg = `Email filed to new locations (already filed in some).`;
                    showStatusNotification(msg, pendingOnSendEvent, true, false, true, 3000);
                    pendingOnSendEvent = null;
                  } else {
                    if (pendingOnSendEvent) {
                      pendingOnSendEvent.completed({ allowEvent: true });
                      pendingOnSendEvent = null;
                    }
                  }
                }).catch(err => {
                  console.error("Filing failed during On-Send:", err);
                  remoteLog("error", `Filing failed during On-Send: ${err.message}`);
                  showStatusNotification(`Filing failed: ${err.message}`, pendingOnSendEvent, false, true, true, 3000);
                  pendingOnSendEvent = null;
                });
              } else {
                console.error("Payload missing during On-Send");
                remoteLog("error", "Payload missing during On-Send");
                if (pendingOnSendEvent) {
                  pendingOnSendEvent.completed({ allowEvent: true });
                  pendingOnSendEvent = null;
                }
              }
            }).catch(err => {
              console.error("Payload extraction failed during On-Send:", err);
              remoteLog("error", `Payload extraction failed during On-Send: ${err.message}`);
              if (pendingOnSendEvent) {
                pendingOnSendEvent.completed({ allowEvent: true });
                pendingOnSendEvent = null;
              }
            });
          } catch (syncErr) {
            console.error("Synchronous error processing fileEmail message:", syncErr);
            try { remoteLog("error", `Synchronous error in On-Send handler: ${syncErr.message || syncErr}`); } catch(e) {}
            if (pendingOnSendEvent) {
              pendingOnSendEvent.completed({ allowEvent: true });
              pendingOnSendEvent = null;
            }
          }
        }
      });

      onSendDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (pendingOnSendEvent) {
          pendingOnSendEvent.completed({ allowEvent: true });
          pendingOnSendEvent = null;
        }
      });
    }
  );
  } catch (err) {
    console.error("Synchronous error opening On-Send dialog: ", err);
    if (pendingOnSendEvent) {
      pendingOnSendEvent.completed({ allowEvent: true });
      pendingOnSendEvent = null;
    }
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
