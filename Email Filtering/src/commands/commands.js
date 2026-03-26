/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { buildCurrentEmailPayload } from "../taskpane/services/mailboxService";

Office.onReady(() => {
  // Check connectivity and update ribbon on startup
  updateRibbon();
});

/**
 * Checks connectivity to all locations via the backend API.
 * @returns {Promise<boolean>} True if all locations are connected.
 */
async function checkConnectivity() {
  try {
    // Note: We use the absolute URL because commands.js runs in a background context (commands.html)
    const response = await fetch("https://localhost:3000/api/locations/status");
    if (!response.ok) return false;
    const status = await response.json();
    const values = Object.values(status);
    return values.length > 0 && values.every(v => v === true);
  } catch (error) {
    console.error("Connectivity check failed:", error);
    return false;
  }
}

/**
 * Updates the ribbon 'Status' button icon based on connectivity result.
 */
async function updateRibbon() {
  try {
    const isOk = await checkConnectivity();
    
    // Office.ribbon.requestUpdate is available in Requirement Set Ribbon 1.1+
    if (typeof Office !== "undefined" && Office.ribbon && Office.ribbon.requestUpdate) {
      await Office.ribbon.requestUpdate({
        tabs: [
          {
            id: "TabDefault",
            groups: [
              {
                id: "MailManager.Group",
                controls: [
                  {
                    id: "MailManager.Status.Button",
                    icon: isOk ? "Icon.Status.Ok" : "Icon.Status"
                  }
                ]
              }
            ]
          }
        ]
      });
    }
  } catch (error) {
    console.error("Ribbon update failed:", error);
  }
}

function showMilestoneNotification(event, featureName) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: `${featureName} is available as a button in Milestone 2 but full functionality is planned for the next milestone.`,
    icon: "Icon.80x80",
    persistent: false,
  };

  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    `${featureName}MilestoneNotification`,
    message
  );

  event.completed();
}

function searchAction(event) {
  showMilestoneNotification(event, "Search");
}

function optionsAction(event) {
  showMilestoneNotification(event, "Options");
}

let dialog;

async function openFilingDialogAction(event) {
  // Clear any existing stale payload
  localStorage.removeItem("currentEmailPayload");

  // Cache the current email payload for the dialog (which lacks mailbox access)
  try {
    const payload = await buildCurrentEmailPayload();
    const cacheData = {
      payload,
      timestamp: Date.now()
    };
    localStorage.setItem("currentEmailPayload", JSON.stringify(cacheData));
  } catch (error) {
    console.warn("Failed to cache email payload:", error);
  }

  // Use the origin of the current command to derive the dialog URL
  const dialogUrl = `${window.location.origin}/taskpane.html?mode=file`;

  // displayInIframe is needed for some environments, but 80% width/height gives a good desktop size
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 70, width: 70, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Dialog failed to open: " + asyncResult.error.message);
      } else {
        dialog = asyncResult.value;
        // Optionally handle messages from the dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          if (arg.message === "close") {
            dialog.close();
          }
        });
      }
      
      // Update ribbon status after interactive actions
      updateRibbon();

      // Complete the ribbon command action
      if (event && event.completed) {
        event.completed();
      }
    }
  );
}

function suggestedAction(event) {
  showMilestoneNotification(event, "Suggested locations");
}


async function statusAction(event) {
  await updateRibbon();
  const isOk = await checkConnectivity();
  const statusMsg = isOk 
    ? "Connectivity Status: All locations are currently connected." 
    : "Connectivity Status: Some locations are disconnected. Please check your network drives.";
  
  showMilestoneNotification(event, statusMsg);
}

function labelAction(event) {
  showMilestoneNotification(event, "Label");
}

function toolsAction(event) {
  showMilestoneNotification(event, "Tools");
}

function helpAction(event) {
  showMilestoneNotification(event, "Help");
}

Office.actions.associate("searchAction", searchAction);
Office.actions.associate("optionsAction", optionsAction); // retained just in case
Office.actions.associate("suggestedAction", suggestedAction);
Office.actions.associate("statusAction", statusAction);
Office.actions.associate("labelAction", labelAction);
Office.actions.associate("toolsAction", toolsAction);
Office.actions.associate("helpAction", helpAction);
Office.actions.associate("openFilingDialogAction", openFilingDialogAction);
