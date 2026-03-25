/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

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

function openFilingDialogAction(event) {
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

function personalAction(event) {
  showMilestoneNotification(event, "Personal filing location");
}

function statusAction(event) {
  showMilestoneNotification(event, "Connectivity Status: All locations are currently connected.");
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
Office.actions.associate("personalAction", personalAction);
Office.actions.associate("statusAction", statusAction);
Office.actions.associate("labelAction", labelAction);
Office.actions.associate("toolsAction", toolsAction);
Office.actions.associate("helpAction", helpAction);
Office.actions.associate("openFilingDialogAction", openFilingDialogAction);
