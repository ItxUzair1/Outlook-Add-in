async function openFilingDialogAction(event) {
  // Clear any existing stale payload
  localStorage.removeItem("currentEmailPayload");
  localStorage.removeItem("multiEmailPayload");

  const isMultiSelectSupported = Office.context.requirements.isSetSupported("Mailbox", "1.13") && !!Office.context.mailbox.getSelectedItemsAsync;

  if (isMultiSelectSupported) {
    Office.context.mailbox.getSelectedItemsAsync(async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 1) {
        await handleMultiSelectFiling(result.value, event);
      } else {
        await handleSingleSelectFiling(event);
      }
    });
  } else {
    await handleSingleSelectFiling(event);
  }
}

async function handleMultiSelectFiling(selectedItems, event) {
  console.log(`[commands] Handling multi-select filing for ${selectedItems.length} items`);
  
  // Save basic metadata for all selected items
  localStorage.setItem("multiEmailPayload", JSON.stringify({
    items: selectedItems,
    timestamp: Date.now()
  }));

  const dialogUrl = `${window.location.origin}/taskpane.html?mode=file_multi`;
  openDialogWithHandlers(dialogUrl, event);
}

async function handleSingleSelectFiling(event) {
  console.log("[commands] Handling single-select filing");
  const { buildEmailMetadata, buildCurrentEmailPayload } = require("../taskpane/services/mailboxService");
  
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

  const dialogUrl = `${window.location.origin}/taskpane.html?mode=file`;
  openDialogWithHandlers(dialogUrl, event);
}

function openDialogWithHandlers(dialogUrl, event) {
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 55, width: 55, displayInIframe: true },
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
          }
        } catch (e) {
          reportActionError("Error processing dialog message: " + e.message);
          console.error(e);
        }
      });
    }
  );
}
