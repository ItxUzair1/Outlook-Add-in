/* global Office console */

export async function insertText(text) {
  // Write text to the cursor point in the compose surface.
  try {
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item?.body.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Text },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            const msg = asyncResult?.error?.message;
            reject(new Error(typeof msg === "string" ? msg : "Failed to insert text."));
            return;
          }
          resolve();
        }
      );
    });
  } catch (error) {
    console.log("Error:", error);
  }
}
