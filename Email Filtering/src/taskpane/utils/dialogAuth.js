/* global Office */

const DIALOG_ERRORS = {
  12002: "The dialog URL is invalid.",
  12003: "The dialog URL uses an unsupported protocol. HTTPS is required.",
  12006: "Login was cancelled.",
  12009: "A dialog is already open in this add-in.",
};

export function openAuthDialogAndGetToken(dialogUrl, options = {}) {
  const { timeoutMs = 120000 } = options;

  return new Promise((resolve, reject) => {
    if (!Office?.context?.ui?.displayDialogAsync) {
      reject(new Error("Office Dialog API is not available in this host."));
      return;
    }

    let settled = false;
    let dialogRef = null;

    const finish = (handler, payload) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      try {
        dialogRef?.close();
      } catch (closeErr) {
        console.warn("[dialogAuth] Failed to close dialog:", closeErr);
      }
      handler(payload);
    };

    const timer = setTimeout(() => {
      finish(reject, new Error("Authentication timed out before completing sign-in."));
    }, timeoutMs);

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      {
        height: 60,
        width: 40,
        promptBeforeOpen: false,
        displayInIframe: false,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          const message = result.error?.message || "Unable to open the authentication dialog.";
          finish(reject, new Error(message));
          return;
        }

        dialogRef = result.value;

        dialogRef.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          let message;
          try {
            message = JSON.parse(args.message);
          } catch (err) {
            finish(reject, new Error("Authentication dialog returned an invalid message."));
            return;
          }

          if (message?.type === "token" && message?.accessToken) {
            finish(resolve, {
              accessToken: message.accessToken,
              account: message.account || null,
              expiresOn: message.expiresOn || null,
            });
            return;
          }

          if (message?.type === "error") {
            finish(reject, new Error(message.message || "Authentication failed in dialog."));
            return;
          }

          finish(reject, new Error("Authentication dialog returned an unexpected response."));
        });

        dialogRef.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
          const mapped = DIALOG_ERRORS[args.error];
          finish(reject, new Error(mapped || `Dialog error: ${args.error}`));
        });
      }
    );
  });
}
