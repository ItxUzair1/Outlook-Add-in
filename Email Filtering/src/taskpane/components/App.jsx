import * as React from "react";
import PropTypes from "prop-types";
import { 
  addLocation, 
  deleteLocation, 
  fileEmail, 
  getLocations, 
  updateLocation,
  getConnectivityStatus,
  exploreLocation,
  removeSuggestion,
  toggleSuggestion,
  markLocationUnused,
} from "../services/backendApi";
import { buildCurrentEmailPayload } from "../services/mailboxService";
import Toolbar from "./Toolbar";
import DetailsSidebar from "./DetailsSidebar";
import LocationTable from "./LocationTable";
import LocationDialog from "./LocationDialog";
import HelpDialog from "./HelpDialog";
import SearchDialog from "./SearchDialog";
import { Button } from "@fluentui/react-components";
import { useMsal } from "@azure/msal-react";
import { getGraphToken } from "../utils/authManager";

/* global Office */

const App = ({ title }) => {
  const { instance } = useMsal();
  // Auth tier label shown in the status bar
  const [authTier, setAuthTier] = React.useState("");
  const autoAuthTriggeredRef = React.useRef(false);
  const [locations, setLocations] = React.useState([]);
  const [selectedIds, setSelectedIds] = React.useState([]);
  const [isMultiSelect, setIsMultiSelect] = React.useState(false);
  const [connectivityStatus, setConnectivityStatus] = React.useState({});

  const getSavedDefault = (key, fallback) => {
    try {
      const saved = localStorage.getItem(`koyomail_default_${key}`);
      return saved !== null ? JSON.parse(saved) : fallback;
    } catch {
      return fallback;
    }
  };

  // Filing Options State
  const [subject, setSubject] = React.useState("");
  const [comment, setComment] = React.useState("");
  const [afterFiling, setAfterFiling] = React.useState(() => getSavedDefault("afterFiling", "none"));
  const [markReviewed, setMarkReviewed] = React.useState(() => getSavedDefault("markReviewed", false));
  const [sendLink, setSendLink] = React.useState(() => getSavedDefault("sendLink", false));
  const [attachmentsOption, setAttachmentsOption] = React.useState(() => getSavedDefault("attachmentsOption", "all"));
  const [emailPayload, setEmailPayload] = React.useState(null);

  const [loading, setLoading] = React.useState(false);
  const [message, setMessage] = React.useState("");
  const [actionError, setActionError] = React.useState("");
  const [ssoWarning, setSsoWarning] = React.useState("");
  const [graphAuthStatus, setGraphAuthStatus] = React.useState("Checking authentication...");
  const [graphAuthOk, setGraphAuthOk] = React.useState(false);

  const saveDefaults = React.useCallback(() => {
    try {
      localStorage.setItem("koyomail_default_afterFiling", JSON.stringify(afterFiling));
      localStorage.setItem("koyomail_default_markReviewed", JSON.stringify(markReviewed));
      localStorage.setItem("koyomail_default_sendLink", JSON.stringify(sendLink));
      localStorage.setItem("koyomail_default_attachmentsOption", JSON.stringify(attachmentsOption));
      setMessage("Default options saved.");
      setTimeout(() => setMessage(""), 3000);
    } catch (e) {
      console.warn("Failed to save defaults:", e);
    }
  }, [afterFiling, markReviewed, sendLink, attachmentsOption]);

  /**
   * Unified token getter — delegates to authManager which runs the three-tier chain:
   *   Tier 1: Office SSO  → Tier 2: NAA (New Outlook)  → Tier 3: MSAL redirect (Classic)
   */
  const getToken = React.useCallback(async ({ interactive = false } = {}) => {
    const result = await getGraphToken({
      msalInstance: instance,
      interactive,
      loginHint: Office?.context?.mailbox?.userProfile?.emailAddress,
    });
    setAuthTier(result.tier);
    return result.token;
  }, [instance]);

  const toRestId = React.useCallback((itemId) => {
    try {
      if (Office?.context?.mailbox?.convertToRestId) {
        return Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
      }
    } catch (error) {
      console.warn("[App] Failed to convert to REST ID:", error.message);
    }
    return itemId;
  }, []);

  const runGraphMove = React.useCallback(async (accessToken, itemId, action) => {
    const restId = toRestId(itemId);
    const destinationId = action === "archive" ? "archive" : "deleteditems";
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}/move`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ destinationId }),
    });

    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data?.error?.message || `Graph move failed: ${response.statusText}`);
    }
  }, [toRestId]);

  const openComposeWindow = React.useCallback((links, emailSubject) => {
    if (!links || links.length === 0) return;
    
    // Formatting for high-fidelity Outlook display
    const formattedLinks = links.map(l => `- ${l}`).join("<br/>");
    const htmlBody = `
      <div style="font-family: 'Segoe UI', sans-serif; font-size: 11pt; color: #323130;">
        <p>The following email has been filed to a shared location:</p>
        <p><strong>${formattedLinks}</strong></p>
        <p><i>Generated by Koyomail</i></p>
      </div>
    `;
    
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [],
      subject: `Filed Link: ${emailSubject}`,
      htmlBody: htmlBody
    });
  }, []);
  
  // Poll for errors from the parent context (commands.js)
  React.useEffect(() => {
    if (afterFiling === "none" || !loading) return;

    const interval = setInterval(() => {
      // Check for background script heartbeat
      const heartbeat = localStorage.getItem("koyomailCommandsHeartbeat");
      if (!heartbeat || Date.now() - parseInt(heartbeat) > 5000) {
        setActionError("Warning: Background script (commands.js) does not appear to be running. Filing actions like 'Delete' may not work.");
      } else {
        setActionError(""); // Clear if alive
      }

      const stored = localStorage.getItem("koyomailActionError");
      if (stored) {
        try {
          const { message: errMsgs, timestamp } = JSON.parse(stored);
          if (Date.now() - timestamp < 30000) {
            const safeError = typeof errMsgs === "string" ? errMsgs : JSON.stringify(errMsgs);
            setActionError(safeError);
            localStorage.removeItem("koyomailActionError");
          }
        } catch (e) { /* ignore */ }
      }
    }, 1000);

    // Global error interceptors to expose raw [object Object] payload details
    window.onerror = function(message, source, lineno, colno, error) {
      try {
        const detail = error ? JSON.stringify(error, Object.getOwnPropertyNames(error)) : String(message);
        setActionError(`Global Error: ${detail}`);
      } catch (e) {
        setActionError(`Global Error: ${String(message)}`);
      }
    };
    window.onunhandledrejection = function(event) {
      try {
        const detail = event.reason ? JSON.stringify(event.reason, Object.getOwnPropertyNames(event.reason)) : "Unknown rejection";
        setActionError(`Unhandled Promise: ${detail}`);
      } catch (e) {
        setActionError(`Unhandled Promise: ${String(event.reason)}`);
      }
    };

    return () => {
      clearInterval(interval);
      window.onerror = null;
      window.onunhandledrejection = null;
    };
  }, [afterFiling, loading]);
  
  // Dialog State
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [isHelpOpen, setIsHelpOpen] = React.useState(false);
  const [isSearchOpen, setIsSearchOpen] = React.useState(false);
  const [editingLocation, setEditingLocation] = React.useState(null);

  const loadLocations = React.useCallback(async () => {
    try {
      const rows = await getLocations();
      setLocations(rows);
      const status = await getConnectivityStatus();
      setConnectivityStatus(status);
    } catch (error) {
      console.error("[App] Load failed:", error);
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Load failed: ${errorMsg}`);
    }
  }, []);

  React.useEffect(() => {
    loadLocations();

    // Fetch email metadata once on mount and persist in state (skip if in help mode)
    const mode = new URLSearchParams(window.location.search).get("mode");
    if (mode === "help") {
      setIsHelpOpen(true);
      return;
    }
    if (mode === "search") {
      setIsSearchOpen(true);
      return;
    }

    const fetchData = async () => {
      try {
        const payload = await buildCurrentEmailPayload();
        if (payload) {
          setEmailPayload(payload);
          setSubject(payload.subject || "");

          // Do not show SSO warnings until full payload is available.
          if (!payload.isPartial) {
            if (payload.ssoTokenError) {
              setSsoWarning(`⚠️ SSO Authentication Warning: ${payload.ssoTokenError}. The add-in will use MSAL fallback automatically when needed.`);
            } else if (!payload.ssoToken) {
              setSsoWarning("⚠️ SSO token not available. The add-in will try MSAL fallback automatically for Graph operations.");
            } else {
              setSsoWarning("");
            }
          }
          
          // If the payload is partial, poll for the full enrichment from background
          if (payload.isPartial) {
            console.log("[App] Partial data found, polling for full enrichment...");
            const pollInterval = setInterval(async () => {
              try {
                const enriched = await buildCurrentEmailPayload();
                if (enriched && !enriched.isPartial) {
                  console.log("[App] Full enrichment received (Body & Attachments).");
                  setEmailPayload(enriched);

                  if (enriched.ssoTokenError) {
                    setSsoWarning(`⚠️ SSO Authentication Warning: ${enriched.ssoTokenError}. The add-in will use MSAL fallback automatically when needed.`);
                  } else if (!enriched.ssoToken) {
                    setSsoWarning("⚠️ SSO token not available. The add-in will try MSAL fallback automatically for Graph operations.");
                  } else {
                    setSsoWarning("");
                  }

                  clearInterval(pollInterval);
                }
              } catch (pollErr) {
                console.warn("[App] Polling enrichment failed:", pollErr.message);
                // Keep polling or clear if error is fatal
              }
            }, 1000);
            
            // Stop polling after 15 seconds to prevent memory leak
            setTimeout(() => clearInterval(pollInterval), 15000);
          }
        }
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : (typeof err === "object" ? JSON.stringify(err) : String(err));
        console.warn("[App] Initial data gathering failed:", errorMsg);
        setMessage(`Initial load failed: ${errorMsg}`);
      }
    };

    fetchData();
  }, [loadLocations]);

  // ── Auto-authentication on load ─────────────────────────────────────────────
  React.useEffect(() => {
    if (autoAuthTriggeredRef.current) return;
    autoAuthTriggeredRef.current = true;

    const autoAuthenticate = async () => {
      try {
        setGraphAuthStatus("Authenticating...");
        // Silent-only on startup — do not redirect automatically on first load
        const token = await getToken({ interactive: false });
        if (token) {
          setGraphAuthOk(true);
          setGraphAuthStatus(`Signed in ✓`);
        }
      } catch {
        // Silent auth failed — show Sign In button, do not auto-redirect
        setGraphAuthOk(false);
        setGraphAuthStatus("Sign in required");
      }
    };

    autoAuthenticate();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const onSelectionChange = (id) => {
    setSelectedIds((prev) => {
      if (isMultiSelect) {
        return prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id];
      } else {
        return prev.includes(id) ? [] : [id];
      }
    });
  };

  const onSaveLocation = async (data) => {
    try {
      if (editingLocation) {
        await updateLocation(editingLocation.id, data);
        setMessage("Location updated.");
      } else {
        await addLocation(data);
        setMessage("Location added.");
      }
      await loadLocations();
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Save failed: ${errorMsg}`);
    }
  };

  const onDeleteLocation = async () => {
    if (selectedIds.length === 0) {
      setMessage("Please select at least one location to delete.");
      return;
    }

    try {
      setLoading(true);
      setMessage(`Deleting ${selectedIds.length} location(s)...`);
      for (const id of selectedIds) {
        await deleteLocation(id);
      }
      setSelectedIds([]);
      setMessage("Location(s) deleted successfully.");
      await loadLocations();
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      console.error("Delete failed:", error);
      setMessage(`Delete failed: ${errorMsg}`);
    } finally {
      setLoading(false);
    }
  };

  const onFileEmail = async () => {
    setLoading(true);
    setMessage("");

    try {
      const selectedLocations = locations.filter((x) => selectedIds.includes(x.id));
      if (selectedLocations.length === 0) {
        throw new Error("Select at least one target location.");
      }

      // Check connectivity for all selected locations
      const disconnected = selectedLocations.filter(loc => connectivityStatus[loc.id] === false);
      if (disconnected.length > 0) {
        const paths = disconnected.map(d => d.path.split("\\").pop()).join(", ");
        throw new Error(`Filing failed: Location(s) [${paths}] are disconnected. Please check your network connection.`);
      }

      setMessage("Reflecting latest email changes...");
      let latestPayload = await buildCurrentEmailPayload();
      let basePayload = latestPayload || emailPayload;
      if (!basePayload) {
        throw new Error("Email content is not ready yet. Please wait a moment.");
      }
      if (basePayload.isPartial) {
        const refreshedPayload = await buildCurrentEmailPayload({ forceRefresh: true });
        basePayload = refreshedPayload || basePayload;
      }
      if (basePayload.isPartial) {
        setMessage("Body enrichment is taking longer than expected. Filing with available preview content.");
      } else {
        setMessage("");
      }

      let graphAccessToken = null;
      const needsGraphPostActions = afterFiling !== "none" || markReviewed || sendLink;
      if (basePayload?.itemId && (!basePayload?.ssoToken || needsGraphPostActions)) {
        try {
          graphAccessToken = await getToken({ interactive: false });
        } catch (tokenErr) {
          // Non-fatal: if token is unavailable, we'll keep existing guard for pending attachments.
          console.warn("[App] No graph token available before attachment validation:", tokenErr?.message || tokenErr);
        }
      }

      if (attachmentsOption !== "message") {
        const attList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
        const pendingAttachments = attList.filter((att) => {
          const hasContent = !!att?.base64Content;
          const isMetadataOnly = !!att?.isMetadataOnly;
          const isInline = !!att?.isInline;
          const size = Number(att?.size || 0);
          return (isMetadataOnly || !hasContent) && !isInline && size > 0;
        });
        const hasPendingAttachments = pendingAttachments.length > 0;

        // If no Graph-capable token exists, pending attachment metadata may still be incomplete.
        if (!basePayload?.ssoToken && !graphAccessToken && hasPendingAttachments) {
          const retryPayload = await buildCurrentEmailPayload({ forceRefresh: true });
          basePayload = retryPayload || basePayload;

          const retryList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
          const retryPendingAttachments = retryList.filter((att) => {
            const hasContent = !!att?.base64Content;
            const isMetadataOnly = !!att?.isMetadataOnly;
            const isInline = !!att?.isInline;
            const size = Number(att?.size || 0);
            return (isMetadataOnly || !hasContent) && !isInline && size > 0;
          });

          if (!basePayload?.ssoToken && !graphAccessToken && retryPendingAttachments.length > 0) {
            throw new Error("Attachments are still loading. Please wait a few seconds and try again to avoid missing attachments.");
          }
        }
      }
      
      // Filter attachments based on user selection
      let finalAttachments = basePayload.attachments || [];
      if (attachmentsOption === "message") {
        finalAttachments = [];
      } else if (attachmentsOption === "attachments") {
        // Keep as is, but logic favors attachments
      }

      console.log("[App] Filing email with payload:", {
        subject,
        attachmentCount: finalAttachments.length,
        attachmentsOption,
        targetPaths: selectedLocations.map((x) => x.path)
      });

      if (!graphAccessToken && basePayload?.itemId && (!basePayload?.ssoToken || needsGraphPostActions)) {
        try {
          graphAccessToken = await getToken({ interactive: false });
        } catch (tokenErr) {
          // Non-fatal: backend can still use frontend attachment payload fallback.
          console.warn("[App] No graph token available for backend enrichment:", tokenErr?.message || tokenErr);
        }
      }

      const validatedGraphAccessToken = (typeof graphAccessToken === "string" && graphAccessToken.length > 10) 
        ? graphAccessToken 
        : null;

      const response = await fileEmail({
        ...basePayload,
        graphAccessToken: validatedGraphAccessToken,
        attachments: finalAttachments,
        subject,
        comment,
        afterFiling,
        markReviewed,
        sendLink,
        attachmentsOption,
        targetPaths: selectedLocations.map((x) => x.path),
      });

      // Check for post-filing errors returned from backend
      if (response?.postFilingError) {
        setActionError(response.postFilingError);
        setMessage("Email filed successfully, but post-filing action failed.");
      } else {
        setMessage("Email filed successfully.");
      }
      
      // If generate link was requested and we have shared paths, open compose window
      if (sendLink && response?.sharingLinks?.length > 0) {
        openComposeWindow(response.sharingLinks, subject);
      }
      
      // Perform after-filing actions
      const item = Office.context?.mailbox?.item;
      if (afterFiling !== "none" && !basePayload?.ssoToken) {
        const graphItemId = basePayload?.itemId || emailPayload?.itemId || item?.itemId;
        if (graphItemId) {
          try {
            const token = await getToken({ interactive: false });
            await runGraphMove(token, graphItemId, afterFiling);
            setMessage(`Email filed and ${afterFiling === "delete" ? "moved to Deleted Items" : "Archived"} via Microsoft Graph.`);
            return;
          } catch (msalPostActionErr) {
            const errMsg = msalPostActionErr instanceof Error ? msalPostActionErr.message : String(msalPostActionErr);
            setActionError(`Graph ${afterFiling} failed: ${errMsg}`);
          }
        }

        // Do not use removeAsync for delete fallback here: in some Outlook hosts it can hard-delete.
        if (item && afterFiling === "delete") {
          setActionError("Automatic local delete was skipped to prevent permanent deletion in this Outlook host.");
          setMessage("Email filed successfully. Please move the email to Deleted Items manually.");
          return;
        }

        if (item && afterFiling === "archive") {
          if (item.archiveAsync) {
            item.archiveAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Failed) {
                setMessage("Email filed, but failed to Archive: " + (result.error?.message || "Unknown error"));
              } else {
                setMessage("Email filed and Archived.");
              }
            });
          } else {
            setMessage("Email filed, but 'Archive' action is not supported in this version of Outlook.");
          }
          return;
        }
      } else if (afterFiling !== "none" && basePayload?.ssoToken) {
        setMessage(`Email filed and ${afterFiling === "delete" ? "moved to Deleted Items" : "Archived"} via Microsoft Graph.`);
      } else if (afterFiling !== "none") {
        // We are likely in a dialog, message the parent to handle the action
        if (Office.context.ui && Office.context.ui.messageParent) {
          setMessage(`Email filed. Requesting Outlook to ${afterFiling === "delete" ? "move to Deleted Items" : "Archive"}...`);
          Office.context.ui.messageParent(JSON.stringify({ action: "afterFiling", value: afterFiling }));
          
          // Wait up to 10 seconds for the parent to close the dialog
          let secondsPassed = 0;
          while (secondsPassed < 10) {
            await new Promise(resolve => setTimeout(resolve, 1000));
            secondsPassed++;
            
            // Check for errors reported by the parent
            const storedError = localStorage.getItem("koyomailActionError");
            if (storedError) {
              const { message: parentError } = JSON.parse(storedError);
              localStorage.removeItem("koyomailActionError");
              // Filing already succeeded; treat post-filing move/archive issues as warnings.
              setActionError(parentError);
              setMessage("Email filed successfully. Automatic move/archive could not be completed in this Outlook host.");
              return;
            }
          }
          
          setMessage(`Filing complete, but Outlook is taking longer than expected to ${afterFiling === "delete" ? "delete" : "archive"} the email. You may close this window manually.`);
        } else {
          setMessage("Email filed, but could not request move/archive (parent context not found).");
        }
      }

    } catch (error) {
      console.error("[App] Filing failed:", error);
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Filing failed: ${errorMsg}`);
    } finally {
      setLoading(false);
    }
  };

  const onFileToPath = async (targetPath) => {
    setLoading(true);
    setMessage("");

    try {
      // Check connectivity for the target path
      const loc = locations.find(x => x.path === targetPath);
      if (loc && connectivityStatus[loc.id] === false) {
        throw new Error(`Filing failed: Location is disconnected. Please check your network connection.`);
      }

      setMessage("Reflecting latest email changes...");
      let latestPayload = await buildCurrentEmailPayload();
      let basePayload = latestPayload || emailPayload;
      if (!basePayload) {
        throw new Error("Email content is not ready yet. Please wait a moment.");
      }
      if (basePayload.isPartial) {
        const refreshedPayload = await buildCurrentEmailPayload({ forceRefresh: true });
        basePayload = refreshedPayload || basePayload;
      }
      if (basePayload.isPartial) {
        setMessage("Body enrichment is taking longer than expected. Filing with available preview content.");
      } else {
        setMessage("");
      }

      let graphAccessToken = null;
      const needsGraphPostActions = afterFiling !== "none" || markReviewed || sendLink;
      if (basePayload?.itemId && (!basePayload?.ssoToken || needsGraphPostActions)) {
        try {
          graphAccessToken = await getToken({ interactive: false });
        } catch (tokenErr) {
          console.warn("[App] No graph token available before attachment validation:", tokenErr?.message || tokenErr);
        }
      }

      if (attachmentsOption !== "message") {
        const attList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
        const pendingAttachments = attList.filter((att) => {
          const hasContent = !!att?.base64Content;
          const isMetadataOnly = !!att?.isMetadataOnly;
          const isInline = !!att?.isInline;
          const size = Number(att?.size || 0);
          return (isMetadataOnly || !hasContent) && !isInline && size > 0;
        });
        const hasPendingAttachments = pendingAttachments.length > 0;

        if (!basePayload?.ssoToken && !graphAccessToken && hasPendingAttachments) {
          const retryPayload = await buildCurrentEmailPayload({ forceRefresh: true });
          basePayload = retryPayload || basePayload;

          const retryList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
          const retryPendingAttachments = retryList.filter((att) => {
            const hasContent = !!att?.base64Content;
            const isMetadataOnly = !!att?.isMetadataOnly;
            const isInline = !!att?.isInline;
            const size = Number(att?.size || 0);
            return (isMetadataOnly || !hasContent) && !isInline && size > 0;
          });

          if (!basePayload?.ssoToken && !graphAccessToken && retryPendingAttachments.length > 0) {
            throw new Error("Attachments are still loading. Please wait a few seconds and try again to avoid missing attachments.");
          }
        }
      }

      // Filter attachments based on user selection
      let finalAttachments = basePayload.attachments || [];
      if (attachmentsOption === "message") {
        finalAttachments = [];
      }

      if (!graphAccessToken && basePayload?.itemId && (!basePayload?.ssoToken || needsGraphPostActions)) {
        try {
          graphAccessToken = await getToken({ interactive: false });
        } catch (tokenErr) {
          console.warn("[App] No graph token available for backend enrichment:", tokenErr?.message || tokenErr);
        }
      }

      const validatedGraphAccessToken = (typeof graphAccessToken === "string" && graphAccessToken.length > 10) 
        ? graphAccessToken 
        : null;

      const response = await fileEmail({
        ...basePayload,
        graphAccessToken: validatedGraphAccessToken,
        attachments: finalAttachments,
        subject,
        comment,
        afterFiling,
        markReviewed,
        sendLink,
        attachmentsOption,
        targetPaths: [targetPath],
      });

      // If generate link was requested and we have shared paths, open compose window
      if (sendLink && response?.sharingLinks?.length > 0) {
        openComposeWindow(response.sharingLinks, subject);
      }

      const item = Office.context?.mailbox?.item;
      if (afterFiling !== "none" && !basePayload?.ssoToken) {
        const graphItemId = basePayload?.itemId || emailPayload?.itemId || item?.itemId;
        if (graphItemId) {
          try {
            const token = await getToken({ interactive: false });
            await runGraphMove(token, graphItemId, afterFiling);
            setMessage(`Email filed and ${afterFiling === "delete" ? "moved to Deleted Items" : "Archived"} via Microsoft Graph.`);
            await loadLocations();
            return;
          } catch (msalPostActionErr) {
            const errMsg = msalPostActionErr instanceof Error ? msalPostActionErr.message : String(msalPostActionErr);
            setActionError(`Graph ${afterFiling} failed: ${errMsg}`);
          }
        }

        if (item && afterFiling === "delete") {
          setActionError("Automatic local delete was skipped to prevent permanent deletion in this Outlook host.");
          setMessage("Email filed successfully. Please move the email to Deleted Items manually.");
          await loadLocations();
          return;
        }

        if (item && afterFiling === "archive") {
          if (item.archiveAsync) {
            item.archiveAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Failed) {
                setMessage("Email filed, but failed to Archive: " + (result.error?.message || "Unknown error"));
              } else {
                setMessage("Email filed and Archived.");
              }
            });
          } else {
            setMessage("Email filed, but 'Archive' action is not supported in this version of Outlook.");
          }
          await loadLocations();
          return;
        }
      } else if (afterFiling !== "none" && basePayload?.ssoToken) {
        setMessage(`Email filed and ${afterFiling === "delete" ? "moved to Deleted Items" : "Archived"} via Microsoft Graph.`);
      } else if (afterFiling !== "none") {
        if (Office.context.ui && Office.context.ui.messageParent) {
          setMessage(`Email filed. Requesting Outlook to ${afterFiling === "delete" ? "move to Deleted Items" : "Archive"}...`);
          Office.context.ui.messageParent(JSON.stringify({ action: "afterFiling", value: afterFiling }));
          
          let secondsPassed = 0;
          while (secondsPassed < 10) {
            await new Promise(resolve => setTimeout(resolve, 1000));
            secondsPassed++;
            const storedError = localStorage.getItem("koyomailActionError");
            if (storedError) {
              const { message: parentError } = JSON.parse(storedError);
              localStorage.removeItem("koyomailActionError");
              setActionError(parentError);
              setMessage("Email filed successfully. Automatic move/archive could not be completed in this Outlook host.");
              await loadLocations();
              return;
            }
          }
          setMessage(`Filing complete, but Outlook is taking longer than expected to ${afterFiling === "delete" ? "delete" : "archive"} the email. You may close this window manually.`);
        } else {
          setMessage("Email filed, but could not request move/archive (parent context not found).");
        }
      }
      
      await loadLocations(); // Refresh to update lastUsedAt
    } catch (error) {
      console.error("[App] Filing to path failed:", error);
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Filing failed: ${errorMsg}`);
    } finally {
      setLoading(false);
    }
  };

  const onExplore = async () => {
    if (selectedIds.length !== 1) {
      setMessage("Please select exactly one location to explore.");
      return;
    }
    const loc = locations.find((x) => x.id === selectedIds[0]);
    if (loc) {
      try {
        await exploreLocation(loc.path);
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
        setMessage(`Explore failed: ${errorMsg}`);
      }
    }
  };

  const onRemoveSuggestion = async () => {
    if (selectedIds.length === 0) {
      setMessage("Please select at least one suggestion to remove.");
      return;
    }
    try {
      for (const id of selectedIds) {
        await removeSuggestion(id);
      }
      setSelectedIds([]);
      setMessage("Suggestions removed.");
      await loadLocations();
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Remove failed: ${errorMsg}`);
    }
  };

  const onToggleSuggestion = async (id) => {
    try {
      await toggleSuggestion(id);
      await loadLocations();
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Toggle failed: ${errorMsg}`);
    }
  };

  const onMarkUnused = async () => {
    if (selectedIds.length === 0) {
      setMessage("Please select at least one location to mark as unused.");
      return;
    }
    try {
      for (const id of selectedIds) {
        await markLocationUnused(id);
      }
      setMessage("Locations marked as unused.");
      await loadLocations();
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Mark failed: ${errorMsg}`);
    }
  };

  return (
    <div style={{ height: "100vh", display: "flex", flexDirection: "column", fontFamily: "Segoe UI" }}>
      <Toolbar 
        locations={locations}
        onFileToPath={onFileToPath}
        onAdd={() => { setEditingLocation(null); setIsDialogOpen(true); }}
        onEdit={() => {
          if (selectedIds.length === 1) {
            setEditingLocation(locations.find(x => x.id === selectedIds[0]));
            setIsDialogOpen(true);
          }
        }}
        onExplore={onExplore}
        onRefresh={loadLocations}
        onRemoveSuggestion={onRemoveSuggestion}
        onMarkUnused={onMarkUnused}
        onToggleMultiSelect={() => setIsMultiSelect(!isMultiSelect)}
        isMultiSelect={isMultiSelect}
        onDelete={onDeleteLocation}
        onHelp={() => setIsHelpOpen(true)}
      />

      <div style={{ display: "flex", flexWrap: "nowrap", flexGrow: 1, overflow: "hidden" }}>
        <div style={{ flex: "1 1 auto", minWidth: 280, padding: 8, height: "100%", boxSizing: "border-box" }}>
          <LocationTable 
            locations={locations}
            selectedIds={selectedIds}
            onSelectionChange={onSelectionChange}
            connectivityStatus={connectivityStatus}
            onToggleSuggestion={onToggleSuggestion}
          />
        </div>

        <DetailsSidebar 
          subject={subject} setSubject={setSubject}
          comment={comment} setComment={setComment}
          afterFiling={afterFiling} setAfterFiling={setAfterFiling}
          markReviewed={markReviewed} setMarkReviewed={setMarkReviewed}
          sendLink={sendLink} setSendLink={setSendLink}
          attachmentsOption={attachmentsOption} setAttachmentsOption={setAttachmentsOption}
          onSaveDefaults={saveDefaults}
        />
      </div>

      <div style={{ padding: 12, borderTop: "1px solid #edebe9", display: "flex", flexDirection: "column", gap: 8, backgroundColor: "#f3f2f1" }}>
        <div style={{ 
          fontSize: 13, 
          color: graphAuthOk ? "#107c10" : "#8a6d00", 
          backgroundColor: graphAuthOk ? "#e8f5e8" : "#fff4ce", 
          padding: "4px 8px", 
          borderRadius: 4,
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          minHeight: "32px"
        }}>
          <span>{graphAuthStatus}</span>
          {!graphAuthOk && !graphAuthStatus.includes("✓") && !graphAuthStatus.includes("Authenticating") && (
            <Button 
              size="small"
              appearance="primary"
              onClick={() => {
                setGraphAuthStatus("Signing in...");
                getToken({ interactive: true })
                  .then(() => {
                    setGraphAuthOk(true);
                    setGraphAuthStatus("Signed in ✓");
                  })
                  .catch(err => {
                    // If it's a redirect error, the page navigates — don't update state
                    if (!err.message?.includes("Redirecting")) {
                      setGraphAuthStatus(`Sign in failed: ${err.message}`);
                    }
                  });
              }}
              style={{ padding: "0 8px", height: "24px", minWidth: "auto" }}
            >
              Sign In
            </Button>
          )}
        </div>
        
        {ssoWarning && !graphAuthOk && <div style={{ fontSize: 13, color: "#7f6700", backgroundColor: "#fef3cd", padding: "4px 8px", borderRadius: 4 }}>{ssoWarning}</div>}
        {actionError && <div style={{ fontSize: 13, color: "#a4262c", backgroundColor: "#fde7e9", padding: "4px 8px", borderRadius: 4 }}>{actionError}</div>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          {message && <span style={{ flexGrow: 1, alignSelf: "center", fontSize: 13, color: message.includes("failed") ? "#a4262c" : "#107c10" }}>{message}</span>}
          <Button appearance="primary" style={{ width: 80 }} onClick={onFileEmail} disabled={loading || selectedIds.length === 0}>
            {loading ? "Filing..." : "File"}
          </Button>
          <Button style={{ width: 80, border: "1px solid #c8c6c4" }} onClick={() => {
            if (Office.context.ui && Office.context.ui.messageParent) {
              Office.context.ui.messageParent("close");
            } else {
              window.close();
            }
          }}>Cancel</Button>
        </div>
      </div>

      <LocationDialog 
        isOpen={isDialogOpen}
        onOpenChange={setIsDialogOpen}
        onSave={onSaveLocation}
        initialData={editingLocation}
      />

      <HelpDialog 
        isOpen={isHelpOpen}
        onOpenChange={(isOpen) => {
          setIsHelpOpen(isOpen);
          const urlParams = new URLSearchParams(window.location.search);
          if (!isOpen && urlParams.get("mode") === "help") {
            if (Office.context.ui && Office.context.ui.messageParent) {
              Office.context.ui.messageParent("close");
            } else {
              window.close();
            }
          }
        }}
      />

      {/* Full-screen Search mode — rendered as a dialog from the ribbon */}
      {isSearchOpen && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 9999,
          backgroundColor: "#f8f8f8",
        }}>
          <SearchDialog
            onClose={() => {
              if (Office.context.ui && Office.context.ui.messageParent) {
                Office.context.ui.messageParent("close");
              } else {
                setIsSearchOpen(false);
              }
            }}
          />
        </div>
      )}
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
