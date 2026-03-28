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
import { Button, Spinner } from "@fluentui/react-components";

/* global Office */

const App = ({ title }) => {
  const [locations, setLocations] = React.useState([]);
  const [selectedIds, setSelectedIds] = React.useState([]);
  const [isMultiSelect, setIsMultiSelect] = React.useState(false);
  const [connectivityStatus, setConnectivityStatus] = React.useState({});
  
  // Filing Options State
  const [subject, setSubject] = React.useState("");
  const [comment, setComment] = React.useState("");
  const [afterFiling, setAfterFiling] = React.useState("none");
  const [markReviewed, setMarkReviewed] = React.useState(false);
  const [sendLink, setSendLink] = React.useState(false);
  const [attachmentsOption, setAttachmentsOption] = React.useState("all");
  const [emailPayload, setEmailPayload] = React.useState(null);

  const [loading, setLoading] = React.useState(false);
  const [message, setMessage] = React.useState("");
  const [actionError, setActionError] = React.useState("");
  
  // Poll for errors from the parent context (commands.js)
  React.useEffect(() => {
    if (afterFiling === "none" || !loading) return;

    const interval = setInterval(() => {
      // Check for background script heartbeat
      const heartbeat = localStorage.getItem("mailManagerCommandsHeartbeat");
      if (!heartbeat || Date.now() - parseInt(heartbeat) > 5000) {
        setActionError("Warning: Background script (commands.js) does not appear to be running. Filing actions like 'Delete' may not work.");
      } else {
        setActionError(""); // Clear if alive
      }

      const stored = localStorage.getItem("mailManagerActionError");
      if (stored) {
        try {
          const { message: errMsgs, timestamp } = JSON.parse(stored);
          if (Date.now() - timestamp < 30000) {
            setActionError(errMsgs);
            localStorage.removeItem("mailManagerActionError");
          }
        } catch (e) { /* ignore */ }
      }
    }, 1000);

    return () => clearInterval(interval);
  }, [afterFiling, loading]);
  
  // Dialog State
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [isHelpOpen, setIsHelpOpen] = React.useState(false);
  const [editingLocation, setEditingLocation] = React.useState(null);

  const loadLocations = React.useCallback(async () => {
    try {
      const rows = await getLocations();
      setLocations(rows);
      const status = await getConnectivityStatus();
      setConnectivityStatus(status);
    } catch (error) {
      setMessage(`Load failed: ${error.message}`);
    }
  }, []);

  React.useEffect(() => {
    loadLocations();

    // Check if we should start in help mode
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get("mode") === "help") {
      setIsHelpOpen(true);
    }

    // Fetch email metadata once on mount and persist in state
    buildCurrentEmailPayload()
      .then(payload => {
        setEmailPayload(payload);
        setSubject(payload.subject || "");
      })
      .catch((err) => {
        setMessage(`Initial load failed: ${err.message}`);
      });
  }, [loadLocations]);

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
      setMessage(`Save failed: ${error.message}`);
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
      console.error("Delete failed:", error);
      setMessage(`Delete failed: ${error.message}`);
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

      const basePayload = emailPayload;
      if (!basePayload) {
        throw new Error("Email content is not ready yet. Please wait a moment.");
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

      const response = await fileEmail({
        ...basePayload,
        attachments: finalAttachments,
        subject,
        comment,
        afterFiling,
        markReviewed,
        sendLink,
        attachmentsOption,
        targetPaths: selectedLocations.map((x) => x.path),
      });

      setMessage("Email filed successfully.");
      
      // Perform after-filing actions in Outlook
      const item = Office.context?.mailbox?.item;
      if (item && afterFiling !== "none") {
        if (afterFiling === "delete") {
          item.removeAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              setMessage("Email filed, but failed to move to Deleted Items: " + (result.error?.message || "Unknown error"));
            } else {
              setMessage("Email filed and moved to Deleted Items.");
            }
          });
        } else if (afterFiling === "archive") {
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
        }
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
            const storedError = localStorage.getItem("mailManagerActionError");
            if (storedError) {
              const { message: parentError } = JSON.parse(storedError);
              localStorage.removeItem("mailManagerActionError");
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
      setMessage(`Filing failed: ${error.message}`);
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

      const basePayload = emailPayload;
      if (!basePayload) {
        throw new Error("Email content is not ready yet. Please wait a moment.");
      }

      // Filter attachments based on user selection
      let finalAttachments = basePayload.attachments || [];
      if (attachmentsOption === "message") {
        finalAttachments = [];
      }

      await fileEmail({
        ...basePayload,
        attachments: finalAttachments,
        subject,
        comment,
        afterFiling,
        markReviewed,
        sendLink,
        attachmentsOption,
        targetPaths: [targetPath],
      });

      const item = Office.context?.mailbox?.item;
      if (item && afterFiling !== "none") {
        if (afterFiling === "delete") {
          item.removeAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              setMessage("Email filed, but failed to move to Deleted Items: " + (result.error?.message || "Unknown error"));
            } else {
              setMessage("Email filed and moved to Deleted Items.");
            }
          });
        } else if (afterFiling === "archive") {
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
        }
      } else if (afterFiling !== "none") {
        if (Office.context.ui && Office.context.ui.messageParent) {
          setMessage(`Email filed. Requesting Outlook to ${afterFiling === "delete" ? "move to Deleted Items" : "Archive"}...`);
          Office.context.ui.messageParent(JSON.stringify({ action: "afterFiling", value: afterFiling }));
          
          let secondsPassed = 0;
          while (secondsPassed < 10) {
            await new Promise(resolve => setTimeout(resolve, 1000));
            secondsPassed++;
            const storedError = localStorage.getItem("mailManagerActionError");
            if (storedError) {
              const { message: parentError } = JSON.parse(storedError);
              localStorage.removeItem("mailManagerActionError");
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
      setMessage(`Filing failed: ${error.message}`);
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
        setMessage(`Explore failed: ${error.message}`);
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
      setMessage(`Remove failed: ${error.message}`);
    }
  };

  const onToggleSuggestion = async (id) => {
    try {
      await toggleSuggestion(id);
      await loadLocations();
    } catch (error) {
      setMessage(`Toggle failed: ${error.message}`);
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
      setMessage(`Mark failed: ${error.message}`);
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
        />
      </div>

      <div style={{ padding: 12, borderTop: "1px solid #edebe9", display: "flex", flexDirection: "column", gap: 8, backgroundColor: "#f3f2f1" }}>
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
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
