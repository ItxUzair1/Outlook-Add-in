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
  toggleSuggestion
} from "../services/backendApi";
import { buildCurrentEmailPayload } from "../services/mailboxService";
import Toolbar from "./Toolbar";
import DetailsSidebar from "./DetailsSidebar";
import LocationTable from "./LocationTable";
import LocationDialog from "./LocationDialog";
import { Button, Spinner } from "@fluentui/react-components";

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

  const [loading, setLoading] = React.useState(false);
  const [message, setMessage] = React.useState("");
  
  // Dialog State
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
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
    // Initial fetch of email metadata
    buildCurrentEmailPayload().then(payload => {
      setSubject(payload.subject || "");
    }).catch(() => {});
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
    if (selectedIds.length === 0) return;
    if (!window.confirm("Are you sure you want to delete the selected location(s)?")) return;

    try {
      for (const id of selectedIds) {
        await deleteLocation(id);
      }
      setSelectedIds([]);
      setMessage("Location(s) deleted.");
      await loadLocations();
    } catch (error) {
      setMessage(`Delete failed: ${error.message}`);
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

      const basePayload = await buildCurrentEmailPayload();
      const response = await fileEmail({
        ...basePayload,
        subject, // Use the subject from state (may be edited)
        comment,
        afterFiling,
        markReviewed,
        sendLink,
        attachmentsOption,
        targetPaths: selectedLocations.map((x) => x.path),
      });

      setMessage("Email filed successfully.");
      
      // Close the dialog if we are running in one
      if (Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent("close");
      }
      
      // Perform after-filing actions in Outlook
      if (afterFiling === "delete") {
        Office.context.mailbox.item.removeAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            setMessage("Email filed, but failed to move to Deleted Items.");
          }
        });
      }

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

  return (
    <div style={{ height: "100vh", display: "flex", flexDirection: "column", fontFamily: "Segoe UI" }}>
      <Toolbar 
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
        onMarkUnused={() => setMessage("Mark as unused not implemented")}
        onToggleMultiSelect={() => setIsMultiSelect(!isMultiSelect)}
        isMultiSelect={isMultiSelect}
        onDelete={onDeleteLocation}
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

      <div style={{ padding: 12, borderTop: "1px solid #edebe9", display: "flex", justifyContent: "flex-end", gap: 8, backgroundColor: "#f3f2f1" }}>
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

      <LocationDialog 
        isOpen={isDialogOpen}
        onOpenChange={setIsDialogOpen}
        onSave={onSaveLocation}
        initialData={editingLocation}
      />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
