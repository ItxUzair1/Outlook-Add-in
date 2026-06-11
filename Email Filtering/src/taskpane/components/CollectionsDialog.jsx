import * as React from "react";
import PropTypes from "prop-types";
import { 
  DocumentAdd24Regular, 
  FolderAdd24Regular, 
  Save24Regular, 
  Dismiss24Regular, 
  Add24Regular, 
  Edit24Regular, 
  Delete24Regular,
  FolderOpen24Regular,
  FolderProhibited24Regular,
  SelectAllOn24Regular,
  ClipboardPaste24Regular,
  Cut24Regular,
  Copy24Regular,
  Search24Regular,
  TextDescription24Regular,
  Checkmark16Regular,
  ArrowClockwise24Regular
} from "@fluentui/react-icons";
import { Input, Table, TableHeader, TableRow, TableHeaderCell, TableBody, TableCell, Dialog, DialogSurface, DialogTitle, DialogBody, DialogContent, DialogActions, Button, Select, Label, Checkbox } from "@fluentui/react-components";
import brandMarkUrl from "../../../assets/koyomail_icon_v2.png";
import { API_BASE_URL } from "../../taskpane/services/backendApi.js";

const RibbonButton = ({ icon, label, onClick, disabled }) => (
  <button 
    onClick={onClick} 
    disabled={disabled}
    style={{ 
      display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start",
      background: "transparent", border: "1px solid transparent", cursor: disabled ? "not-allowed" : "pointer", 
      padding: "2px 4px", minWidth: 48, boxSizing: "border-box",
      opacity: disabled ? 0.45 : 1
    }}
    onMouseOver={(e) => !disabled && Object.assign(e.currentTarget.style, { backgroundColor: "#c1ddf1", border: "1px solid #7cbbed" })}
    onMouseOut={(e) => !disabled && Object.assign(e.currentTarget.style, { backgroundColor: "transparent", border: "1px solid transparent" })}
  >
    <div style={{ color: "#0078d4", marginBottom: 2 }}>{icon}</div>
    <span style={{ fontSize: 11, fontFamily: "'Exo 2', 'Segoe UI', sans-serif", textAlign: "center", lineHeight: "1.1", color: "#323130" }}>
      {label}
    </span>
  </button>
);

const RibbonGroup = ({ label, children }) => (
  <div style={{ display: "flex", flexDirection: "column", borderRight: "1px solid #c8c6c4", padding: "2px 8px 0 8px", height: "100%", justifyContent: "center" }}>
    <div style={{ display: "flex", flexGrow: 1, gap: 2, alignItems: "flex-start" }}>
      {children}
    </div>
    <div style={{ fontSize: 10, color: "#605e5c", textAlign: "center", marginTop: "auto", paddingBottom: 2 }}>
      {label}
    </div>
  </div>
);

const CollectionsDialog = ({ isOpen, onOpenChange }) => {
  const [collections, setCollections] = React.useState([]);
  const [selectedCollectionId, setSelectedCollectionId] = React.useState(null);
  const [collectionsFilter, setCollectionsFilter] = React.useState("");
  const [locationsFilter, setLocationsFilter] = React.useState("");
  const [selectedLocationId, setSelectedLocationId] = React.useState(null);
  const [editingLocationId, setEditingLocationId] = React.useState(null);
  const [collectionToDelete, setCollectionToDelete] = React.useState(null);
  const [locationToDelete, setLocationToDelete] = React.useState(null);

  const loadPersistedCollections = async () => {
    try {
      const loadedStr = localStorage.getItem("koyomail_loaded_collections");
      if (loadedStr) {
        const filePaths = JSON.parse(loadedStr);
        const loadedCollections = [];
        for (const filePath of filePaths) {
          const loadResp = await fetch(`${API_BASE_URL}/api/collections/load`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ filePath })
          });
          if (loadResp.ok) {
            const data = await loadResp.json();
            const filename = filePath.split('\\').pop().split('/').pop();
            loadedCollections.push({
              id: filePath,
              name: filename.replace(/\.mmcollection$/i, ''),
              locations: data.locations || []
            });
          }
        }
        setCollections(loadedCollections);
      }
    } catch (err) {
      console.error("Failed to load persisted collections:", err);
    }
  };

  React.useEffect(() => {
    if (isOpen) {
      loadPersistedCollections();
    }
  }, [isOpen]);

  const saveToLocalStorage = (cols) => {
    const filePaths = cols.map(c => c.id);
    localStorage.setItem("koyomail_loaded_collections", JSON.stringify(filePaths));
  };

  const [isNewDialogOpen, setIsNewDialogOpen] = React.useState(false);
  const [newCollectionType, setNewCollectionType] = React.useState("Local and network folder");
  const [newCollectionPath, setNewCollectionPath] = React.useState("");
  const [newCollectionFilename, setNewCollectionFilename] = React.useState("");

  const [isAddLocationDialogOpen, setIsAddLocationDialogOpen] = React.useState(false);
  const [addLocationType, setAddLocationType] = React.useState("Local and network folder");
  const [addLocationPath, setAddLocationPath] = React.useState("");
  const [addLocationDesc, setAddLocationDesc] = React.useState("");

  const handleAddCollectionClick = async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/api/search/browse-file`);
      if (!resp.ok) throw new Error("Unable to open file picker");
      const result = await resp.json();
      if (result?.path) {
        const filePath = String(result.path).trim();
        
        // Prevent duplicates
        if (collections.some(c => c.id === filePath)) {
          setSelectedCollectionId(filePath);
          return;
        }

        // Load collection data from file
        const loadResp = await fetch(`${API_BASE_URL}/api/collections/load`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ filePath })
        });

        if (loadResp.ok) {
          const data = await loadResp.json();
          const filename = filePath.split('\\').pop().split('/').pop();
          const newCol = {
            id: filePath,
            name: filename.replace(/\.mmcollection$/i, ''),
            locations: data.locations || []
          };
          const updatedCols = [...collections, newCol];
          setCollections(updatedCols);
          saveToLocalStorage(updatedCols);
          setSelectedCollectionId(newCol.id);
        } else {
          alert("Failed to load the selected collection file.");
        }
      }
    } catch (err) {
      console.error("Browse file failed:", err);
    }
  };

  const handleDeleteCollectionClick = () => {
    if (!selectedCollectionId) return;
    setCollectionToDelete(collections.find(c => c.id === selectedCollectionId));
  };

  const confirmDeleteCollection = () => {
    if (!collectionToDelete) return;
    const updatedCols = collections.filter(c => c.id !== collectionToDelete.id);
    setCollections(updatedCols);
    saveToLocalStorage(updatedCols);
    if (selectedCollectionId === collectionToDelete.id) {
      setSelectedCollectionId(null);
      setSelectedLocationId(null);
    }
    setCollectionToDelete(null);
  };

  const handleDeleteLocationClick = () => {
    if (!selectedCollectionId || !selectedLocationId) return;
    const selectedCollection = collections.find(c => c.id === selectedCollectionId);
    const locToDelete = selectedCollection?.locations.find(l => l.id === selectedLocationId);
    if (locToDelete) setLocationToDelete({ collection: selectedCollection, location: locToDelete });
  };

  const confirmDeleteLocation = async () => {
    if (!locationToDelete) return;
    const { collection, location } = locationToDelete;

    const updatedCollections = collections.map(c => {
      if (c.id === collection.id) {
        return { ...c, locations: c.locations.filter(l => l.id !== location.id) };
      }
      return c;
    });

    try {
      const colToSave = updatedCollections.find(c => c.id === collection.id);
      await fetch(`${API_BASE_URL}/api/collections/save`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filePath: colToSave.id, locations: colToSave.locations })
      });

      setCollections(updatedCollections);
      if (selectedLocationId === location.id) setSelectedLocationId(null);
      setLocationToDelete(null);
    } catch (err) {
      console.error("Failed to delete location:", err);
      setLocationToDelete(null);
    }
  };

  const handleEditPathClick = () => {
    const loc = selectedCollection?.locations.find(l => l.id === selectedLocationId);
    if (loc) {
      setEditingLocationId(loc.id);
      setAddLocationType(loc.type === "msg" ? "Local and network folder" : loc.type);
      setAddLocationPath(loc.folder || loc.path || "");
      setAddLocationDesc(loc.description || "");
      setIsAddLocationDialogOpen(true);
    }
  };

  const handleAddLocationClick = () => {
    setEditingLocationId(null);
    setAddLocationType("Local and network folder");
    setAddLocationPath("");
    setAddLocationDesc("");
    setIsAddLocationDialogOpen(true);
  };

  const [isSaveDialogOpen, setIsSaveDialogOpen] = React.useState(false);

  const selectedCollection = collections.find(c => c.id === selectedCollectionId);

  const handleBrowse = async (setter) => {
    try {
      let url = `${API_BASE_URL}/api/search/browse-folder`;
      if (selectedCollection?.id) {
        const pathParts = selectedCollection.id.split('\\');
        pathParts.pop();
        const startPath = pathParts.join('\\');
        url += `?startPath=${encodeURIComponent(startPath)}`;
      }
      
      const resp = await fetch(url);
      if (!resp.ok) throw new Error("Unable to open folder picker");
      const result = await resp.json();
      if (result?.path) {
        setter(String(result.path).trim());
      }
    } catch (err) {
      console.error("Browse failed:", err);
    }
  };

  const handleCloseClick = () => {
    // Show save dialog if there are collections, else just close
    if (collections.length > 0) {
      setIsSaveDialogOpen(true);
    } else {
      onOpenChange(false);
    }
  };

  const handleNewOk = async () => {
    if (!newCollectionPath || !newCollectionFilename) return;
    
    // Construct full path, ensuring it ends with .mmcollection
    let filename = newCollectionFilename;
    if (!filename.toLowerCase().endsWith('.mmcollection')) {
      filename += '.mmcollection';
    }
    const sep = newCollectionPath.endsWith('\\') || newCollectionPath.endsWith('/') ? '' : '\\';
    const fullPath = `${newCollectionPath}${sep}${filename}`;

    try {
      // Create empty collection file using backend
      await fetch(`${API_BASE_URL}/api/collections/save`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filePath: fullPath, locations: [] })
      });

      // Add to UI
      const newCol = {
        id: fullPath, // use path as ID for now
        name: filename.replace(/\.mmcollection$/i, ''),
        locations: []
      };
      const updatedCols = [...collections, newCol];
      setCollections(updatedCols);
      saveToLocalStorage(updatedCols);
      setSelectedCollectionId(newCol.id);
      setIsNewDialogOpen(false);
    } catch (err) {
      console.error("Failed to create new collection:", err);
      alert("Failed to create new collection. See console.");
    }
  };

  const handleAddLocationOk = async () => {
    if (!selectedCollection || !addLocationPath) return;

    const newLocation = {
      id: crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(),
      type: addLocationType === "Local and network folder" ? "msg" : addLocationType,
      folder: addLocationPath,
      description: addLocationDesc || addLocationPath.split('\\').pop() || addLocationPath
    };

    const updatedCollections = collections.map(c => {
      if (c.id === selectedCollection.id) {
        if (editingLocationId) {
          return { ...c, locations: c.locations.map(l => l.id === editingLocationId ? newLocation : l) };
        } else {
          return { ...c, locations: [...c.locations, newLocation] };
        }
      }
      return c;
    });

    try {
      // Save to backend immediately
      const colToSave = updatedCollections.find(c => c.id === selectedCollection.id);
      await fetch(`${API_BASE_URL}/api/collections/save`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filePath: colToSave.id, locations: colToSave.locations })
      });

      setCollections(updatedCollections);
      setIsAddLocationDialogOpen(false);
      setAddLocationPath("");
      setAddLocationDesc("");
    } catch (err) {
      console.error("Failed to add location:", err);
      alert("Failed to add location. See console.");
    }
  };

  // Not rendering if not open
  if (!isOpen) return null;

  return (
    <div style={{ height: "100vh", display: "flex", flexDirection: "column", fontFamily: "'Exo 2', 'Segoe UI', sans-serif" }}>
      
      {/* Ribbon Toolbar */}
      <div style={{ display: "flex", minHeight: 80, height: 80, backgroundColor: "#f3f2f1", borderBottom: "1px solid #edebe9", padding: "0", boxSizing: "border-box", alignItems: "center" }}>
        
        <RibbonGroup label="Collections">
          <RibbonButton icon={<DocumentAdd24Regular />} label="New" onClick={() => setIsNewDialogOpen(true)} />
          <RibbonButton icon={<FolderAdd24Regular />} label="Add" onClick={handleAddCollectionClick} />
          <RibbonButton icon={<FolderProhibited24Regular style={{color: "#a4262c"}}/>} label={<>Delete<br/>Collection</>} disabled={!selectedCollectionId} onClick={handleDeleteCollectionClick} />
          <RibbonButton icon={<ArrowClockwise24Regular style={{color: "#0078d4"}}/>} label="Refresh" onClick={loadPersistedCollections} />
          <RibbonButton icon={<Dismiss24Regular />} label="Close" onClick={handleCloseClick} />
        </RibbonGroup>

        <RibbonGroup label="Locations">
          <RibbonButton icon={<Add24Regular style={{color: "#107c10"}}/>} label={<>Add<br/>Location</>} disabled={!selectedCollection} onClick={handleAddLocationClick} />
          <RibbonButton icon={<Edit24Regular style={{color: "#d83b01"}}/>} label="Edit" disabled={!selectedCollection || !selectedLocationId} onClick={handleEditPathClick} />
          <RibbonButton icon={<Delete24Regular style={{color: "#a4262c"}}/>} label={<>Delete<br/>Location</>} disabled={!selectedCollection || !selectedLocationId} onClick={handleDeleteLocationClick} />
        </RibbonGroup>

        {/* Brand */}
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", justifyContent: "flex-end", flexShrink: 0, gap: 8, padding: "0 16px", backgroundColor: "transparent" }}>
          <img src={brandMarkUrl} alt="" style={{ width: 68, height: 68, objectFit: "contain" }} />
          <span style={{ fontSize: 22, fontWeight: 700, color: "#000", lineHeight: 1.1, letterSpacing: "2px", textTransform: "uppercase" }}>
            Koyomail
          </span>
        </div>
      </div>

      {/* Main Content Area */}
      <div style={{ display: "flex", flexGrow: 1, overflow: "hidden", backgroundColor: "#fff" }}>
        
        {/* Left Pane: Collections */}
        <div style={{ width: 280, borderRight: "1px solid #edebe9", display: "flex", flexDirection: "column" }}>
          <div style={{ padding: "8px", borderBottom: "1px solid #edebe9", display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 12, color: "#323130", whiteSpace: "nowrap" }}>Collections filter:</span>
            <Input 
              style={{ flexGrow: 1, minWidth: 0 }} 
              value={collectionsFilter} 
              onChange={(e, data) => setCollectionsFilter(data.value)} 
              appearance="outline"
              contentAfter={<Dismiss24Regular style={{ fontSize: 14, cursor: "pointer", color: "#605e5c", visibility: collectionsFilter ? "visible" : "hidden" }} onClick={() => setCollectionsFilter("")} />}
            />
          </div>
          <div style={{ flexGrow: 1, overflowY: "auto" }}>
            <Table size="small">
              <TableHeader>
                <TableRow>
                  <TableHeaderCell style={{ width: 50 }}>Status</TableHeaderCell>
                  <TableHeaderCell>Collection</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {collections.filter(c => !collectionsFilter || c.name.toLowerCase().includes(collectionsFilter.toLowerCase())).map((c) => (
                  <TableRow 
                    key={c.id} 
                    style={{ 
                      backgroundColor: selectedCollectionId === c.id ? "#c7e0f4" : "transparent",
                      cursor: "pointer" 
                    }}
                    onClick={() => {
                      setSelectedCollectionId(c.id);
                      setSelectedLocationId(null);
                    }}
                  >
                    <TableCell><Checkmark16Regular style={{ color: "#107c10" }} /></TableCell>
                    <TableCell>{c.name}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
          <div style={{ padding: "4px 8px", borderTop: "1px solid #edebe9", fontSize: 11, color: "#605e5c", backgroundColor: "#f3f2f1" }}>
            {collections.filter(c => !collectionsFilter || c.name.toLowerCase().includes(collectionsFilter.toLowerCase())).length} collections
          </div>
        </div>

        {/* Right Pane: Locations */}
        <div style={{ flexGrow: 1, display: "flex", flexDirection: "column" }}>
          <div style={{ padding: "8px", borderBottom: "1px solid #edebe9", display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 12, color: "#323130", whiteSpace: "nowrap" }}>Locations filter:</span>
            <Input 
              style={{ flexGrow: 1, minWidth: 0 }} 
              value={locationsFilter} 
              onChange={(e, data) => setLocationsFilter(data.value)} 
              appearance="outline"
              contentAfter={<Dismiss24Regular style={{ fontSize: 14, cursor: "pointer", color: "#605e5c", visibility: locationsFilter ? "visible" : "hidden" }} onClick={() => setLocationsFilter("")} />}
            />
          </div>
          <div style={{ flexGrow: 1, overflowY: "auto" }}>
            <Table size="small">
              <TableHeader>
                <TableRow>
                  <TableHeaderCell style={{ width: 50 }}>Status</TableHeaderCell>
                  <TableHeaderCell>Description</TableHeaderCell>
                  <TableHeaderCell>Path</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {selectedCollection?.locations.filter(loc => 
                  !locationsFilter || 
                  (loc.description || "").toLowerCase().includes(locationsFilter.toLowerCase()) || 
                  (loc.folder || loc.path || "").toLowerCase().includes(locationsFilter.toLowerCase())
                ).length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={3}>
                      <div style={{ padding: 20, textAlign: "center", color: "#605e5c", fontStyle: "italic", fontSize: 12 }}>
                        No locations in this collection.
                      </div>
                    </TableCell>
                  </TableRow>
                ) : (
                  selectedCollection?.locations.filter(loc => 
                    !locationsFilter || 
                    (loc.description || "").toLowerCase().includes(locationsFilter.toLowerCase()) || 
                    (loc.folder || loc.path || "").toLowerCase().includes(locationsFilter.toLowerCase())
                  ).map((loc, idx) => (
                    <TableRow 
                      key={loc.id || idx}
                      style={{ 
                        backgroundColor: selectedLocationId === (loc.id || idx) ? "#e1dfdd" : "transparent",
                        cursor: "pointer" 
                      }}
                      onClick={() => setSelectedLocationId(loc.id || idx)}
                    >
                      <TableCell>
                        {(loc.folder || loc.path) && <Checkmark16Regular style={{ color: "#107c10" }} />}
                      </TableCell>
                      <TableCell>{loc.description}</TableCell>
                      <TableCell>{loc.folder || loc.path}</TableCell>
                    </TableRow>
                  ))
                )}
              </TableBody>
            </Table>
          </div>
        </div>

      </div>

      {/* New Collection Dialog */}
      <Dialog open={isNewDialogOpen} onOpenChange={(e, data) => setIsNewDialogOpen(data.open)}>
        <DialogSurface style={{ maxWidth: 450 }}>
          <DialogBody>
            <DialogTitle>New Collection</DialogTitle>
            <DialogContent style={{ display: "flex", flexDirection: "column", gap: 16, paddingTop: 12 }}>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr", alignItems: "center", gap: 8 }}>
                <Label size="small">Folder type:</Label>
                <Select value={newCollectionType} onChange={(e) => setNewCollectionType(e.target.value)}>
                  <option>Local and network folder</option>
                </Select>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr auto", alignItems: "center", gap: 8 }}>
                <Label size="small">Path:</Label>
                <Input value={newCollectionPath} onChange={(e, data) => setNewCollectionPath(data.value)} />
                <Button onClick={() => handleBrowse(setNewCollectionPath)}>Browse...</Button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr", alignItems: "center", gap: 8 }}>
                <Label size="small">Collection name:</Label>
                <Input value={newCollectionFilename} onChange={(e, data) => setNewCollectionFilename(data.value)} />
              </div>
            </DialogContent>
            <DialogActions style={{ marginTop: 24 }}>
              <Button appearance="primary" onClick={handleNewOk}>OK</Button>
              <Button onClick={() => setIsNewDialogOpen(false)}>Cancel</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Add Location Dialog */}
      <Dialog open={isAddLocationDialogOpen} onOpenChange={(e, data) => setIsAddLocationDialogOpen(data.open)}>
        <DialogSurface style={{ maxWidth: 450 }}>
          <DialogBody>
            <DialogTitle>{editingLocationId ? "Edit" : "Add Location"}</DialogTitle>
            <DialogContent style={{ display: "flex", flexDirection: "column", gap: 16, paddingTop: 12 }}>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr", alignItems: "center", gap: 8 }}>
                <Label size="small">Type:</Label>
                <Select value={addLocationType} onChange={(e) => setAddLocationType(e.target.value)}>
                  <option>Local and network folder</option>
                </Select>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr auto", alignItems: "center", gap: 8 }}>
                <Label size="small">Location:</Label>
                <Input value={addLocationPath} onChange={(e, data) => setAddLocationPath(data.value)} />
                <Button onClick={() => handleBrowse(setAddLocationPath)}>Browse...</Button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr auto", alignItems: "center", gap: 8 }}>
                <Label size="small">Description:</Label>
                <Input value={addLocationDesc} onChange={(e, data) => setAddLocationDesc(data.value)} />
                <div style={{ display: "flex", gap: 4 }}>
                  <Button icon={<span style={{fontSize: 10}}>&lt;</span>} style={{ minWidth: 32, padding: 0 }} />
                  <Button icon={<span style={{fontSize: 10}}>&gt;</span>} style={{ minWidth: 32, padding: 0 }} />
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "100px 1fr", alignItems: "center", gap: 8 }}>
                <Label size="small">Collection:</Label>
                <Select value={selectedCollection?.id || ""} disabled>
                  {selectedCollection && <option value={selectedCollection.id}>✓ {selectedCollection.name}</option>}
                </Select>
              </div>
            </DialogContent>
            <DialogActions style={{ marginTop: 24, display: "flex", justifyContent: "flex-end" }}>
              <div style={{ display: "flex", gap: 8 }}>
                <Button appearance="primary" onClick={handleAddLocationOk}>OK</Button>
                <Button onClick={() => setIsAddLocationDialogOpen(false)}>Cancel</Button>
              </div>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Save All Dialog */}
      <Dialog open={isSaveDialogOpen} onOpenChange={(e, data) => setIsSaveDialogOpen(data.open)}>
        <DialogSurface style={{ maxWidth: 450 }}>
          <DialogBody>
            <DialogTitle>Save All</DialogTitle>
            <DialogContent style={{ display: "flex", flexDirection: "column", gap: 8, paddingTop: 12 }}>
              <div style={{ fontSize: 13, color: "#323130", fontWeight: "600", marginBottom: 8, backgroundColor: "#ffff00", display: "inline-block", padding: "2px 4px" }}>
                Save changes to the following items?
              </div>
              <div style={{ border: "1px solid #c8c6c4", height: 120, overflowY: "auto", backgroundColor: "#fff" }}>
                {collections.map(c => (
                  <div key={c.id} style={{ display: "flex", alignItems: "center", padding: "4px 8px", backgroundColor: "#e1dfdd" }}>
                    <Checkbox checked={true} />
                    <span style={{ fontSize: 13, marginLeft: 8 }}>{c.name}.mmcollection</span>
                  </div>
                ))}
              </div>
            </DialogContent>
            <DialogActions style={{ marginTop: 16 }}>
              <Button appearance="primary" onClick={() => { setIsSaveDialogOpen(false); onOpenChange(false); }}>Yes</Button>
              <Button onClick={() => { setIsSaveDialogOpen(false); onOpenChange(false); }}>No</Button>
              <Button onClick={() => setIsSaveDialogOpen(false)}>Cancel</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Delete Collection Dialog */}
      <Dialog open={!!collectionToDelete} onOpenChange={(e, data) => !data.open && setCollectionToDelete(null)}>
        <DialogSurface style={{ maxWidth: 400 }}>
          <DialogBody>
            <DialogTitle>Delete Collection</DialogTitle>
            <DialogContent>
              Are you sure you want to remove the collection "{collectionToDelete?.name}" from your list?
            </DialogContent>
            <DialogActions style={{ marginTop: 24, display: "flex", justifyContent: "flex-end" }}>
              <div style={{ display: "flex", gap: 8 }}>
                <Button appearance="primary" style={{ backgroundColor: "#a4262c", color: "white" }} onClick={confirmDeleteCollection}>Delete</Button>
                <Button onClick={() => setCollectionToDelete(null)}>Cancel</Button>
              </div>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Delete Location Dialog */}
      <Dialog open={!!locationToDelete} onOpenChange={(e, data) => !data.open && setLocationToDelete(null)}>
        <DialogSurface style={{ maxWidth: 400 }}>
          <DialogBody>
            <DialogTitle>Delete Location</DialogTitle>
            <DialogContent>
              Are you sure you want to delete the location "{locationToDelete?.location?.description || locationToDelete?.location?.folder || locationToDelete?.location?.path}"?
            </DialogContent>
            <DialogActions style={{ marginTop: 24, display: "flex", justifyContent: "flex-end" }}>
              <div style={{ display: "flex", gap: 8 }}>
                <Button appearance="primary" style={{ backgroundColor: "#a4262c", color: "white" }} onClick={confirmDeleteLocation}>Delete</Button>
                <Button onClick={() => setLocationToDelete(null)}>Cancel</Button>
              </div>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

    </div>
  );
};

CollectionsDialog.propTypes = {
  isOpen: PropTypes.bool.isRequired,
  onOpenChange: PropTypes.func.isRequired,
};

export default CollectionsDialog;
