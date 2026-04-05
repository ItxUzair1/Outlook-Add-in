import * as React from "react";
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogActions,
  DialogBody,
  Button,
  Input,
  Select,
} from "@fluentui/react-components";
import { 
  Checkmark20Regular, 
  Search20Regular, 
  ChevronLeft20Regular, 
  ChevronRight20Regular, 
  QuestionCircle16Regular 
} from "@fluentui/react-icons";

const Row = ({ label, children }) => (
  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
    <span style={{ width: 85, fontSize: 13, fontFamily: "Segoe UI", textAlign: "left" }}>{label}</span>
    <div style={{ flexGrow: 1, display: "flex", alignItems: "center", gap: 8 }}>
      {children}
    </div>
  </div>
);

const LocationDialog = ({ isOpen, onOpenChange, onSave, initialData }) => {
  const [data, setData] = React.useState({
    type: "Local or Network location",
    path: "",
    description: "",
    collection: "Private",
  });

  const fileInputRef = React.useRef(null);

  React.useEffect(() => {
    if (initialData) {
      setData(initialData);
    } else {
      setData({
        type: "Local or Network location",
        path: "",
        description: "",
        collection: "Private",
      });
    }
  }, [initialData, isOpen]);

  const handleBrowse = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  const handlePaste = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        setData({ ...data, path: text.trim() });
      }
    } catch (err) {
      console.error("Failed to read clipboard:", err);
    }
  };

  const onFileChange = (e) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      // Browsers hide the full path (like D:\...) for security.
      // We take the folder name the user picked.
      const folderName = files[0].webkitRelativePath.split("/")[0] || files[0].name;
      
      // We only fill the folder name so the user can easily add C:\ or D:\ in front.
      setData({ ...data, path: folderName }); 
    }
  };

  const handleSave = () => {
    onSave(data);
    onOpenChange(false);
  };

  return (
    <Dialog open={isOpen} onOpenChange={(e, d) => onOpenChange(d.open)}>
      <DialogSurface style={{ minWidth: 520 }}>
        <DialogBody>
          <DialogTitle>{initialData ? "Edit Location" : "Add Location"}</DialogTitle>
          <DialogContent style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 12 }}>
            
            <input 
              type="file" 
              ref={fileInputRef} 
              style={{ display: "none" }} 
              webkitdirectory="true" 
              onChange={onFileChange} 
            />

            <Row label="Type:">
              <Select size="small" style={{ flexGrow: 1 }} value={data.type} onChange={(e) => setData({ ...data, type: e.target.value })}>
                <option>Local or Network location</option>
                <option>OneDrive</option>
                <option>SharePoint</option>
              </Select>
            </Row>
            
            <Row label="Location:">
              <Input size="small" style={{ flexGrow: 1 }} value={data.path} onChange={(e) => setData({ ...data, path: e.target.value })} />
              <div style={{ display: "flex", gap: 4 }}>
                <Button size="small" onClick={handlePaste} style={{ width: 60, border: "1px solid #c8c6c4" }}>Paste</Button>
                <Button size="small" onClick={handleBrowse} style={{ width: 80, border: "1px solid #c8c6c4" }}>Browse...</Button>
              </div>
            </Row>

            <Row label="Description:">
              <Input size="small" style={{ flexGrow: 1 }} value={data.description} onChange={(e) => setData({ ...data, description: e.target.value })} />
              <div style={{ display: "flex", gap: 4 }}>
                <Button size="small" icon={<ChevronLeft20Regular />} style={{ minWidth: 32, padding: 0, border: "1px solid #c8c6c4" }} />
                <Button size="small" icon={<ChevronRight20Regular />} style={{ minWidth: 32, padding: 0, border: "1px solid #c8c6c4" }} />
              </div>
            </Row>

            <Row label="Portfolio:">
              <div style={{ display: "flex", alignItems: "center", flexGrow: 1, border: "1px solid #d1d1d1", borderRadius: 4, paddingLeft: 8, backgroundColor: "#fff" }}>
                <Checkmark20Regular style={{ color: "#107c10", marginRight: 4 }} />
                <Select size="small" style={{ border: "none", flexGrow: 1, boxShadow: "none" }} value={data.collection} onChange={(e) => setData({ ...data, collection: e.target.value })}>
                  <option>Private</option>
                  <option>Portfolio</option>
                  <option>Archive</option>
                </Select>
              </div>
            </Row>



            <div style={{ display: "flex", alignItems: "center", marginTop: 4 }}>
              <QuestionCircle16Regular style={{ color: "#0078d4", marginRight: 6 }} />
              <a href="#" style={{ fontSize: 13, fontFamily: "Segoe UI", color: "#0078d4", textDecoration: "none" }}>Help for sharing locations</a>
            </div>

          </DialogContent>
          <DialogActions style={{ marginTop: 24 }}>
            <Button appearance="secondary" style={{ width: 85, border: "1px solid #c8c6c4" }} onClick={handleSave}>OK</Button>
            <Button appearance="subtle" style={{ width: 85, border: "1px solid #c8c6c4" }} onClick={() => onOpenChange(false)}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default LocationDialog;
