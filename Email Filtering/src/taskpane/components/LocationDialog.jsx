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
  ChevronLeft20Regular, 
  ChevronRight20Regular, 
  QuestionCircle16Regular 
} from "@fluentui/react-icons";
import { API_BASE_URL } from "../services/backendApi.js";

const Row = ({ label, children, isNarrow, required, errorMsg }) => (
  <div style={{ 
    display: "flex", 
    flexDirection: isNarrow ? "column" : "row", 
    alignItems: isNarrow ? "stretch" : "flex-start", 
    gap: isNarrow ? 4 : 12 
  }}>
    <span style={{ 
      width: isNarrow ? "auto" : 85, 
      fontSize: 13, 
      fontFamily: "Segoe UI", 
      textAlign: "left",
      fontWeight: isNarrow ? "600" : "normal",
      marginTop: isNarrow ? 0 : 4
    }}>
      {label}
      {required && <span style={{ color: "#d13438", marginLeft: 4 }}>*</span>}
    </span>
    <div style={{ 
      flexGrow: 1, 
      display: "flex", 
      flexDirection: "column",
      gap: 2
    }}>
      <div style={{ 
        display: "flex", 
        flexDirection: isNarrow ? "column" : "row", 
        alignItems: isNarrow ? "stretch" : "center", 
        gap: 8 
      }}>
        {children}
      </div>
      {errorMsg && <span style={{ fontSize: 11, color: "#d13438", marginTop: 2 }}>{errorMsg}</span>}
    </div>
  </div>
);

function normalizePathByType(rawPath, pathType) {
  const value = String(rawPath || "").trim();
  if (!value) return "";

  if (pathType === "UNC") {
    const driveMatch = value.match(/^([a-zA-Z]):[\\/](.*)$/);
    if (driveMatch) {
      const drive = driveMatch[1].toUpperCase();
      const rest = driveMatch[2].replace(/\//g, "\\");
      return `\\\\localhost\\${drive}$\\${rest}`;
    }
    return value;
  }

  const uncToDrive = value.match(/^\\\\localhost\\([a-zA-Z])\$\\(.*)$/i);
  if (uncToDrive) {
    const drive = uncToDrive[1].toUpperCase();
    const rest = uncToDrive[2].replace(/\//g, "\\");
    return `${drive}:\\${rest}`;
  }

  return value;
}

const LocationDialog = ({ isOpen, onOpenChange, onSave, initialData }) => {
  const [data, setData] = React.useState({
    type: "Local or Network location",
    path: "",
    description: "",
    collection: "Private",
  });
  
  const [touched, setTouched] = React.useState({ path: false, description: false });

  const [width, setWidth] = React.useState(() => typeof window !== "undefined" ? window.innerWidth : 850);

  React.useEffect(() => {
    const handleResize = () => {
      setWidth(window.innerWidth);
    };
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  const isNarrow = width < 500;

  React.useEffect(() => {
    setTouched({ path: false, description: false });
    if (initialData) {
      let selectedPathType = "Drive";
      try {
        const stored = localStorage.getItem("koyomail_options");
        const parsed = stored ? JSON.parse(stored) : {};
        selectedPathType = parsed.pathType || "Drive";
      } catch {}

      // Normalise: Personal and Private are the same — always display as "Private"
      const normalisedCollection =
        initialData.collection && initialData.collection.toLowerCase() === "personal"
          ? "Private"
          : (initialData.collection || "Private");

      setData({
        ...initialData,
        collection: normalisedCollection,
        path: normalizePathByType(initialData.path, selectedPathType)
      });
    } else {
      setData({
        type: "Local or Network location",
        path: "",
        description: "",
        collection: "Private",
      });
    }
  }, [initialData, isOpen]);

  const pathInputRef = React.useRef(null);

  const handlePathChange = (val) => {
    setData((prev) => ({ ...prev, path: val }));
  };

  const handleBrowse = async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/api/search/browse-folder`);
      if (!resp.ok) {
        throw new Error("Unable to open folder picker");
      }
      const result = await resp.json();
      if (result?.path) {
        handlePathChange(String(result.path).trim());
        setTouched(prev => ({ ...prev, path: true }));
        // Force WebView2 to completely repaint after native dialog closes
        setTimeout(() => {
          if (pathInputRef.current) {
            pathInputRef.current.blur();
            pathInputRef.current.focus();
          }
          window.dispatchEvent(new Event('resize'));
        }, 150);
      }
    } catch (err) {
      console.error("Browse failed:", err);
    }
  };

  const handlePaste = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        handlePathChange(text.trim());
        setTouched(prev => ({ ...prev, path: true }));
      }
    } catch (err) {
      console.error("Failed to read clipboard:", err);
    }
  };

  const handleSave = () => {
    if (!data.path.trim() || !data.description.trim()) {
      setTouched({ path: true, description: true });
      return;
    }

    let selectedPathType = "Drive";
    try {
      const stored = localStorage.getItem("koyomail_options");
      const parsed = stored ? JSON.parse(stored) : {};
      selectedPathType = parsed.pathType || "Drive";
    } catch {}

    // Normalise: Personal and Private are the same — always save as "Private"
    const rawCollection = data.collection;
    const normalisedCollection =
      !rawCollection || rawCollection.toLowerCase() === "personal"
        ? "Private"
        : rawCollection;

    onSave({
      ...data,
      collection: normalisedCollection,
      path: normalizePathByType(data.path, selectedPathType),
    });
    onOpenChange(false);
  };

  const isPathError = touched.path && !data.path.trim();
  const isDescError = touched.description && !data.description.trim();
  const isSaveDisabled = !data.path.trim() || !data.description.trim();

  return (
    <Dialog open={isOpen} onOpenChange={(e, d) => onOpenChange(d.open)}>
      <DialogSurface style={{ width: "95%", maxWidth: 520, boxSizing: "border-box" }}>
        <DialogBody>
          <DialogTitle>{initialData ? "Edit Location" : "Add Location"}</DialogTitle>
          <DialogContent style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 12 }}>
            
            <Row label="Type:" isNarrow={isNarrow}>
              <Select size="small" style={{ flexGrow: 1 }} value={data.type} onChange={(e) => setData({ ...data, type: e.target.value })}>
                <option>Local or Network location</option>
                <option>OneDrive</option>
                <option>SharePoint</option>
              </Select>
            </Row>
            
            <Row 
              label="Location:" 
              isNarrow={isNarrow} 
              required 
              errorMsg={isPathError ? "Location path is required." : null}
            >
              <Input 
                ref={pathInputRef} 
                size="small" 
                style={{ flexGrow: 1, border: isPathError ? "1px solid #d13438" : undefined }} 
                value={data.path} 
                onChange={(e) => handlePathChange(e.target.value)} 
                onBlur={() => setTouched(prev => ({ ...prev, path: true }))}
              />
              <div style={{ display: "flex", gap: 4, justifyContent: isNarrow ? "flex-end" : "flex-start" }}>
                <Button size="small" onClick={handlePaste} style={{ width: 60, border: "1px solid #c8c6c4" }}>Paste</Button>
                <Button size="small" onClick={handleBrowse} style={{ width: 80, border: "1px solid #c8c6c4" }}>Browse...</Button>
              </div>
            </Row>

            <Row 
              label="Description:" 
              isNarrow={isNarrow} 
              required 
              errorMsg={isDescError ? "Description is required." : null}
            >
              <Input 
                size="small" 
                style={{ flexGrow: 1, border: isDescError ? "1px solid #d13438" : undefined }} 
                value={data.description} 
                onChange={(e) => setData({ ...data, description: e.target.value })} 
                onBlur={() => setTouched(prev => ({ ...prev, description: true }))}
              />
              <div style={{ display: "flex", gap: 4, justifyContent: isNarrow ? "flex-end" : "flex-start" }}>
                <Button size="small" icon={<ChevronLeft20Regular />} style={{ minWidth: 32, padding: 0, border: "1px solid #c8c6c4" }} />
                <Button size="small" icon={<ChevronRight20Regular />} style={{ minWidth: 32, padding: 0, border: "1px solid #c8c6c4" }} />
              </div>
            </Row>

            <Row label="Portfolio:" isNarrow={isNarrow}>
              <div style={{ display: "flex", alignItems: "center", flexGrow: 1, border: "1px solid #d1d1d1", borderRadius: 4, paddingLeft: 8, backgroundColor: "#fff" }}>
                <Checkmark20Regular style={{ color: "#107c10", marginRight: 4 }} />
                <Select size="small" style={{ border: "none", flexGrow: 1, boxShadow: "none" }} value={data.collection} onChange={(e) => setData({ ...data, collection: e.target.value })}>
                  <option>Private</option>
                  <option>Portfolio</option>
                  <option>Archive</option>
                  {data.collection && !["Private", "Portfolio", "Archive"].includes(data.collection) && (
                    <option value={data.collection}>{data.collection}</option>
                  )}
                </Select>
              </div>
            </Row>

            <div style={{ display: "flex", alignItems: "center", marginTop: 4 }}>
              <QuestionCircle16Regular style={{ color: "#0078d4", marginRight: 6 }} />
              <a href="#" style={{ fontSize: 13, fontFamily: "Segoe UI", color: "#0078d4", textDecoration: "none" }}>Help for sharing locations</a>
            </div>

          </DialogContent>
          <DialogActions style={{ marginTop: 24 }}>
            <Button appearance="primary" disabled={isSaveDisabled} style={{ width: 85 }} onClick={handleSave}>OK</Button>
            <Button appearance="subtle" style={{ width: 85, border: "1px solid #c8c6c4" }} onClick={() => onOpenChange(false)}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default LocationDialog;
