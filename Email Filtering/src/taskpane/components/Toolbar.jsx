import * as React from "react";
import {
  Add24Regular,
  Edit24Regular,
  FolderOpen24Regular,
  ArrowClockwise24Regular,
  StarOff24Regular,
  EyeOff24Regular,
  SelectAllOn24Regular,
  QuestionCircle24Regular,
  AppsListDetail24Regular
} from "@fluentui/react-icons";

const RibbonButton = ({ icon, label, onClick }) => (
  <button 
    onClick={onClick} 
    style={{ 
      display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start",
      background: "transparent", border: "1px solid transparent", cursor: "pointer", 
      padding: "2px 4px", minWidth: 48, boxSizing: "border-box"
    }}
    onMouseOver={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "#c1ddf1", border: "1px solid #7cbbed" })}
    onMouseOut={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "transparent", border: "1px solid transparent" })}
  >
    <div style={{ color: "#0078d4", marginBottom: 2 }}>{icon}</div>
    <span style={{ fontSize: 11, fontFamily: "Segoe UI", textAlign: "center", lineHeight: "1.1", color: "#323130" }}>
      {label}
    </span>
  </button>
);

const RibbonGroup = ({ label, children }) => (
  <div style={{ display: "flex", flexDirection: "column", borderRight: "1px solid #c8c6c4", padding: "2px 8px 0 8px", height: "100%" }}>
    <div style={{ display: "flex", flexGrow: 1, gap: 2, alignItems: "flex-start" }}>
      {children}
    </div>
    <div style={{ fontSize: 11, fontFamily: "Segoe UI", color: "#605e5c", textAlign: "center", marginTop: "auto", paddingBottom: 2 }}>
      {label}
    </div>
  </div>
);

const Toolbar = ({ onAdd, onEdit, onManage, onExplore, onRefresh, onRemoveSuggestion, onMarkUnused, onToggleMultiSelect, isMultiSelect }) => {
  return (
    <div style={{ display: "flex", height: 80, backgroundColor: "#f3f2f1", borderBottom: "1px solid #edebe9", padding: "4px 0" }}>
      
      <RibbonGroup label="Locations">
        <RibbonButton icon={<Add24Regular />} label="Add" onClick={onAdd} />
        <RibbonButton icon={<Edit24Regular />} label="Edit" onClick={onEdit} />
        <RibbonButton icon={<FolderOpen24Regular />} label="Explore" onClick={onExplore} />
        <RibbonButton icon={<ArrowClockwise24Regular />} label="Refresh" onClick={onRefresh} />
      </RibbonGroup>

      <RibbonGroup label="Usage">
        <RibbonButton icon={<StarOff24Regular style={{ color: "#a4262c" }}/>} label={<>Remove<br/>suggestion</>} onClick={onRemoveSuggestion} />
        <RibbonButton icon={<EyeOff24Regular />} label={<>Mark as<br/>Unused</>} onClick={onMarkUnused} />
      </RibbonGroup>

      <RibbonGroup label="Selection">
        <RibbonButton icon={<SelectAllOn24Regular style={isMultiSelect ? {color: "#107c10"} : {}}/>} label={<>Select<br/>Multiple</>} onClick={onToggleMultiSelect} />
      </RibbonGroup>

      <RibbonGroup label="Help">
        <RibbonButton icon={<QuestionCircle24Regular />} label={<>Filing<br/>Help</>} />
      </RibbonGroup>

    </div>
  );
};

export default Toolbar;
