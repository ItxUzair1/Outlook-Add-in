import * as React from "react";
/* 256px source scaled in CSS — clearer in Outlook / WebView than loading the 128px file small */
import brandMarkUrl from "../../../assets/Koyomail-02-appicon-256.png";
import {
  Add24Regular,
  Edit24Regular,
  FolderOpen24Regular,
  ArrowClockwise24Regular,
  StarOff24Regular,
  EyeOff24Regular,
  SelectAllOn24Regular,
  QuestionCircle24Regular,
  AppsListDetail24Regular,
  Star24Regular,
  History24Regular,
  ChevronDown24Regular,
  Delete24Regular
} from "@fluentui/react-icons";
import {
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  MenuDivider,
} from "@fluentui/react-components";

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
  <div style={{ display: "flex", flexDirection: "column", borderRight: "1px solid #c8c6c4", padding: "2px 8px 0 8px", height: "100%", justifyContent: "center" }}>
    <div style={{ display: "flex", flexGrow: 1, gap: 2, alignItems: "flex-start" }}>
      {children}
    </div>
  </div>
);

const Toolbar = ({ 
  locations = [], 
  onFileToPath, 
  onAdd, 
  onEdit, 
  onExplore, 
  onRefresh, 
  onRemoveSuggestion, 
  onMarkUnused, 
  onDelete,
  onToggleMultiSelect, 
  onHelp,
  isMultiSelect 
}) => {
  const suggested = locations.filter(l => l.isSuggested);
  const recentlyUsed = locations
    .filter(l => l.lastUsedAt)
    .sort((a, b) => new Date(b.lastUsedAt) - new Date(a.lastUsedAt))
    .slice(0, 5);

  return (
    <div style={{ display: "flex", minHeight: 104, height: 104, backgroundColor: "#f3f2f1", borderBottom: "1px solid #edebe9", padding: "8px 0", boxSizing: "border-box", alignItems: "center" }}>
      
      <RibbonGroup label="File Email">
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <button 
              style={{ 
                display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start",
                background: "transparent", border: "1px solid transparent", cursor: "pointer", 
                padding: "2px 4px", minWidth: 64, boxSizing: "border-box"
              }}
              onMouseOver={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "#c1ddf1", border: "1px solid #7cbbed" })}
              onMouseOut={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "transparent", border: "1px solid transparent" })}
            >
              <div style={{ color: "#0078d4", marginBottom: 2, display: "flex", alignItems: "center" }}>
                <AppsListDetail24Regular />
                <ChevronDown24Regular style={{ fontSize: 12, marginLeft: 2 }} />
              </div>
              <span style={{ fontSize: 11, fontFamily: "Segoe UI", textAlign: "center", lineHeight: "1.1", color: "#323130" }}>
                File<br/>Email
              </span>
            </button>
          </MenuTrigger>

          <MenuPopover>
            <MenuList style={{ minWidth: 250 }}>
              <div style={{ padding: "4px 12px", fontSize: 12, fontWeight: "bold", color: "#605e5c" }}>Suggested Locations</div>
              {suggested.length > 0 ? (
                suggested.map(loc => (
                  <MenuItem 
                    key={loc.id} 
                    icon={<Star24Regular style={{ color: "#ffb900" }} />}
                    onClick={() => onFileToPath(loc.path)}
                  >
                    {loc.description || loc.path.split("\\").pop()}
                  </MenuItem>
                ))
              ) : (
                <div style={{ padding: "4px 32px", fontSize: 11, color: "#a19f9d", fontStyle: "italic" }}>No suggested locations</div>
              )}

              <MenuDivider />
              <div style={{ padding: "4px 12px", fontSize: 12, fontWeight: "bold", color: "#605e5c" }}>Recently Used</div>
              {recentlyUsed.length > 0 ? (
                recentlyUsed.map(loc => (
                  <MenuItem 
                    key={loc.id} 
                    icon={<History24Regular />}
                    onClick={() => onFileToPath(loc.path)}
                  >
                    {loc.description || loc.path.split("\\").pop()}
                  </MenuItem>
                ))
              ) : (
                <div style={{ padding: "4px 32px", fontSize: 11, color: "#a19f9d", fontStyle: "italic" }}>No recently used locations</div>
              )}
            </MenuList>
          </MenuPopover>
        </Menu>
      </RibbonGroup>

      <RibbonGroup label="Actions">
        <RibbonButton icon={<Add24Regular />} label="Add" onClick={onAdd} />
        <RibbonButton icon={<Edit24Regular />} label="Edit" onClick={onEdit} />
        <RibbonButton 
          icon={<Delete24Regular style={{ color: "#a4262c" }} />} 
          label="Delete" 
          onClick={onDelete} 
        />
        <RibbonButton icon={<FolderOpen24Regular />} label="Explore" onClick={onExplore} />
        <RibbonButton icon={<ArrowClockwise24Regular />} label="Refresh" onClick={onRefresh} />
        <RibbonButton icon={<Star24Regular style={{ color: "#ffb900" }}/>} label={<>Set as<br/>favourite</>} onClick={onRemoveSuggestion} />
        <RibbonButton icon={<EyeOff24Regular />} label={<>Set location<br/>unused</>} onClick={onMarkUnused} />
        <RibbonButton icon={<SelectAllOn24Regular style={isMultiSelect ? {color: "#107c10"} : {}}/>} label={<>Choose multiple<br/>locations</>} onClick={onToggleMultiSelect} />
      </RibbonGroup>

      <RibbonGroup label="Help">
        <RibbonButton 
          icon={<QuestionCircle24Regular />} 
          label={<>Koyomail<br/>help</>} 
          onClick={onHelp}
        />
      </RibbonGroup>

      {/* Brand — high-res PNG via webpack; explicit px box so IE/WebView cannot shrink the mark */}
      <div
        style={{
          marginLeft: "auto",
          display: "flex",
          alignItems: "center",
          justifyContent: "flex-end",
          flexShrink: 0,
          gap: 14,
          padding: "0 6px 0 18px",
          paddingRight: 16,
          borderLeft: "1px solid #c8c6c4",
          backgroundColor: "transparent",
        }}
      >
        <div
          style={{
            width: 88,
            height: 88,
            minWidth: 88,
            minHeight: 88,
            flexShrink: 0,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <img
            src={brandMarkUrl}
            alt=""
            style={{
              width: 88,
              height: 88,
              minWidth: 88,
              minHeight: 88,
              display: "block",
              objectFit: "contain",
              backgroundColor: "transparent",
              border: "none",
              outline: "none",
            }}
          />
        </div>
        <span
          style={{
            fontSize: 26,
            fontWeight: 600,
            color: "#0078d4",
            fontFamily: "Segoe UI, sans-serif",
            lineHeight: 1.1,
            flexShrink: 0,
          }}
        >
          Koyomail
        </span>
      </div>

    </div>
  );
};

export default Toolbar;
