import * as React from "react";
/* 512px source scaled in CSS — much clearer on high-DPI displays */
import brandMarkUrl from "../../../assets/koyomail_icon_v2.png";
import {
  Add24Regular,
  Edit24Regular,
  FolderOpen24Regular,
  ArrowClockwise24Regular,
  StarOff24Regular,
  EyeOff24Regular,
  Eye24Regular,
  SelectAllOn24Regular,
  QuestionCircle24Regular,
  AppsListDetail24Regular,
  Star24Regular,
  History24Regular,
  ChevronDown24Regular,
  Delete24Regular,
  MoreHorizontal24Regular
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
  isMultiSelect,
  isAuthOk = false,
  hasUnusedSelected = false,
  hasCollectionSelected = false
}) => {
  const suggested = locations.filter(l => l.isSuggested);
  const recentlyUsed = locations
    .filter(l => l.lastUsedAt && !l.isUnused)
    .sort((a, b) => new Date(b.lastUsedAt) - new Date(a.lastUsedAt))
    .slice(0, 5);

  const [width, setWidth] = React.useState(() => typeof window !== "undefined" ? window.innerWidth : 850);

  React.useEffect(() => {
    const handleResize = () => {
      setWidth(window.innerWidth);
    };
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  const isNarrow = width < 450;
  const isMedium = width >= 450 && width < 720;
  const isWide = width >= 720;

  return (
    <div style={{ display: "flex", minHeight: 80, height: 80, overflowX: "auto", overflowY: "hidden", backgroundColor: "#f3f2f1", borderBottom: "1px solid #edebe9", padding: "0", boxSizing: "border-box", alignItems: "center" }}>
      
      <RibbonGroup label="File Email">
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <button 
              disabled={!isAuthOk}
              style={{ 
                display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start",
                background: "transparent", border: "1px solid transparent", cursor: isAuthOk ? "pointer" : "not-allowed", 
                padding: "2px 4px", minWidth: 64, boxSizing: "border-box",
                opacity: isAuthOk ? 1 : 0.45
              }}
              title={!isAuthOk ? "Sign in to file emails" : ""}
              onMouseOver={(e) => isAuthOk && Object.assign(e.currentTarget.style, { backgroundColor: "#c1ddf1", border: "1px solid #7cbbed" })}
              onMouseOut={(e) => isAuthOk && Object.assign(e.currentTarget.style, { backgroundColor: "transparent", border: "1px solid transparent" })}
            >
              <div style={{ color: "#0078d4", marginBottom: 2, display: "flex", alignItems: "center" }}>
                <AppsListDetail24Regular />
                <ChevronDown24Regular style={{ fontSize: 12, marginLeft: 2 }} />
              </div>
              <span style={{ fontSize: 11, fontFamily: "'Exo 2', 'Segoe UI', sans-serif", textAlign: "center", lineHeight: "1.1", color: "#323130" }}>
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
                    disabled={!isAuthOk}
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
                    disabled={!isAuthOk}
                  >
                    {loc.description || loc.path.split("\\").pop()}
                  </MenuItem>
                ))
              ) : (
                <div style={{ padding: "4px 32px", fontSize: 11, color: "#a19f9d", fontStyle: "italic" }}>No recently used locations</div>
              )}

              <MenuDivider />
              <MenuItem 
                icon={<ArrowClockwise24Regular style={{ color: "#0078d4" }} />}
                persistOnClick
                onClick={(e) => {
                  e.stopPropagation();
                  onRefresh();
                }}
              >
                Refresh
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </RibbonGroup>

      {isNarrow && (
        <>
          <RibbonGroup label="Actions">
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <button 
                  style={{ 
                    display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start",
                    background: "transparent", border: "1px solid transparent", cursor: "pointer", 
                    padding: "2px 4px", minWidth: 48, boxSizing: "border-box"
                  }}
                  onMouseOver={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "#c1ddf1", border: "1px solid #7cbbed" })}
                  onMouseOut={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "transparent", border: "1px solid transparent" })}
                >
                  <div style={{ color: "#0078d4", marginBottom: 2, display: "flex", alignItems: "center" }}>
                    <MoreHorizontal24Regular />
                  </div>
                  <span style={{ fontSize: 11, fontFamily: "'Exo 2', 'Segoe UI', sans-serif", textAlign: "center", lineHeight: "1.1", color: "#323130" }}>
                    More
                  </span>
                </button>
              </MenuTrigger>
              <MenuPopover>
                <MenuList style={{ minWidth: 200 }}>
                  <MenuItem icon={<Add24Regular style={{ color: "#107c10" }} />} onClick={onAdd}>Add</MenuItem>
                  <MenuItem icon={<Edit24Regular style={{ color: "#d83b01" }} />} onClick={onEdit}>Edit</MenuItem>
                  <MenuItem icon={<Delete24Regular style={{ color: "#a4262c" }} />} onClick={onDelete}>Delete</MenuItem>
                  <MenuDivider />
                  <MenuItem icon={<FolderOpen24Regular style={{ color: "#0078d4" }} />} onClick={onExplore}>Explore</MenuItem>
                  <MenuItem icon={<ArrowClockwise24Regular style={{ color: "#008272" }} />} onClick={onRefresh}>Refresh</MenuItem>
                  <MenuDivider />
                  <MenuItem icon={<Star24Regular style={{ color: "#ffb900" }} />} onClick={onRemoveSuggestion}>Set as favourite</MenuItem>
                  <MenuItem 
                    icon={hasUnusedSelected ? <Eye24Regular style={{ color: "#881798" }} /> : <EyeOff24Regular style={{ color: "#881798" }} />} 
                    onClick={onMarkUnused}
                  >
                    {hasUnusedSelected ? "Set location used" : "Set location unused"}
                  </MenuItem>
                  <MenuItem 
                    icon={<SelectAllOn24Regular style={isMultiSelect ? {color: "#107c10"} : {color: "#605e5c"}} />} 
                    onClick={onToggleMultiSelect}
                  >
                    Choose multiple locations
                  </MenuItem>
                  <MenuDivider />
                  <MenuItem icon={<QuestionCircle24Regular style={{ color: "#0078d4" }} />} onClick={onHelp}>Help</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </RibbonGroup>
        </>
      )}

      {isMedium && (
        <>
          <RibbonGroup label="Actions">
            <RibbonButton icon={<Add24Regular style={{ color: "#107c10" }} />} label="Add" onClick={onAdd} />
            <RibbonButton icon={<Edit24Regular style={{ color: "#d83b01" }} />} label="Edit" onClick={onEdit} />
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <button 
                  style={{ 
                    display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start",
                    background: "transparent", border: "1px solid transparent", cursor: "pointer", 
                    padding: "2px 4px", minWidth: 48, boxSizing: "border-box"
                  }}
                  onMouseOver={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "#c1ddf1", border: "1px solid #7cbbed" })}
                  onMouseOut={(e) => Object.assign(e.currentTarget.style, { backgroundColor: "transparent", border: "1px solid transparent" })}
                >
                  <div style={{ color: "#0078d4", marginBottom: 2, display: "flex", alignItems: "center" }}>
                    <MoreHorizontal24Regular />
                  </div>
                  <span style={{ fontSize: 11, fontFamily: "'Exo 2', 'Segoe UI', sans-serif", textAlign: "center", lineHeight: "1.1", color: "#323130" }}>
                    More
                  </span>
                </button>
              </MenuTrigger>
              <MenuPopover>
                <MenuList style={{ minWidth: 200 }}>
                  <MenuItem icon={<Delete24Regular style={{ color: "#a4262c" }} />} onClick={onDelete}>Delete</MenuItem>
                  <MenuItem icon={<FolderOpen24Regular style={{ color: "#0078d4" }} />} onClick={onExplore}>Explore</MenuItem>
                  <MenuItem icon={<ArrowClockwise24Regular style={{ color: "#008272" }} />} onClick={onRefresh}>Refresh</MenuItem>
                  <MenuItem icon={<Star24Regular style={{ color: "#ffb900" }} />} onClick={onRemoveSuggestion}>Set as favourite</MenuItem>
                  <MenuItem 
                    icon={hasUnusedSelected ? <Eye24Regular style={{ color: "#881798" }} /> : <EyeOff24Regular style={{ color: "#881798" }} />} 
                    onClick={onMarkUnused}
                  >
                    {hasUnusedSelected ? "Set location used" : "Set location unused"}
                  </MenuItem>
                  <MenuItem 
                    icon={<SelectAllOn24Regular style={isMultiSelect ? {color: "#107c10"} : {color: "#605e5c"}} />} 
                    onClick={onToggleMultiSelect}
                  >
                    Choose multiple locations
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </RibbonGroup>

          <RibbonGroup label="Help">
            <RibbonButton 
              icon={<QuestionCircle24Regular style={{ color: "#0078d4" }} />} 
              label="Help" 
              onClick={onHelp}
            />
          </RibbonGroup>
        </>
      )}

      {isWide && (
        <>
          <RibbonGroup label="Actions">
            <RibbonButton icon={<Add24Regular style={{ color: "#107c10" }} />} label="Add" onClick={onAdd} />
            <RibbonButton icon={<Edit24Regular style={{ color: "#d83b01" }} />} label="Edit" onClick={onEdit} />
            <RibbonButton icon={<Delete24Regular style={{ color: "#a4262c" }} />} label="Delete" onClick={onDelete} />
            <RibbonButton icon={<FolderOpen24Regular style={{ color: "#0078d4" }} />} label="Explore" onClick={onExplore} />
            <RibbonButton icon={<ArrowClockwise24Regular style={{ color: "#008272" }} />} label="Refresh" onClick={onRefresh} />
            <RibbonButton icon={<Star24Regular style={{ color: "#ffb900" }}/>} label={<>Set as<br/>favourite</>} onClick={onRemoveSuggestion} />
            <RibbonButton icon={hasUnusedSelected ? <Eye24Regular style={{ color: "#881798" }} /> : <EyeOff24Regular style={{ color: "#881798" }} />} label={hasUnusedSelected ? <>Set location<br/>used</> : <>Set location<br/>unused</>} onClick={onMarkUnused} />
            <RibbonButton icon={<SelectAllOn24Regular style={isMultiSelect ? {color: "#107c10"} : {color: "#605e5c"}}/>} label={<>Choose multiple<br/>locations</>} onClick={onToggleMultiSelect} />
          </RibbonGroup>

          <RibbonGroup label="Help">
            <RibbonButton 
              icon={<QuestionCircle24Regular style={{ color: "#0078d4" }} />} 
              label={<>Koyomail<br/>help</>} 
              onClick={onHelp}
            />
          </RibbonGroup>
        </>
      )}

      {/* Brand — high-res PNG via webpack; explicit px box so IE/WebView cannot shrink the mark */}
      <div
        style={{
          marginLeft: "auto",
          display: "flex",
          alignItems: "center",
          justifyContent: "flex-end",
          flexShrink: 0,
          gap: 8,
          padding: "0 16px",
          backgroundColor: "transparent",
        }}
      >
        <div
          style={{
            width: isWide ? 58 : 48,
            height: isWide ? 58 : 48,
            minWidth: isWide ? 58 : 48,
            minHeight: isWide ? 58 : 48,
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
              width: "100%",
              height: "100%",
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
            fontSize: isWide ? 20 : 18,
            fontWeight: 700,
            color: "#000000",
            fontFamily: "'Exo 2', 'Segoe UI', sans-serif",
            lineHeight: 1.1,
            flexShrink: 0,
            letterSpacing: "1px",
            textTransform: "uppercase",
          }}
        >
          Koyomail
        </span>
      </div>

    </div>
  );
};

export default Toolbar;
