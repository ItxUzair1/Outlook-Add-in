import * as React from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  Option,
  RadioGroup,
  Radio
} from "@fluentui/react-components";

const OptionsDialog = ({ isOpen, onOpenChange, initialTab = "Local & Network folders" }) => {
  const [selectedMainTab, setSelectedMainTab] = React.useState(initialTab);
  
  React.useEffect(() => {
    if (isOpen) {
      setSelectedMainTab(initialTab);
    }
  }, [isOpen, initialTab]);
  
  // State for Local & Network folders
  const [discoverLocations, setDiscoverLocations] = React.useState(false);
  const [applyReadOnly, setApplyReadOnly] = React.useState(false);
  const [pathType, setPathType] = React.useState("UNC");

  // State for Search
  const [enableSearching, setEnableSearching] = React.useState(true);
  const [searchScope, setSearchScope] = React.useState("locations_i_use");
  const [disableDelete, setDisableDelete] = React.useState(false);
  const [disableMoveTo, setDisableMoveTo] = React.useState(false);

  React.useEffect(() => {
    try {
      const stored = localStorage.getItem('koyomail_options');
      if (stored) {
        const parsed = JSON.parse(stored);
        setEnableSearching(parsed.enableSearching ?? true);
        setSearchScope(parsed.searchScope || "locations_i_use");
        setDisableDelete(parsed.disableDelete || false);
        setDisableMoveTo(parsed.disableMoveTo || false);
      }
    } catch (e) {
      console.error(e);
    }
  }, []);

  const updateOption = (key, value) => {
    try {
      const stored = localStorage.getItem('koyomail_options');
      let current = stored ? JSON.parse(stored) : { enableSearching: true, disableDelete: false, disableMoveTo: false, searchScope: "locations_i_use" };
      current[key] = value;
      localStorage.setItem('koyomail_options', JSON.stringify(current));
      window.dispatchEvent(new Event('koyomail_options_updated'));
    } catch (e) {
      console.error(e);
    }
  };

  const urlMode = new URLSearchParams(window.location.search).get("mode");
  const isFromSearch = urlMode === "search";

  if (!isOpen) return null;

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 10000, display: "flex", flexDirection: "column", height: "100vh", backgroundColor: "#ffffff" }}>
      <div style={{ display: "flex", flexGrow: 1, backgroundColor: "#ffffff", overflow: "hidden" }}>
            {/* Sidebar */}
            <div style={{ width: "220px", borderRight: "1px solid #edebe9", padding: "16px 0", backgroundColor: "#ffffff" }}>
              {isFromSearch && (
                <div 
                  style={{ 
                    padding: "8px 16px 16px 16px", 
                    fontSize: "14px", 
                    cursor: "pointer", 
                    color: "#0078d4",
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                    fontWeight: "600",
                    borderBottom: "1px solid #edebe9",
                    marginBottom: "16px"
                  }} 
                  onClick={() => onOpenChange(false)}
                >
                  <span style={{ fontSize: "16px" }}>←</span> Back to Search
                </div>
              )}
              
              <div style={{ padding: "8px 16px", fontWeight: "600", fontSize: "16px" }}>General</div>
              <div 
                style={{ 
                  padding: "8px 16px 8px 32px", 
                  fontSize: "14px", 
                  cursor: "pointer", 
                  backgroundColor: selectedMainTab === "Filing" ? "#d2e4f5" : "transparent" 
                }} 
                onClick={() => setSelectedMainTab("Filing")}
              >
                Filing
              </div>
              <div 
                style={{ 
                  padding: "8px 16px 8px 32px", 
                  fontSize: "14px", 
                  cursor: "pointer", 
                  backgroundColor: selectedMainTab === "Search" ? "#d2e4f5" : "transparent" 
                }} 
                onClick={() => setSelectedMainTab("Search")}
              >
                Search
              </div>
              
              <div style={{ padding: "8px 16px", fontWeight: "600", fontSize: "16px", marginTop: "16px" }}>Integrations</div>
              <div 
                style={{ 
                  padding: "8px 16px 8px 32px", 
                  fontSize: "14px", 
                  cursor: "pointer", 
                  backgroundColor: selectedMainTab === "Local & Network folders" ? "#cceaea" : "transparent" 
                }} 
                onClick={() => setSelectedMainTab("Local & Network folders")}
              >
                Local &amp; Network folders
              </div>
            </div>
            
            {/* Content Area */}
            <div style={{ flexGrow: 1, padding: "24px 32px", backgroundColor: "#fbfbfb", overflowY: "auto" }}>
              {selectedMainTab === "Local & Network folders" && (
                <div>
                  <h2 style={{ fontSize: "16px", fontWeight: "600", marginBottom: "24px", color: "#323130", margin: "0 0 24px 0" }}>
                    Integrations - Local and network folders
                  </h2>
                  
                  <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                    <Checkbox 
                      label="Discover filing locations" 
                      checked={discoverLocations}
                      onChange={(e, data) => setDiscoverLocations(data.checked)}
                    />
                    <Checkbox 
                      label="Apply file system 'Read only' attribute to filed items" 
                      checked={applyReadOnly}
                      onChange={(e, data) => setApplyReadOnly(data.checked)}
                    />
                    
                    <div style={{ marginTop: "12px" }}>
                      <label style={{ display: "block", marginBottom: "8px", fontSize: "14px", color: "#323130" }}>Default path type:</label>
                      <Dropdown 
                        value={pathType} 
                        style={{ minWidth: "150px" }}
                        selectedOptions={[pathType]} 
                        onOptionSelect={(e, data) => setPathType(data.optionValue)}
                      >
                        <Option value="UNC">UNC</Option>
                        <Option value="Drive">Drive</Option>
                      </Dropdown>
                    </div>
                  </div>
                </div>
              )}
              {selectedMainTab === "Filing" && (
                <div>
                  <h2 style={{ fontSize: "16px", fontWeight: "600", marginBottom: "24px", color: "#323130", margin: "0 0 24px 0" }}>
                    General - Filing
                  </h2>
                  <p style={{ color: "#605e5c", fontSize: "14px" }}>Filing options...</p>
                </div>
              )}
              {selectedMainTab === "Search" && (
                <div>
                  <h2 style={{ fontSize: "16px", fontWeight: "600", color: "#323130", margin: "0 0 8px 0" }}>
                    Search
                  </h2>
                  <p style={{ color: "#323130", fontSize: "13px", margin: "0 0 24px 0" }}>
                    These settings apply to all integrations
                  </p>

                  <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
                    <Checkbox 
                      label="Enable searching" 
                      checked={enableSearching}
                      onChange={(e, data) => { setEnableSearching(data.checked); updateOption('enableSearching', data.checked); }}
                    />
                    
                    <div style={{ paddingLeft: "28px", opacity: enableSearching ? 1 : 0.6, pointerEvents: enableSearching ? "auto" : "none" }}>
                      <RadioGroup value={searchScope} onChange={(e, data) => { setSearchScope(data.value); updateOption('searchScope', data.value); }}>
                        <Radio 
                          value="locations_i_use" 
                          label="Only search locations I use" 
                          style={{ marginBottom: "8px" }}
                        />
                        
                        <Radio 
                          value="all_locations" 
                          label="Search all available locations" 
                        />
                        <div style={{ marginLeft: "28px", color: "#d13438", fontSize: "12px", marginTop: "-4px" }}>
                          Searches will take longer with this option.
                        </div>
                      </RadioGroup>
                    </div>

                    <div style={{ marginTop: "16px", color: "#323130", fontSize: "13px" }}>
                      Search window
                    </div>
                    
                    <Checkbox 
                      label="Disable the Delete option" 
                      checked={disableDelete}
                      onChange={(e, data) => { setDisableDelete(data.checked); updateOption('disableDelete', data.checked); }}
                      style={{ marginTop: "-4px" }}
                    />
                    <Checkbox 
                      label="Disable the Move to option" 
                      checked={disableMoveTo}
                      onChange={(e, data) => { setDisableMoveTo(data.checked); updateOption('disableMoveTo', data.checked); }}
                      style={{ marginTop: "-8px" }}
                    />
                  </div>
                </div>
              )}
            </div>
          </div>
          
          <div style={{ padding: "16px 24px", borderTop: "1px solid #edebe9", backgroundColor: "#ffffff", display: "flex", justifyContent: "flex-end" }}>
        <Button appearance="secondary" onClick={() => onOpenChange(false)}>
          {isFromSearch ? "Back to Search" : "Close"}
        </Button>
      </div>
    </div>
  );
};

export default OptionsDialog;
