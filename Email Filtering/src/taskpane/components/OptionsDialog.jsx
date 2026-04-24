import * as React from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  Option,
  RadioGroup,
  Radio,
  Input
} from "@fluentui/react-components";
import { API_BASE_URL, getPreferences, updatePreferences } from "../services/backendApi.js";

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
  const [discoverStatus, setDiscoverStatus] = React.useState("");

  // State for Search
  const [enableSearching, setEnableSearching] = React.useState(true);
  const [searchScope, setSearchScope] = React.useState("locations_i_use");
  const [disableDelete, setDisableDelete] = React.useState(false);
  const [disableMoveTo, setDisableMoveTo] = React.useState(false);

  // State for Filing
  const [duplicateStrategy, setDuplicateStrategy] = React.useState("rename");
  const [defaultAttachments, setDefaultAttachments] = React.useState("all");

  // --- New Filing Options State ---
  const [afterFilingAction, setAfterFilingAction] = React.useState("none");
  const [filedFolderPrefix, setFiledFolderPrefix] = React.useState("*");
  const [deleteEmptyFolders, setDeleteEmptyFolders] = React.useState(false);
  const [addFiledCategory, setAddFiledCategory] = React.useState(true);
  const [fileReplyingTo, setFileReplyingTo] = React.useState(false);
  const [sendLink, setSendLink] = React.useState(false);
  const [emailFont, setEmailFont] = React.useState("Times New Roman");
  const [fontSize, setFontSize] = React.useState("10");
  const [markReviewed, setMarkReviewed] = React.useState(false);
  const [assistantCategories, setAssistantCategories] = React.useState("");
  const [enableDoubleClickFiling, setEnableDoubleClickFiling] = React.useState(false);
  const [includeCollectionName, setIncludeCollectionName] = React.useState(false);
  const [onlyFileUsingDialog, setOnlyFileUsingDialog] = React.useState(false);
  const [alwaysShowFilingOptions, setAlwaysShowFilingOptions] = React.useState(false);
  const [useUtcTime, setUseUtcTime] = React.useState(false);

  // Load options from localStorage and merge backend preferences.
  React.useEffect(() => {
    const loadOptions = async () => {
      try {
        const stored = localStorage.getItem('koyomail_options');
        const localParsed = stored ? JSON.parse(stored) : {};
        let backendParsed = {};
        try {
          backendParsed = await getPreferences();
        } catch {
          backendParsed = {};
        }
        const parsed = { ...backendParsed, ...localParsed };
        // Search options
        setEnableSearching(parsed.enableSearching ?? true);
        setSearchScope(parsed.searchScope || "locations_i_use");
        setDisableDelete(parsed.disableDelete || false);
        setDisableMoveTo(parsed.disableMoveTo || false);
        // Local & Network folders options
        setDiscoverLocations(parsed.discoverLocations || false);
        setApplyReadOnly(parsed.applyReadOnly || false);
        setPathType(parsed.pathType || "UNC");
        // Filing options
        setDuplicateStrategy(parsed.duplicateStrategy || "rename");
        setDefaultAttachments(parsed.defaultAttachments || "all");
        
        if (parsed.afterFilingAction) {
          const normalizedAfter = parsed.afterFilingAction === "move_deleted" ? "delete" : parsed.afterFilingAction;
          setAfterFilingAction(normalizedAfter);
        }
        if (parsed.filedFolderPrefix !== undefined) setFiledFolderPrefix(parsed.filedFolderPrefix);
        if (parsed.deleteEmptyFolders !== undefined) setDeleteEmptyFolders(parsed.deleteEmptyFolders);
        if (parsed.addFiledCategory !== undefined) setAddFiledCategory(parsed.addFiledCategory);
        if (parsed.fileReplyingTo !== undefined) setFileReplyingTo(parsed.fileReplyingTo);
        if (parsed.sendLink !== undefined) setSendLink(parsed.sendLink);
        if (parsed.emailFont) setEmailFont(parsed.emailFont);
        if (parsed.fontSize) setFontSize(parsed.fontSize);
        if (parsed.markReviewed !== undefined) setMarkReviewed(parsed.markReviewed);
        
        if (parsed.assistantCategories !== undefined) setAssistantCategories(parsed.assistantCategories);

        if (parsed.enableDoubleClickFiling !== undefined) setEnableDoubleClickFiling(parsed.enableDoubleClickFiling);
        if (parsed.includeCollectionName !== undefined) setIncludeCollectionName(parsed.includeCollectionName);
        if (parsed.onlyFileUsingDialog !== undefined) setOnlyFileUsingDialog(parsed.onlyFileUsingDialog);
        if (parsed.alwaysShowFilingOptions !== undefined) setAlwaysShowFilingOptions(parsed.alwaysShowFilingOptions);
        if (parsed.useUtcTime !== undefined) setUseUtcTime(parsed.useUtcTime);
        localStorage.setItem('koyomail_options', JSON.stringify(parsed));
      } catch (e) {
        console.error(e);
      }
    };
    loadOptions();
  }, []);

  const updateOption = (key, value) => {
    try {
      const stored = localStorage.getItem('koyomail_options');
      let current = stored ? JSON.parse(stored) : {
        enableSearching: true,
        disableDelete: false,
        disableMoveTo: false,
        searchScope: "locations_i_use",
        discoverLocations: false,
        applyReadOnly: false,
        pathType: "UNC",
        duplicateStrategy: "rename",
        defaultAttachments: "all",
      };
      const normalizedValue = key === "afterFilingAction" && value === "move_deleted" ? "delete" : value;
      current[key] = normalizedValue;
      localStorage.setItem('koyomail_options', JSON.stringify(current));
      window.dispatchEvent(new Event('koyomail_options_updated'));
      updatePreferences({ [key]: normalizedValue }).catch((err) => {
        console.warn("Failed to persist preference to backend:", err?.message || err);
      });
    } catch (e) {
      console.error(e);
    }
  };

  // Discover filing locations — calls backend to scan the search index for unique directories
  const handleDiscoverLocations = async (enabled) => {
    setDiscoverLocations(enabled);
    updateOption('discoverLocations', enabled);

    if (enabled) {
      setDiscoverStatus("Scanning for filing locations...");
      try {
        const resp = await fetch(`${API_BASE_URL}/api/locations/discover`, { method: "POST" });
        if (resp.ok) {
          const data = await resp.json();
          setDiscoverStatus(`Done! Found ${data.addedCount} new location(s).`);
          setTimeout(() => setDiscoverStatus(""), 5000);
        } else {
          setDiscoverStatus("Discovery failed. Check server connection.");
          setTimeout(() => setDiscoverStatus(""), 5000);
        }
      } catch (e) {
        setDiscoverStatus(`Discovery failed: ${e.message}`);
        setTimeout(() => setDiscoverStatus(""), 5000);
      }
    } else {
      setDiscoverStatus("");
    }
  };

  const urlMode = new URLSearchParams(window.location.search).get("mode");
  const isFromSearch = urlMode === "search";

  if (!isOpen) return null;

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 10000, display: "flex", flexDirection: "column", height: "100vh", backgroundColor: "#ffffff" }}>
      <div style={{ display: "flex", flexGrow: 1, backgroundColor: "#ffffff", overflow: "hidden" }}>
            {/* Sidebar */}
            <div style={{ width: "220px", minWidth: "220px", flexShrink: 0, borderRight: "1px solid #edebe9", padding: "16px 0", backgroundColor: "#ffffff" }}>
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
            <div style={{ flexGrow: 1, minWidth: 0, padding: "24px 32px", backgroundColor: "#fbfbfb", overflowY: "auto", overflowX: "hidden" }}>
              {selectedMainTab === "Local & Network folders" && (
                <div>
                  <h2 style={{ fontSize: "16px", fontWeight: "600", marginBottom: "24px", color: "#323130", margin: "0 0 24px 0" }}>
                    Integrations - Local and network folders
                  </h2>
                  
                  <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                    <div>
                      <Checkbox 
                        label="Discover filing locations" 
                        checked={discoverLocations}
                        onChange={(e, data) => handleDiscoverLocations(data.checked)}
                      />
                      {discoverStatus && (
                        <div style={{ marginLeft: "28px", marginTop: "4px", fontSize: "12px", color: discoverStatus.includes("failed") ? "#a4262c" : "#107c10", fontWeight: "600" }}>
                          {discoverStatus}
                        </div>
                      )}
                      <div style={{ marginLeft: "28px", marginTop: "4px", fontSize: "12px", color: "#605e5c" }}>
                        Automatically scan filed emails to discover and add filing locations.
                      </div>
                    </div>
                    <div>
                      <Checkbox 
                        label="Apply file system 'Read only' attribute to filed items" 
                        checked={applyReadOnly}
                        onChange={(e, data) => { setApplyReadOnly(data.checked); updateOption('applyReadOnly', data.checked); }}
                      />
                      <div style={{ marginLeft: "28px", marginTop: "4px", fontSize: "12px", color: "#605e5c" }}>
                        Prevents accidental modification of filed emails by marking them as read-only on disk.
                      </div>
                    </div>
                    
                    <div style={{ marginTop: "12px" }}>
                      <label style={{ display: "block", marginBottom: "8px", fontSize: "14px", color: "#323130" }}>Default path type:</label>
                      <Dropdown 
                        value={pathType} 
                        style={{ minWidth: "150px" }}
                        selectedOptions={[pathType]} 
                        onOptionSelect={(e, data) => { setPathType(data.optionValue); updateOption('pathType', data.optionValue); }}
                      >
                        <Option value="UNC">UNC</Option>
                        <Option value="Drive">Drive</Option>
                      </Dropdown>
                      <div style={{ marginTop: "8px", fontSize: "12px", color: "#605e5c" }}>
                        {pathType === "UNC" 
                          ? "UNC paths use the format \\\\server\\share\\folder for network locations." 
                          : "Drive paths use the format C:\\folder for local and mapped drive locations."
                        }
                      </div>
                    </div>
                  </div>
                </div>
              )}
              {selectedMainTab === "Filing" && (
                <div>
                  <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>

                  <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                    <h3 style={{ fontSize: "14px", fontWeight: "600", margin: 0, paddingBottom: "4px" }}>After filing</h3>
                    <RadioGroup value={afterFilingAction} onChange={(e, d) => { setAfterFilingAction(d.value); updateOption('afterFilingAction', d.value); }}>
                      <Radio value="none" label="Keep in Inbox" />
                      <Radio value="add_date" label="Add filed date and time to subject, but don't move" />
                      <Radio value="delete" label="Move to Deleted Items folder" />
                      <Radio value="move_filed_items" label="Move to Filed Items, an Inbox sub-folder" />
                      <Radio value="move_filed_folders" label="Move to Filed folders, multiple Inbox sub-folders with the same description as the filing location" />
                      <Radio value="archive" label="Archive" />
                    </RadioGroup>
                    
                    <div style={{ paddingLeft: "28px", display: "flex", alignItems: "center", gap: "8px", opacity: afterFilingAction === "move_filed_folders" ? 1 : 0.5, pointerEvents: afterFilingAction === "move_filed_folders" ? "auto" : "none" }}>
                      <span style={{ fontSize: "12px", color: "#605e5c" }}>Filed folder Prefix:</span>
                      <Input value={filedFolderPrefix} onChange={(e, d) => { setFiledFolderPrefix(d.value); updateOption('filedFolderPrefix', d.value); }} style={{ width: "60px", height: "24px", minHeight: "24px" }} />
                    </div>
                    <div style={{ paddingLeft: "28px", opacity: afterFilingAction === "move_filed_folders" ? 1 : 0.5, pointerEvents: afterFilingAction === "move_filed_folders" ? "auto" : "none" }}>
                      <Checkbox label="Delete empty Filed folders" checked={deleteEmptyFolders} onChange={(e, d) => { setDeleteEmptyFolders(d.checked); updateOption('deleteEmptyFolders', d.checked); }} />
                    </div>

                    <Checkbox label="Add filed category" checked={addFiledCategory} onChange={(e, d) => { setAddFiledCategory(d.checked); updateOption('addFiledCategory', d.checked); }} />
                    <Checkbox label="And file the message I'm replying to" checked={fileReplyingTo} onChange={(e, d) => { setFileReplyingTo(d.checked); updateOption('fileReplyingTo', d.checked); }} />
                    
                    <Checkbox label="Send a link to the filed item after filing" checked={sendLink} onChange={(e, d) => { setSendLink(d.checked); updateOption('sendLink', d.checked); }} />
                    <div style={{ paddingLeft: "28px", display: "flex", alignItems: "center", gap: "16px", opacity: sendLink ? 1 : 0.5, pointerEvents: sendLink ? "auto" : "none" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                        <span style={{ fontSize: "12px", color: "#605e5c" }}>Email font</span>
                        <Dropdown value={emailFont} selectedOptions={[emailFont]} onOptionSelect={(e, d) => { setEmailFont(d.optionValue); updateOption('emailFont', d.optionValue); }} style={{ minWidth: "150px" }}>
                          <Option value="Times New Roman">Times New Roman</Option>
                          <Option value="Arial">Arial</Option>
                          <Option value="Calibri">Calibri</Option>
                        </Dropdown>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                        <span style={{ fontSize: "12px", color: "#605e5c" }}>Font size</span>
                        <Dropdown value={fontSize} selectedOptions={[fontSize]} onOptionSelect={(e, d) => { setFontSize(d.optionValue); updateOption('fontSize', d.optionValue); }} style={{ minWidth: "60px" }}>
                          <Option value="8">8</Option>
                          <Option value="10">10</Option>
                          <Option value="12">12</Option>
                          <Option value="14">14</Option>
                        </Dropdown>
                      </div>
                    </div>

                    <Checkbox label="Mark item subject as reviewed" checked={markReviewed} onChange={(e, d) => { setMarkReviewed(d.checked); updateOption('markReviewed', d.checked); }} />
                  </div>



                  <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                    <h3 style={{ fontSize: "14px", fontWeight: "600", margin: 0, paddingBottom: "4px" }}>Advanced Settings</h3>

                    <div style={{ marginTop: "16px" }}>
                      <div style={{ fontWeight: "600", fontSize: "14px" }}>Attachments</div>
                      <div style={{ fontSize: "12px", color: "#605e5c", marginBottom: "4px" }}>Select default from the menu:</div>
                      <Dropdown 
                        value={defaultAttachments === "message" ? "File message only" : defaultAttachments === "attachments" ? "File attachments separately" : "File message with attachments"} 
                        selectedOptions={[defaultAttachments === "message" ? "File message only" : defaultAttachments === "attachments" ? "File attachments separately" : "File message with attachments"]} 
                        onOptionSelect={(e, d) => { 
                          const val = d.optionValue === "File message only" ? "message" : d.optionValue === "File attachments separately" ? "attachments" : "all";
                          setDefaultAttachments(val); 
                          updateOption('defaultAttachments', val); 
                        }} 
                        style={{ minWidth: "250px" }}
                      >
                        <Option value="File message with attachments">File message with attachments</Option>
                        <Option value="File attachments separately">File attachments separately</Option>
                        <Option value="File message only">File message only</Option>
                      </Dropdown>
                    </div>

                    <div style={{ marginTop: "16px" }}>
                      <div style={{ fontWeight: "600", fontSize: "14px" }}>Duplicate handling</div>
                      <div style={{ fontSize: "12px", color: "#605e5c", marginBottom: "4px" }}>When the same file name already exists:</div>
                      <Dropdown
                        value={duplicateStrategy === "skip" ? "Skip" : duplicateStrategy === "overwrite" ? "Overwrite" : "Rename"}
                        selectedOptions={[duplicateStrategy === "skip" ? "Skip" : duplicateStrategy === "overwrite" ? "Overwrite" : "Rename"]}
                        onOptionSelect={(e, d) => {
                          const val = d.optionValue === "Skip" ? "skip" : d.optionValue === "Overwrite" ? "overwrite" : "rename";
                          setDuplicateStrategy(val);
                          updateOption("duplicateStrategy", val);
                        }}
                        style={{ minWidth: "180px" }}
                      >
                        <Option value="Rename">Rename</Option>
                        <Option value="Skip">Skip</Option>
                        <Option value="Overwrite">Overwrite</Option>
                      </Dropdown>
                    </div>

                    <div style={{ marginTop: "16px" }}>
                      <div style={{ fontWeight: "600", fontSize: "14px" }}>Categories</div>
                      <div style={{ display: "flex", gap: "8px", marginTop: "4px", alignItems: "center" }}>
                        <Input value={assistantCategories} onChange={(e, d) => { setAssistantCategories(d.value); updateOption('assistantCategories', d.value); }} style={{ flexGrow: 1 }} />
                        <Button onClick={() => updateOption("assistantCategories", assistantCategories)}>Update...</Button>
                      </div>
                    </div>



                    <div style={{ marginTop: "16px" }}>
                      <div style={{ fontWeight: "600", fontSize: "14px", marginBottom: "8px" }}>Other settings</div>
                      <Checkbox label="Enable filing by double clicking desired location in the Koyomail filing dialog" checked={enableDoubleClickFiling} onChange={(e, d) => { setEnableDoubleClickFiling(d.checked); updateOption('enableDoubleClickFiling', d.checked); }} />
                      <Checkbox label="Include the collection name when listing filing locations (requires restarting Outlook)" checked={includeCollectionName} onChange={(e, d) => { setIncludeCollectionName(d.checked); updateOption('includeCollectionName', d.checked); }} />
                      
                      <Checkbox label="Only file using the filing dialog (requires restarting Outlook)" checked={onlyFileUsingDialog} onChange={(e, d) => { setOnlyFileUsingDialog(d.checked); updateOption('onlyFileUsingDialog', d.checked); }} />
                      <div style={{ paddingLeft: "28px", opacity: onlyFileUsingDialog ? 1 : 0.5, pointerEvents: onlyFileUsingDialog ? "auto" : "none" }}>
                        <Checkbox label="Always show filing options" checked={alwaysShowFilingOptions} onChange={(e, d) => { setAlwaysShowFilingOptions(d.checked); updateOption('alwaysShowFilingOptions', d.checked); }} />
                      </div>

                      <Checkbox label="Use UTC filing time" checked={useUtcTime} onChange={(e, d) => { setUseUtcTime(d.checked); updateOption('useUtcTime', d.checked); }} />
                    </div>
                  </div>

                  </div>
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
