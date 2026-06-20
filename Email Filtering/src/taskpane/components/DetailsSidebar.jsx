import * as React from "react";
import { Input, Label, Select, Checkbox } from "@fluentui/react-components";
import { ArrowLeft20Regular } from "@fluentui/react-icons";

const DetailsSidebar = ({ 
  subject, setSubject, 
  comment, setComment, 
  afterFiling, setAfterFiling, 
  markReviewed, setMarkReviewed, 
  sendLink, setSendLink, 
  attachmentsOption, setAttachmentsOption,
  onSaveDefaults,
  mode,
  isNarrow,
  onBack
}) => {
  const [isOptionsExpanded, setIsOptionsExpanded] = React.useState(false);

  const overlayStyle = {
    position: "absolute", top: 0, left: 0, right: 0, bottom: 0, zIndex: 1000, 
    backgroundColor: "rgba(0, 0, 0, 0.4)", backdropFilter: "blur(2px)", 
    display: "flex", justifyContent: "flex-end"
  };

  const containerStyle = isNarrow
    ? { width: "75%", height: "100%", display: "flex", flexDirection: "column", backgroundColor: "#faf9f8", overflowY: "auto", overflowX: "hidden", boxShadow: "-4px 0 16px rgba(0,0,0,0.2)" }
    : { flex: "0 0 280px", borderLeft: "1px solid #edebe9", display: "flex", flexDirection: "column", backgroundColor: "#ffffff", overflowY: "auto", overflowX: "hidden", boxShadow: "-2px 0 12px rgba(0,0,0,0.05)" };

  const content = (
    <div style={containerStyle} onClick={isNarrow ? (e) => e.stopPropagation() : undefined}>
      {/* Sidebar Header - Selected Email */}
      <div style={{ padding: "12px 12px 0 12px", display: "flex", flexDirection: "column", gap: 4 }}>
        <div style={{ display: "flex", alignItems: "center" }}>
          {isNarrow && (
            <button 
              onClick={(e) => { e.stopPropagation(); onBack(); }}
              style={{ display: "flex", alignItems: "center", justifyContent: "center", border: "none", background: "none", cursor: "pointer", color: "#0078d4", padding: "0 8px 0 0" }}
              title="Back"
            >
              <ArrowLeft20Regular />
            </button>
          )}
          <div 
            onClick={() => setIsOptionsExpanded(!isOptionsExpanded)}
            style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", userSelect: "none" }}
          >
            <span style={{ fontSize: 10 }}>{isOptionsExpanded ? "▼" : "▶"}</span>
            <Label size="small" weight="semibold" style={{ fontSize: 13, fontFamily: "'Exo 2', 'Segoe UI', sans-serif", cursor: "pointer" }}>Selected Email</Label>
          </div>
        </div>
        {isOptionsExpanded && (
          <div style={{ fontSize: 12, marginLeft: 16 }}>
            <span 
              onClick={onSaveDefaults}
              style={{ color: "#0078d4", textDecoration: "underline", cursor: "pointer", display: "inline-block", marginBottom: 4 }}
            >
              Change defaults
            </span>
            <div style={{ color: "#605e5c" }}>Select 'Change defaults' to remember your choice.</div>
          </div>
        )}
      </div>

      <div style={{ padding: "12px", display: "flex", flexDirection: "column", gap: 8 }}>
        {/* All fields now visible in a unified list */}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <Label size="small" style={{ marginTop: 4 }}>Subject:</Label>
          <Input size="small" value={String(subject || "")} onChange={(e) => setSubject(e.target.value)} />
          
          <Label size="small">Comments:</Label>
          <Input size="small" value={String(comment || "")} onChange={(e) => setComment(e.target.value)} placeholder="Enter comment" />
          
          <>
            <Label size="small">Actions after filing:</Label>
            <Select size="small" value={afterFiling} onChange={(e) => setAfterFiling(e.target.value)}>
              <option value="none">{mode === "onsend" ? "Keep in Sent Items" : "Keep in Inbox"}</option>
              <option value="add_date">Add filed date to subject</option>
              <option value="delete">Transfer email to Deleted Items</option>
              <option value="move_filed_items">Transfer to Filed Items folder</option>
              <option value="move_filed_folders">Transfer to Filed sub-folders</option>
              <option value="archive">Archive</option>
            </Select>
          </>

          <div style={{ display: "flex", flexDirection: "column", gap: 4, marginTop: 4 }}>
            <Checkbox size="small" label="Email has been reviewed" checked={markReviewed} onChange={(e, data) => setMarkReviewed(data.checked)} />
            <Checkbox size="small" label="Generate email link" checked={sendLink} onChange={(e, data) => setSendLink(data.checked)} />
          </div>

          <Label size="small" style={{ marginTop: 8 }}>Filing options:</Label>
          <Select size="small" value={attachmentsOption} onChange={(e) => setAttachmentsOption(e.target.value)}>
            <option value="all">File message with attachments</option>
            <option value="message">File message only</option>
            <option value="attachments">File attachments only</option>
          </Select>
        </div>
      </div>
    </div>
  );

  return isNarrow ? (
    <div style={overlayStyle} onClick={onBack}>
      {content}
    </div>
  ) : content;
};

export default DetailsSidebar;
