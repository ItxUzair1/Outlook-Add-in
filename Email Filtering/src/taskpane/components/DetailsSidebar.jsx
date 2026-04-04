import * as React from "react";
import { Input, Label, Select, Checkbox } from "@fluentui/react-components";

const DetailsSidebar = ({ 
  subject, setSubject, 
  comment, setComment, 
  afterFiling, setAfterFiling, 
  markReviewed, setMarkReviewed, 
  sendLink, setSendLink, 
  attachmentsOption, setAttachmentsOption,
  onSaveDefaults
}) => {
  const [isOptionsExpanded, setIsOptionsExpanded] = React.useState(false);

  return (
    <div style={{ flex: "0 0 260px", borderLeft: "1px solid #edebe9", display: "flex", flexDirection: "column", backgroundColor: "#faf9f8", overflowY: "auto", overflowX: "hidden" }}>
      
      {/* Sidebar Header - Selected Email */}
      <div style={{ padding: "12px 12px 0 12px", display: "flex", flexDirection: "column", gap: 4 }}>
        <div 
          onClick={() => setIsOptionsExpanded(!isOptionsExpanded)}
          style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", userSelect: "none" }}
        >
          <span style={{ fontSize: 10 }}>{isOptionsExpanded ? "▼" : "▶"}</span>
          <Label size="small" weight="semibold" style={{ fontSize: 13, fontFamily: "Segoe UI", cursor: "pointer" }}>Selected Email</Label>
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
          
          <Label size="small">Actions after filing:</Label>
          <Select size="small" value={afterFiling} onChange={(e) => setAfterFiling(e.target.value)}>
            <option value="none">Keep in Inbox</option>
            <option value="delete">Transfer email to Deleted Items</option>
            <option value="folder">Transfer email to Outlook folder</option>
            <option value="archive">Archive</option>
          </Select>

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
};

export default DetailsSidebar;
