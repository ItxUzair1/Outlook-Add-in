import * as React from "react";
import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogBody,
  DialogActions,
  Button,
} from "@fluentui/react-components";
import {
  QuestionCircle24Regular,
  Star24Regular,
  Add24Regular,
  Edit24Regular,
  Delete24Regular,
  FolderOpen24Regular,
  EyeOff24Regular,
  StarOff24Regular,
  SelectAllOn24Regular,
} from "@fluentui/react-icons";

const HelpSection = ({ title, icon, children }) => (
  <div style={{ marginBottom: "20px" }}>
    <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "8px", borderBottom: "1px solid #edebe9", paddingBottom: "4px" }}>
      <div style={{ color: "#0078d4" }}>{icon}</div>
      <h3 style={{ margin: 0, fontSize: "16px", fontWeight: "600", color: "#323130" }}>{title}</h3>
    </div>
    <div style={{ paddingLeft: "32px", fontSize: "14px", color: "#605e5c", lineHeight: "1.5" }}>
      {children}
    </div>
  </div>
);

const HelpDialog = ({ isOpen, onOpenChange }) => {
  return (
    <Dialog open={isOpen} onOpenChange={(e, data) => onOpenChange(data.open)}>
      <DialogSurface style={{ maxWidth: "600px", width: "95%" }}>
        <DialogBody>
          <DialogTitle>Filing Help & Documentation</DialogTitle>
          <DialogContent style={{ maxHeight: "70vh", overflowY: "auto", paddingRight: "10px" }}>
            
            <HelpSection title="How to File an Email (Step-by-Step)" icon={<QuestionCircle24Regular />}>
              <p>Follow these simple steps to file your current email:</p>
              <div style={{ backgroundColor: "#eff6fc", padding: "12px", borderRadius: "4px", marginBottom: "12px", borderLeft: "4px solid #0078d4" }}>
                <ol style={{ paddingLeft: "20px", margin: 0 }}>
                  <li><b>Search or Select Folder</b>: Use the table on the left to find your destination folder. You can use the search bar at the top to filter by name.</li>
                  <li><b>Click the Checkbox</b>: Tick the box next to the folder(s) you want to file into.</li>
                  <li><b>Verify Details (Right Panel)</b>: Check the Subject and Comment in the right sidebar. Adjust the "After Filing" action (e.g. Move to Deleted Items) if needed.</li>
                  <li><b>Press FILE</b>: Click the large blue <b>"File"</b> button at the bottom right.</li>
                </ol>
              </div>
              <p>Once finished, you will see a green "Email filed successfully" confirmation at the bottom.</p>
            </HelpSection>

            <HelpSection title="Filing Options & Sidebar" icon={<Edit24Regular />}>
              <p>The right sidebar allows you to customize the filing process:</p>
              <ul>
                <li><b>After Filing</b>: 
                  <ul>
                    <li><i>Keep in Inbox</i>: Leaves the original email alone.</li>
                    <li><i>Move to Deleted Items</i>: Deletes the email from Outlook after filing.</li>
                    <li><i>Archive</i>: Moves the email to your Outlook Archive folder.</li>
                  </ul>
                </li>
                <li><b>Mark message as reviewed</b>: Adds a metadata tag to the saved file for tracking.</li>
                <li><b>Attachments</b>: Use the "Options" dropdown in the sidebar to choose if you want to file the <i>Whole Message</i>, <i>Attachments Only</i>, or <i>Message Only</i>.</li>
              </ul>
            </HelpSection>

            <HelpSection title="Managing Your Locations" icon={<Add24Regular />}>
              <p>Use the toolbar buttons to manage your destination list:</p>
              <ul>
                <li><Add24Regular style={{fontSize: 14}}/> <b>Add</b>: Create a new filing destination.</li>
                <li><Edit24Regular style={{fontSize: 14}}/> <b>Edit</b>: Rename or change the path of a folder.</li>
                <li><Delete24Regular style={{fontSize: 14}}/> <b>Delete</b>: Remove a location from the table.</li>
                <li><FolderOpen24Regular style={{fontSize: 14}}/> <b>Explore</b>: Opens the destination folder on your computer in File Explorer.</li>
              </ul>
            </HelpSection>

            <HelpSection title="Smart Suggestions" icon={<Star24Regular />}>
              <p>Mail Manager learns which folders you use most often:</p>
              <ul>
                <li><Star24Regular style={{fontSize: 14, color: "#ffb900"}}/> <b>Stars/Rank</b>: Favorite or frequently used folders appear at the top.</li>
                <li><StarOff24Regular style={{fontSize: 14}}/> <b>Remove Suggestion</b>: Clears the "Suggested" status from a folder.</li>
                <li><EyeOff24Regular style={{fontSize: 14}}/> <b>Mark as Unused</b>: Hides a folder from suggestions without deleting it.</li>
              </ul>
            </HelpSection>

            <HelpSection title="Filing to Multiple Folders" icon={<SelectAllOn24Regular />}>
              <p>Click <b>"Select Multiple"</b> in the toolbar to enable checkboxes for multiple folders. This allows you to file the same email into several locations with one click.</p>
            </HelpSection>

          </DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="primary" onClick={() => onOpenChange(false)}>Got it</Button>
            </DialogTrigger>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default HelpDialog;
