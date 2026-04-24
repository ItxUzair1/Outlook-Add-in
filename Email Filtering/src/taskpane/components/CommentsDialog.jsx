import * as React from "react";
import { 
  Button, 
  Label, 
  Textarea,
  Field
} from "@fluentui/react-components";

/* global Office */

const CommentsDialog = ({ initialComment = "", onSave, onCancel }) => {
  const [comment, setComment] = React.useState(initialComment);

  const handleSave = () => {
    if (Office.context.ui && Office.context.ui.messageParent) {
      Office.context.ui.messageParent(`setComment:${comment}`);
    }
    if (onSave) onSave(comment);
  };

  const handleCancel = () => {
    if (Office.context.ui && Office.context.ui.messageParent) {
      Office.context.ui.messageParent("close");
    }
    if (onCancel) onCancel();
  };

  return (
    <div style={{ 
      padding: "20px", 
      display: "flex", 
      flexDirection: "column", 
      gap: "16px",
      height: "100vh",
      boxSizing: "border-box",
      backgroundColor: "#fff"
    }}>
      <h2 style={{ fontSize: "18px", fontWeight: "600", margin: "0" }}>Add Comment</h2>
      <p style={{ fontSize: "14px", color: "#605e5c", margin: "0" }}>
        Enter a comment to be associated with this email when it is filed.
      </p>
      
      <Field label="Comment">
        <Textarea 
          value={comment} 
          onChange={(e) => setComment(e.target.value)} 
          placeholder="Type your comment here..."
          resize="vertical"
          style={{ minHeight: "100px" }}
        />
      </Field>

      <div style={{ 
        marginTop: "auto", 
        display: "flex", 
        justifyContent: "flex-end", 
        gap: "8px" 
      }}>
        <Button appearance="secondary" onClick={handleCancel}>Cancel</Button>
        <Button appearance="primary" onClick={handleSave}>Save Comment</Button>
      </div>
    </div>
  );
};

export default CommentsDialog;
