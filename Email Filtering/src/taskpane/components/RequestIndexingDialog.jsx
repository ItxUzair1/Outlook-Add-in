import * as React from "react";
import { Dismiss20Regular, Send20Regular, CheckmarkCircle20Regular, ErrorCircle20Regular } from "@fluentui/react-icons";
import { API_BASE_URL } from "../services/backendApi";

export default function RequestIndexingDialog({ onClose }) {
  const [projectName, setProjectName] = React.useState("");
  const [userEmail, setUserEmail] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [status, setStatus] = React.useState("idle"); // idle, success, error
  const [errorMsg, setErrorMsg] = React.useState("");

  React.useEffect(() => {
    // Try to auto-fetch the user email from Office JS
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
        const email = Office.context.mailbox.userProfile.emailAddress;
        if (email) {
          setUserEmail(email);
        }
      }
    } catch (e) {
      console.warn("Could not fetch user email automatically.", e);
    }
  }, []);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!projectName.trim() || !userEmail.trim()) {
      setErrorMsg("Please provide both a project name and your email address.");
      return;
    }

    setLoading(true);
    setStatus("idle");
    setErrorMsg("");

    try {
      const resp = await fetch(`${API_BASE_URL}/api/indexing-requests`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          projectName: projectName.trim(),
          userEmail: userEmail.trim(),
        })
      });

      if (!resp.ok) {
        const data = await resp.json().catch(() => ({}));
        throw new Error(data.error || "Failed to submit request.");
      }

      setStatus("success");
    } catch (err) {
      setStatus("error");
      setErrorMsg(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      position: "fixed", top: 0, left: 0, right: 0, bottom: 0,
      backgroundColor: "rgba(0,0,0,0.4)", display: "flex", justifyContent: "center",
      alignItems: "center", zIndex: 2000, borderRadius: 8
    }}>
      <div style={{
        backgroundColor: "#fff", borderRadius: 8, width: 440, maxWidth: "90%",
        boxShadow: "0 8px 32px rgba(0,0,0,0.2)", display: "flex", flexDirection: "column",
        overflow: "hidden"
      }}>
        {/* Header */}
        <div style={{
          padding: "16px 24px", borderBottom: "1px solid #edebe9",
          display: "flex", alignItems: "center", justifyContent: "space-between",
          backgroundColor: "#faf9f8"
        }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 600, color: "#323130" }}>Request Indexing</h2>
          <Dismiss20Regular style={{ cursor: "pointer", color: "#605e5c" }} onClick={onClose} />
        </div>

        {/* Body */}
        <div style={{ padding: "24px", flex: 1, overflowY: "auto" }}>
          {status === "success" ? (
            <div style={{ textAlign: "center", padding: "20px 0" }}>
              <CheckmarkCircle20Regular style={{ fontSize: 48, color: "#107c10", marginBottom: 16 }} />
              <h3 style={{ margin: "0 0 8px 0", color: "#323130" }}>Request Submitted!</h3>
              <p style={{ color: "#605e5c", fontSize: 14, margin: 0 }}>
                Your request for <strong>{projectName}</strong> has been sent to the admin. You will receive an email at <strong>{userEmail}</strong> once it's indexed.
              </p>
              <button 
                onClick={onClose}
                style={{
                  marginTop: 24, padding: "8px 24px", borderRadius: 4, border: "none",
                  backgroundColor: "#0078d4", color: "#fff", cursor: "pointer",
                  fontWeight: 600, fontFamily: "Segoe UI, sans-serif"
                }}
              >
                Close
              </button>
            </div>
          ) : (
            <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <p style={{ margin: "0 0 8px 0", color: "#605e5c", fontSize: 13, lineHeight: "1.5" }}>
                Can't find your project? Enter the project name or job number below, and our team will add its location to the search index.
              </p>

              {status === "error" && (
                <div style={{ backgroundColor: "#fde7e9", color: "#a4262c", padding: "12px", borderRadius: 4, display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                  <ErrorCircle20Regular style={{ flexShrink: 0 }} />
                  <span>{errorMsg}</span>
                </div>
              )}

              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                <label style={{ fontSize: 13, fontWeight: 600, color: "#323130" }}>
                  Project Name or Job Number <span style={{ color: "#a4262c" }}>*</span>
                </label>
                <input 
                  type="text" 
                  value={projectName}
                  onChange={(e) => setProjectName(e.target.value)}
                  placeholder="e.g. 12345 - City Center"
                  autoFocus
                  style={{
                    padding: "8px 12px", border: "1px solid #8a8886", borderRadius: 4,
                    fontSize: 14, fontFamily: "Segoe UI, sans-serif"
                  }}
                />
              </div>

              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                <label style={{ fontSize: 13, fontWeight: 600, color: "#323130" }}>
                  Your Email Address <span style={{ color: "#a4262c" }}>*</span>
                </label>
                <input 
                  type="email" 
                  value={userEmail}
                  onChange={(e) => setUserEmail(e.target.value)}
                  placeholder="name@company.com"
                  style={{
                    padding: "8px 12px", border: "1px solid #8a8886", borderRadius: 4,
                    fontSize: 14, fontFamily: "Segoe UI, sans-serif",
                    backgroundColor: "#f3f2f1" // Indicate it's auto-filled but editable
                  }}
                />
                <span style={{ fontSize: 11, color: "#605e5c" }}>We'll email you when the location is ready.</span>
              </div>

              <div style={{ display: "flex", gap: 12, justifyContent: "flex-end", marginTop: 16 }}>
                <button 
                  type="button"
                  onClick={onClose}
                  style={{
                    padding: "8px 20px", borderRadius: 4, border: "1px solid #8a8886",
                    backgroundColor: "#fff", color: "#323130", cursor: "pointer",
                    fontWeight: 600, fontFamily: "Segoe UI, sans-serif"
                  }}
                >
                  Cancel
                </button>
                <button 
                  type="submit"
                  disabled={loading || !projectName.trim() || !userEmail.trim()}
                  style={{
                    padding: "8px 20px", borderRadius: 4, border: "none",
                    backgroundColor: (loading || !projectName.trim() || !userEmail.trim()) ? "#c8c6c4" : "#0078d4", 
                    color: "#fff", cursor: (loading || !projectName.trim() || !userEmail.trim()) ? "default" : "pointer",
                    fontWeight: 600, display: "flex", alignItems: "center", gap: 6,
                    fontFamily: "Segoe UI, sans-serif"
                  }}
                >
                  {loading ? "Submitting..." : (
                    <>
                      <Send20Regular />
                      Submit Request
                    </>
                  )}
                </button>
              </div>
            </form>
          )}
        </div>
      </div>
    </div>
  );
}
