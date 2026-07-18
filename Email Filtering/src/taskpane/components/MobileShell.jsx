/**
 * MobileShell.jsx
 *
 * Root container for the mobile Koyomail experience.
 *
 * Renders a fixed top header, a scrollable content area, and a sticky
 * bottom navigation bar. All features stay within this single taskpane —
 * there are no pop-up dialogs on mobile (Office doesn't support them).
 *
 * Tabs:
 *   file      → MobileFileScreen    (file the current email)
 *   search    → MobileSearchScreen  (search the archive)
 *   locations → MobileLocationsScreen
 *   options   → MobileOptionsScreen (agent URL / token)
 */

import * as React from "react";
import MobileFileScreen from "./MobileFileScreen";
import MobileSearchScreen from "./MobileSearchScreen";
import MobileLocationsScreen from "./MobileLocationsScreen";
import MobileOptionsScreen from "./MobileOptionsScreen";
import AgentStatusDot from "./AgentStatusDot";

const NAV_TABS = [
  { id: "file",      icon: "✉️",  label: "File"      },
  { id: "search",    icon: "🔍",  label: "Search"    },
  { id: "locations", icon: "📂",  label: "Locations" },
  { id: "options",   icon: "⚙️",  label: "Settings"  },
];

const styles = {
  shell: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif",
    background: "#f7f8fa",
    overflow: "hidden",
  },

  // ── Header ─────────────────────────────────────────────────────────────────
  header: {
    display: "flex",
    alignItems: "center",
    padding: "10px 16px",
    background: "#0078d4",
    color: "#fff",
    flexShrink: 0,
    boxShadow: "0 2px 6px rgba(0,0,0,.18)",
    zIndex: 100,
  },
  headerLogo: {
    width: 22,
    height: 22,
    marginRight: 8,
    borderRadius: 4,
  },
  headerTitle: {
    fontWeight: 700,
    fontSize: 15,
    letterSpacing: "0.02em",
    flex: 1,
  },

  // ── Content ────────────────────────────────────────────────────────────────
  content: {
    flex: 1,
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
    // bottom nav height = 60px; built-in bottom padding in each screen
  },

  // ── Bottom nav ─────────────────────────────────────────────────────────────
  bottomNav: {
    display: "flex",
    background: "#fff",
    borderTop: "1px solid #e0e0e0",
    flexShrink: 0,
    height: 60,
    zIndex: 100,
  },
  navItem: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    gap: 2,
    transition: "background .15s",
    userSelect: "none",
    WebkitTapHighlightColor: "transparent",
  },
  navItemActive: {
    background: "#e3f2fd",
  },
  navIcon: {
    fontSize: 18,
    lineHeight: 1,
  },
  navLabel: {
    fontSize: 10,
    fontWeight: 500,
    color: "#555",
  },
  navLabelActive: {
    color: "#0078d4",
    fontWeight: 700,
  },
};

export default function MobileShell({ initialMode }) {
  const defaultTab = initialMode === "mobile_search" ? "search" : "file";
  const [activeTab, setActiveTab] = React.useState(defaultTab);

  return (
    <div style={styles.shell}>
      {/* ── Header ── */}
      <div style={styles.header}>
        <span style={styles.headerTitle}>Koyomail</span>
        <AgentStatusDot />
      </div>

      {/* ── Screen content ── */}
      <div style={styles.content}>
        {activeTab === "file"      && <MobileFileScreen />}
        {activeTab === "search"    && <MobileSearchScreen />}
        {activeTab === "locations" && <MobileLocationsScreen />}
        {activeTab === "options"   && <MobileOptionsScreen />}
      </div>

      {/* ── Bottom navigation ── */}
      <nav style={styles.bottomNav}>
        {NAV_TABS.map((tab) => {
          const isActive = activeTab === tab.id;
          return (
            <div
              key={tab.id}
              style={{
                ...styles.navItem,
                ...(isActive ? styles.navItemActive : {}),
              }}
              onClick={() => setActiveTab(tab.id)}
            >
              <span style={styles.navIcon}>{tab.icon}</span>
              <span style={{
                ...styles.navLabel,
                ...(isActive ? styles.navLabelActive : {}),
              }}>
                {tab.label}
              </span>
            </div>
          );
        })}
      </nav>
    </div>
  );
}
