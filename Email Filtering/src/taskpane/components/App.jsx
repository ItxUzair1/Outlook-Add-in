import * as React from "react";
import PropTypes from "prop-types";
import { 
  addLocation, 
  deleteLocation, 
  fileEmail, 
  createDraftEmail,
  getLocations, 
  getSenderHistory,
  updateLocation,
  removeSuggestion,
  toggleSuggestion,
  markLocationUnused,
  getPreferences,
  updatePreferences,
  checkPathsConnectivity,
  exploreLocation,
  API_BASE_URL,
  remoteLog,
} from "../services/backendApi";
import { buildCurrentEmailPayload, addCategoryToCurrentEmail, ensureMasterCategory, toGraphItemId } from "../services/mailboxService";
import Toolbar from "./Toolbar";
import DetailsSidebar from "./DetailsSidebar";
import LocationTable from "./LocationTable";
import LocationDialog from "./LocationDialog";
import HelpDialog from "./HelpDialog";
import SearchDialog from "./SearchDialog";
import OptionsDialog from "./OptionsDialog";

import LocationsManagerDialog from "./LocationsManagerDialog";
import { Button, Spinner } from "@fluentui/react-components";
import { useMsal } from "@azure/msal-react";
import { getGraphToken, isOutlookIframeHost } from "../utils/authManager";
import {
  reportActionError,
  formatAfterFilingApiError,
  deleteItemViaEws,
  moveItemViaEws,
  recoverPostFilingAfterGraphFailure,
  isGraphPostFilingDeferralError,
} from "../utils/afterFilingUtils";

/* global Office */

function isBenignSsoError(message) {
  const lower = String(message || "").toLowerCase();
  return (
    lower.includes("timeout") ||
    lower.includes("sso token failed") ||
    lower.includes("not supported in this environment")
  );
}

function resolveSsoWarning(payload) {
  if (!payload || payload.isPartial) return "";
  if (isOutlookIframeHost()) {
    // New Outlook / filing dialogs use NAA or MSAL — Office SSO is optional here.
    return "";
  }
  if (payload.ssoTokenError && !isBenignSsoError(payload.ssoTokenError)) {
    return `⚠️ SSO Authentication Warning: ${payload.ssoTokenError}. The add-in will use MSAL fallback automatically when needed.`;
  }
  if (!payload.ssoToken && !payload.ssoTokenError) {
    return "⚠️ SSO token not available. The add-in will try MSAL fallback automatically for Graph operations.";
  }
  return "";
}

const sortLocationsList = (locationsArray, sender, senderStats, generalStats) => {
  const normalizePath = (p) => {
    if (!p) return "";
    return p.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
  };

  return [...locationsArray].sort((a, b) => {
    // 1. Keep unused folders at the bottom
    if (a.isUnused && !b.isUnused) return 1;
    if (!a.isUnused && b.isUnused) return -1;
    if (a.isUnused && b.isUnused) return 0;

    // 2. Prioritize starred/favourites for this sender (or general if sender is empty)
    if (a.isSuggested && !b.isSuggested) return -1;
    if (!a.isSuggested && b.isSuggested) return 1;
    if (a.isSuggested && b.isSuggested) {
      const normA = normalizePath(a.path);
      const normB = normalizePath(b.path);
      const statA = senderStats?.[normA];
      const statB = senderStats?.[normB];
      if (statA && !statB) return -1;
      if (!statA && statB) return 1;
      if (statA && statB) {
        if (statA.count !== statB.count) return statB.count - statA.count;
        if (statA.lastUsed !== statB.lastUsed) return statB.lastUsed - statA.lastUsed;
      }
      const genA = generalStats?.[normA];
      const genB = generalStats?.[normB];
      if (genA && !genB) return -1;
      if (!genA && genB) return 1;
      if (genA && genB) {
        if (genA.count !== genB.count) return genB.count - genA.count;
        if (genA.lastUsed !== genB.lastUsed) return genB.lastUsed - genA.lastUsed;
      }
      return String(a.description || "").localeCompare(String(b.description || ""));
    }

    // 3. Prioritize matching sender history details
    const normA = normalizePath(a.path);
    const normB = normalizePath(b.path);
    const statA = senderStats?.[normA];
    const statB = senderStats?.[normB];

    if (statA && !statB) return -1;
    if (!statA && statB) return 1;
    if (statA && statB) {
      if (statA.count !== statB.count) {
        return statB.count - statA.count;
      }
      return statB.lastUsed - statA.lastUsed;
    }

    // 4. Prioritize general history details
    const genA = generalStats?.[normA];
    const genB = generalStats?.[normB];

    if (genA && !genB) return -1;
    if (!genA && genB) return 1;
    if (genA && genB) {
      if (genA.count !== genB.count) {
        return genB.count - genA.count;
      }
      return genB.lastUsed - genA.lastUsed;
    }

    // 5. Sort by general lastUsedAt descending
    const timeA = a.lastUsedAt ? new Date(a.lastUsedAt).getTime() : 0;
    const timeB = b.lastUsedAt ? new Date(b.lastUsedAt).getTime() : 0;
    if (timeA !== timeB) {
      return timeB - timeA;
    }

    // 6. Fallback to alphabetical description
    return String(a.description || "").localeCompare(String(b.description || ""));
  });
};

const App = ({ title, initialMode: propInitialMode }) => {
  const initialMode = propInitialMode || (typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("mode") : null);
  const { instance } = useMsal();
  // Auth tier label shown in the status bar
  const [authTier, setAuthTier] = React.useState("");
  const autoAuthTriggeredRef = React.useRef(false);
  // emailPayloadRef always holds the latest emailPayload value.
  // Used by loadLocations so the callback stays stable (empty dep array).
  const emailPayloadRef = React.useRef(null);
  const [locations, setLocations] = React.useState([]);
  const [locationsLoading, setLocationsLoading] = React.useState(true);
  const locationsRef = React.useRef([]);
  const senderStatsRef = React.useRef({});
  const generalStatsRef = React.useRef({});

  React.useEffect(() => {
    locationsRef.current = locations;
  }, [locations]);
  const [selectedIds, setSelectedIds] = React.useState([]);
  const [narrowSidebarDismissed, setNarrowSidebarDismissed] = React.useState(false);
  const [isMultiSelect, setIsMultiSelect] = React.useState(false);
  const [multiEmailItems, setMultiEmailItems] = React.useState([]);
  const [connectivityStatus, setConnectivityStatus] = React.useState({});



  const getSavedDefault = (key, fallback) => {
    try {
      const optsStr = localStorage.getItem("koyomail_options");
      const opts = optsStr ? JSON.parse(optsStr) : {};
      
      if (key === "afterFiling" && opts.afterFilingAction !== undefined) {
        return opts.afterFilingAction === "move_deleted" ? "delete" : opts.afterFilingAction;
      }
      if (key === "attachmentsOption" && opts.defaultAttachments !== undefined) return opts.defaultAttachments;
      if (key === "markReviewed" && opts.markReviewed !== undefined) return opts.markReviewed;
      if (key === "sendLink" && opts.sendLink !== undefined) return opts.sendLink;
      
      const saved = localStorage.getItem(`koyomail_default_${key}`);
      return saved !== null ? JSON.parse(saved) : fallback;
    } catch {
      return fallback;
    }
  };

  const instantInfo = React.useMemo(() => {
    try {
      if (typeof Office !== "undefined" && Office.context?.mailbox?.item) {
        const item = Office.context.mailbox.item;
        const subjectVal = typeof item.subject === "string" ? item.subject : "";
        let senderVal = "";
        if (item.from) {
          if (typeof item.from === "object") {
            senderVal = item.from.emailAddress || item.from.displayName || "";
          } else if (typeof item.from === "string") {
            senderVal = item.from;
          }
        }
        if (!senderVal && item.sender) {
          if (typeof item.sender === "object") {
            senderVal = item.sender.emailAddress || item.sender.displayName || "";
          } else if (typeof item.sender === "string") {
            senderVal = item.sender;
          }
        }
        return {
          subject: subjectVal,
          sender: senderVal,
          itemId: item.itemId ? toGraphItemId(item.itemId) : "",
          isPartial: true
        };
      }
    } catch (err) {
      console.warn("[App] Failed to get instant email info:", err);
    }
    return null;
  }, []);

  // Filing Options State
  const [subject, setSubject] = React.useState(() => {
    if (instantInfo?.subject) return instantInfo.subject;
    const urlSubject = new URLSearchParams(window.location.search).get("subject");
    return urlSubject ? decodeURIComponent(urlSubject) : "";
  });
  const [comment, setComment] = React.useState("");
  const [afterFiling, setAfterFiling] = React.useState(() => getSavedDefault("afterFiling", "none"));
  const [markReviewed, setMarkReviewed] = React.useState(() => getSavedDefault("markReviewed", false));
  const [sendLink, setSendLink] = React.useState(() => getSavedDefault("sendLink", false));
  const [attachmentsOption, setAttachmentsOption] = React.useState(() => getSavedDefault("attachmentsOption", "all"));
  const [emailPayload, setEmailPayload] = React.useState(instantInfo);
  const [noItemSelected, setNoItemSelected] = React.useState(false);

  // Keep emailPayloadRef always in sync with the latest emailPayload state.
  // This ref is read inside loadLocations so the callback itself can have an
  // empty dependency array (staying stable), which prevents the mount useEffect
  // from re-running and creating the callId race that dropped collection results.
  React.useEffect(() => { emailPayloadRef.current = emailPayload; }, [emailPayload]);

  const [loading, setLoading] = React.useState(false);
  const [message, setMessage] = React.useState("");
  const [actionError, setActionError] = React.useState("");
  const [ssoWarning, setSsoWarning] = React.useState("");
  const [graphAuthStatus, setGraphAuthStatus] = React.useState("Checking authentication...");
  const [graphAuthOk, setGraphAuthOk] = React.useState(false);
  const [isFiled, setIsFiled] = React.useState(false);
  const [brokenCollectionNames, setBrokenCollectionNames] = React.useState([]);
  const abortControllerRef = React.useRef(null);

  const [koyoOptions, setKoyoOptions] = React.useState(() => {
    try {
      const opts = localStorage.getItem("koyomail_options");
      return opts ? JSON.parse(opts) : {};
    } catch {
      return {};
    }
  });

  const [width, setWidth] = React.useState(() => typeof window !== "undefined" ? window.innerWidth : 850);

  React.useEffect(() => {
    const handleResize = () => {
      setWidth(window.innerWidth);
    };
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  const isNarrow = width < 500;

  const [showNarrowAuthSuccess, setShowNarrowAuthSuccess] = React.useState(true);

  React.useEffect(() => {
    if (graphAuthOk && isNarrow) {
      setShowNarrowAuthSuccess(true);
      const timer = setTimeout(() => {
        setShowNarrowAuthSuccess(false);
      }, 6000); // 6 seconds
      return () => clearTimeout(timer);
    } else {
      setShowNarrowAuthSuccess(true);
    }
  }, [graphAuthOk, isNarrow]);

  React.useEffect(() => {
    const loadOptions = () => {
      try {
        const opts = localStorage.getItem("koyomail_options");
        const parsed = opts ? JSON.parse(opts) : {};
        setKoyoOptions(parsed);
        if (parsed.addFiledCategory !== false) {
          const categoryName = parsed.filedCategoryName || "Filed by Koyomail";
          ensureMasterCategory(categoryName, "Preset3").catch((err) => {
            console.warn("[App] Failed to ensure master category:", err.message);
          });
        }
      } catch {
        setKoyoOptions({});
      }
    };
    
    const handleStorageChange = (e) => {
      if (e.key === "koyomail_options") {
        loadOptions();
      }
    };

    window.addEventListener("koyomail_options_updated", loadOptions);
    window.addEventListener("storage", handleStorageChange);
    

    // Listen for subject passed securely from the parent
    if (typeof Office !== "undefined" && Office.context?.ui?.addHandlerAsync) {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
        try {
          const data = JSON.parse(arg.message);
          if (data.type === "subject" && data.value) {
            setSubject(data.value);
          }
        } catch (e) {
          console.warn("[App] Failed to parse parent message", e);
        }
      });
    }

    return () => {
      window.removeEventListener("koyomail_options_updated", loadOptions);
      window.removeEventListener("storage", handleStorageChange);
    };
  }, []);

  React.useEffect(() => {
    const syncPreferences = async () => {
      try {
        const backendParsed = await getPreferences();
        const stored = localStorage.getItem("koyomail_options");
        const localParsed = stored ? JSON.parse(stored) : {};
        const parsed = { ...backendParsed, ...localParsed };
        localStorage.setItem("koyomail_options", JSON.stringify(parsed));
        setKoyoOptions(parsed);
        window.dispatchEvent(new Event("koyomail_options_updated"));

        // Sync loaded collections: merge backend + local so neither wipes the other.
        // Backend is the durable source; local may have additions made since last save.
        const backendCollections = backendParsed.loadedCollections && Array.isArray(backendParsed.loadedCollections)
          ? backendParsed.loadedCollections
          : [];
        const localCollectionsRaw = localStorage.getItem("koyomail_loaded_collections");
        const localCollections = localCollectionsRaw ? (JSON.parse(localCollectionsRaw) || []) : [];

        // Union: all unique paths from both sources (backend is authoritative, local adds extras)
        const mergedCollections = Array.from(new Set([...backendCollections, ...localCollections]));

        if (mergedCollections.length > 0) {
          const mergedJson = JSON.stringify(mergedCollections);
          localStorage.setItem("koyomail_loaded_collections", mergedJson);

          // If the merged list differs from what the backend has, push the update back
          if (mergedCollections.length !== backendCollections.length) {
            updatePreferences({ loadedCollections: mergedCollections }).catch(() => {});
          }

          // Dispatch storage event to trigger reload of locations list
          window.dispatchEvent(new StorageEvent("storage", {
            key: "koyomail_loaded_collections",
            newValue: mergedJson
          }));

          // Probe each saved collection path — if any are unreachable on this machine, warn the user
          const broken = [];
          for (const filePath of mergedCollections) {
            try {
              const probeResp = await fetch(`${API_BASE_URL}/api/collections/load`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ filePath })
              });
              if (!probeResp.ok) {
                const filename = filePath.split('\\').pop().split('/').pop();
                broken.push(filename.replace(/\.mmcollection$/i, ''));
              }
            } catch (_) {
              const filename = filePath.split('\\').pop().split('/').pop();
              broken.push(filename.replace(/\.mmcollection$/i, ''));
            }
          }
          if (broken.length > 0) {
            setBrokenCollectionNames(broken);
          }
        }
      } catch (err) {
        console.warn("[App] Failed to sync backend preferences on mount:", err.message);
      }
    };
    syncPreferences();
  }, []);

  React.useEffect(() => {
    if (!emailPayload?.itemId) return;

    const itemId = emailPayload.itemId;
    
    const syncComment = () => {
      const stored = localStorage.getItem(`koyomail_comment_${itemId}`);
      if (stored !== null) {
        setComment(stored);
      } else {
        // Fallback for any older comments saved globally
        const temp = localStorage.getItem("koyomail_temp_comment");
        if (temp !== null) {
          setComment(temp);
        } else {
          setComment("");
        }
      }
    };

    // Run once on mount or when email changes
    syncComment();

    window.addEventListener("storage", syncComment);
    window.addEventListener("koyomail_comment_updated", syncComment);

    return () => {
      window.removeEventListener("storage", syncComment);
      window.removeEventListener("koyomail_comment_updated", syncComment);
    };
  }, [emailPayload?.itemId]);

  React.useEffect(() => {
    const handleStorageChange = (e) => {
      if (e.key === "koyomail_loaded_collections") {
        // Always reload when the collection list changes or is restored from backend—
        // do NOT gate on sender here, otherwise a post-update restore is silently ignored
        // in filing mode before the email payload has fully resolved.
        loadLocations(null, { silent: true });
      } else if (e.key === "koyomail_locations_updated") {
        const isReadFilingMode = initialMode === "file" || !initialMode;
        if (!isReadFilingMode || emailPayloadRef.current?.sender) {
          loadLocations(null, { silent: true });
        }
      }
    };
    const handleFocus = () => {
      const isReadFilingMode = initialMode === "file" || !initialMode;
      if (!isReadFilingMode || emailPayloadRef.current?.sender) {
        loadLocations(null, { silent: true });
      }
    };
    window.addEventListener("storage", handleStorageChange);
    window.addEventListener("focus", handleFocus);
    return () => {
      window.removeEventListener("storage", handleStorageChange);
      window.removeEventListener("focus", handleFocus);
    };
  }, [loadLocations, initialMode]);

  const saveDefaults = React.useCallback(() => {
    try {
      const optsStr = localStorage.getItem("koyomail_options");
      const opts = optsStr ? JSON.parse(optsStr) : {};
      
      opts.afterFilingAction = afterFiling;
      opts.markReviewed = markReviewed;
      opts.sendLink = sendLink;
      opts.defaultAttachments = attachmentsOption;
      
      localStorage.setItem("koyomail_options", JSON.stringify(opts));

      setMessage("Default options saved.");
      setTimeout(() => setMessage(""), 3000);
    } catch (e) {
      console.warn("Failed to save defaults:", e);
    }
  }, [afterFiling, markReviewed, sendLink, attachmentsOption]);

  /**
   * Unified token getter — delegates to authManager which runs the three-tier chain:
   *   Tier 1: Office SSO  → Tier 2: NAA (New Outlook)  → Tier 3: MSAL redirect (Classic)
   */
  const getToken = React.useCallback(async ({ interactive = false } = {}) => {
    const result = await getGraphToken({
      msalInstance: instance,
      interactive,
      loginHint: Office?.context?.mailbox?.userProfile?.emailAddress,
    });
    setAuthTier(result.tier);
    return result.token;
  }, [instance]);

  const openComposeWindow = React.useCallback((links, emailSubject) => {
    if (!links || links.length === 0) return;
    
    try {
      const mailbox = Office?.context?.mailbox;
      if (!mailbox || !mailbox.displayNewMessageForm) {
        // Fallback: show links in a message instead
        console.warn("[App] displayNewMessageForm not available. Showing links as message.");
        setMessage(`Filed link(s): ${links.join(", ")}`);
        return;
      }

      // Formatting for high-fidelity Outlook display
      const fontFam = koyoOptions.emailFont || 'Segoe UI';
      const fontSz = koyoOptions.fontSize ? `${koyoOptions.fontSize}pt` : '11pt';
      
      const formattedLinks = links.map(l => `- ${l}`).join("<br/>");
      const htmlBody = `
        <div style="font-family: '${fontFam}', sans-serif; font-size: ${fontSz}; color: #323130;">
          <p>The following email has been filed to a shared location:</p>
          <p><strong>${formattedLinks}</strong></p>
          <p><i>Generated by Koyomail</i></p>
        </div>
      `;
      
      mailbox.displayNewMessageForm({
        toRecipients: [],
        subject: `Filed Link: ${emailSubject}`,
        htmlBody: htmlBody
      });
    } catch (err) {
      console.warn("[App] Failed to open compose window:", err.message);
      setMessage(`Email filed. Link: ${links.join(", ")}`);
    }
  }, [koyoOptions]);
  
  // Poll for errors from the parent context (commands.js)
  React.useEffect(() => {
    if (afterFiling === "none" || !loading) return;

    const interval = setInterval(() => {
      const stored = localStorage.getItem("koyomailActionError");
      if (stored) {
        try {
          const { message: errMsgs, timestamp } = JSON.parse(stored);
          if (Date.now() - timestamp < 30000) {
            const safeError = typeof errMsgs === "string" ? errMsgs : JSON.stringify(errMsgs);
            setActionError(safeError);
            localStorage.removeItem("koyomailActionError");
          }
        } catch (e) { /* ignore */ }
      }
    }, 1000);

    // Global error interceptors to expose raw [object Object] payload details
    window.onerror = function(message, source, lineno, colno, error) {
      try {
        const detail = error ? JSON.stringify(error, Object.getOwnPropertyNames(error)) : String(message);
        setActionError(`Global Error: ${detail}`);
      } catch (e) {
        setActionError(`Global Error: ${String(message)}`);
      }
      return true; // Prevent default error overlay
    };
    window.onunhandledrejection = function(event) {
      if (event && event.preventDefault) {
        event.preventDefault();
      }
      try {
        const detail = event.reason ? JSON.stringify(event.reason, Object.getOwnPropertyNames(event.reason)) : "Unknown rejection";
        setActionError(`Unhandled Promise: ${detail}`);
      } catch (e) {
        setActionError(`Unhandled Promise: ${String(event.reason)}`);
      }
    };

    return () => {
      clearInterval(interval);
      window.onerror = null;
      window.onunhandledrejection = null;
    };
  }, [afterFiling, loading]);
  
  // Dialog State
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [isHelpOpen, setIsHelpOpen] = React.useState(initialMode === "help");
  const [isSearchOpen, setIsSearchOpen] = React.useState(initialMode === "search");
  const [isOptionsOpen, setIsOptionsOpen] = React.useState(initialMode === "options");

  const [optionsInitialTab, setOptionsInitialTab] = React.useState("Local & Network folders");
  const [editingLocation, setEditingLocation] = React.useState(null);

  // Helper to centralize sorting logic
  const sortLocationsList = (rows, sender, senderStats, generalStats) => {
    const normalizePath = (p) => (p || "").replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
    
    return [...rows].sort((a, b) => {
      if (a.isUnused && !b.isUnused) return 1;
      if (!a.isUnused && b.isUnused) return -1;
      if (a.isUnused && b.isUnused) return 0;

      if (a.isSuggested && !b.isSuggested) return -1;
      if (!a.isSuggested && b.isSuggested) return 1;
      if (a.isSuggested && b.isSuggested) {
        const normA = normalizePath(a.path);
        const normB = normalizePath(b.path);
        const statA = senderStats[normA];
        const statB = senderStats[normB];
        if (statA && !statB) return -1;
        if (!statA && statB) return 1;
        if (statA && statB) {
          if (statA.count !== statB.count) return statB.count - statA.count;
          if (statA.lastUsed !== statB.lastUsed) return statB.lastUsed - statA.lastUsed;
        }
        const genA = generalStats[normA];
        const genB = generalStats[normB];
        if (genA && !genB) return -1;
        if (!genA && genB) return 1;
        if (genA && genB) {
          if (genA.count !== genB.count) return genB.count - genA.count;
          if (genA.lastUsed !== genB.lastUsed) return genB.lastUsed - genA.lastUsed;
        }
        return String(a.description || "").localeCompare(String(b.description || ""));
      }

      const normA = normalizePath(a.path);
      const normB = normalizePath(b.path);
      const statA = senderStats[normA];
      const statB = senderStats[normB];
      if (statA && !statB) return -1;
      if (!statA && statB) return 1;
      if (statA && statB) {
        if (statA.count !== statB.count) return statB.count - statA.count;
        return statB.lastUsed - statA.lastUsed;
      }
      const genA = generalStats[normA];
      const genB = generalStats[normB];
      if (genA && !genB) return -1;
      if (!genA && genB) return 1;
      if (genA && genB) {
        if (genA.count !== genB.count) return genB.count - genA.count;
        return genB.lastUsed - genA.lastUsed;
      }
      const timeA = a.lastUsedAt ? new Date(a.lastUsedAt).getTime() : 0;
      const timeB = b.lastUsedAt ? new Date(b.lastUsedAt).getTime() : 0;
      if (timeA !== timeB) return timeB - timeA;
      return String(a.description || "").localeCompare(String(b.description || ""));
    });
  };

  // Ref used to cancel stale concurrent loadLocations calls.
  // When two calls fire simultaneously (initial mount + sender change) the first
  // one is abandoned so only the latest result is applied to state.
  const loadLocationsIdRef = React.useRef(0);

  const loadLocations = React.useCallback(async (senderParam, options = {}) => {
    const callId = ++loadLocationsIdRef.current;

    const silent = !!options.silent;
    const lightweight = !!options.lightweight;
    if (!silent && (!locationsRef.current || locationsRef.current.length === 0)) {
      setLocationsLoading(true);
    }

    try {
      // Read sender from the ref so this callback never needs emailPayload as a
      // dependency — keeping the function reference stable across renders.
      const sender = senderParam || emailPayloadRef.current?.sender;
      
      const [localRows, senderHistoryData] = await Promise.all([
        getLocations().catch((err) => {
          console.warn("[App] Failed to load local locations from backend:", err);
          return [];
        }),
        getSenderHistory(sender || "").catch((err) => {
          console.warn("[App] Failed to fetch sender history:", err);
          return { history: {}, favourites: [], generalHistory: {} };
        })
      ]);

      // Abort if a newer call was started while we were awaiting responses
      if (callId !== loadLocationsIdRef.current) return;

      let rows = [...localRows];

      // Sync locations from loaded Collections (skipped after filing for speed)
      if (!lightweight) {
      try {
        let loadedCollectionsRaw = localStorage.getItem("koyomail_loaded_collections");

        // ── Recovery path: if localStorage was cleared (e.g. after a Koyomail update),
        // the collection file list is gone. Fetch it from the backend preferences store
        // which is persisted in the data directory and survives browser storage wipes.
        if (!loadedCollectionsRaw || loadedCollectionsRaw === "[]" || loadedCollectionsRaw === "null") {
          try {
            remoteLog("info", "[App] koyomail_loaded_collections missing from localStorage — attempting backend preferences restore");
            const prefResp = await fetch(`${API_BASE_URL}/api/preferences?_t=${Date.now()}`);
            if (prefResp.ok) {
              const prefs = await prefResp.json();
              if (prefs.loadedCollections && Array.isArray(prefs.loadedCollections) && prefs.loadedCollections.length > 0) {
                loadedCollectionsRaw = JSON.stringify(prefs.loadedCollections);
                localStorage.setItem("koyomail_loaded_collections", loadedCollectionsRaw);
                remoteLog("info", `[App] Restored ${prefs.loadedCollections.length} collection path(s) from backend preferences`);
              }
            }
          } catch (restoreErr) {
            remoteLog("warn", `[App] Backend preferences restore failed: ${restoreErr.message}`);
          }
        }

        remoteLog("info", `[App] Sync collections raw: ${loadedCollectionsRaw}`);
        if (loadedCollectionsRaw) {
          const filePaths = JSON.parse(loadedCollectionsRaw);
          if (Array.isArray(filePaths)) {
            const baseUrl = API_BASE_URL;
            remoteLog("info", `[App] Loading collections: ${JSON.stringify(filePaths)}`);

            // Load all collections in parallel using Promise.all (safe since each map promise handles errors internally)
            const collectionResults = await Promise.all(
              filePaths.map(async (filePath) => {
                try {
                  const loadResp = await fetch(`${baseUrl}/api/collections/load`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ filePath })
                  });
                  if (!loadResp.ok) {
                    remoteLog("warn", `[App] Failed to load collection ${filePath}: status ${loadResp.status}`);
                    return null;
                  }
                  const data = await loadResp.json();
                  const rawCollectionName = filePath.split('\\').pop().split('/').pop().replace(/\.mmcollection$/i, '');
                  // "Personal" and "Private" are the same — normalise Personal → Private
                  const collectionName = rawCollectionName.toLowerCase() === "personal" ? "Private" : rawCollectionName;
                  if (!data.locations || !Array.isArray(data.locations)) {
                    remoteLog("warn", `[App] No locations found or invalid array in collection ${filePath}`);
                    return null;
                  }
                  const validLocations = data.locations.filter(Boolean);
                  remoteLog("info", `[App] Loaded collection "${rawCollectionName}" (→ "${collectionName}") successfully with ${validLocations.length} locations`);
                  return validLocations.map((loc, idx) => {
                    const originalId = loc.id || idx;
                    // Also normalise any per-location collection field that may say "Personal"
                    const locCollection = loc.collection && loc.collection.toLowerCase() === "personal"
                      ? "Private"
                      : (loc.collection || collectionName);
                    return {
                      ...loc,
                      id: `col_${rawCollectionName}_${originalId}`,
                      path: loc.folder || loc.path,
                      collection: locCollection === rawCollectionName ? collectionName : locCollection
                    };
                  });
                } catch (fetchErr) {
                  remoteLog("error", `[App] Network error while fetching collection ${filePath}: ${fetchErr.message}`);
                  return null;
                }
              })
            );

            for (const value of collectionResults) {
              if (value && Array.isArray(value)) {
                // Collection file locations must REPLACE any "Discovered" DB entry for
                // the same path — otherwise the auto-discovered entry shadows the real
                // project name/number from the .mmcollection file.
                const collectionPaths = new Set(
                  value.filter(cl => cl && cl.path).map(cl => String(cl.path).toLowerCase())
                );

                // Strip out Discovered DB entries that are now covered by a collection file
                rows = rows.filter(r => {
                  if (!r || !r.path) return true;
                  const isDiscovered = String(r.collection || "").toLowerCase() === "discovered";
                  return !(isDiscovered && collectionPaths.has(String(r.path).toLowerCase()));
                });

                // Now add collection entries that are not already present (non-Discovered
                // DB entries like Private/Portfolio still take priority)
                const existingPaths = new Set(rows.map(r => r && r.path ? String(r.path).toLowerCase() : ""));
                const unique = value.filter(cl => cl && cl.path && !existingPaths.has(String(cl.path).toLowerCase()));
                remoteLog("info", `[App] Collection unique locations count: ${unique.length} out of ${value.length}`);
                rows = [...rows, ...unique];
              }
            }
          }
        }
      } catch (err) {
        remoteLog("error", `[App] Failed to load collection locations into main list: ${err.message}`);
      }
      }

      // Abort again if a newer call overtook us during collection fetching
      if (callId !== loadLocationsIdRef.current) return;

      const normalizePath = (p) => {
        if (!p) return "";
        return p.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase().trim();
      };

      const senderStats = senderHistoryData?.history || {};
      const senderFavourites = senderHistoryData?.favourites || [];
      const generalStats = senderHistoryData?.generalHistory || {};

      senderStatsRef.current = senderStats;
      generalStatsRef.current = generalStats;

      const normalizedFavourites = senderFavourites.map(p => normalizePath(p));

      // Map dynamic isSuggested and isSenderSuggested flags
      const mappedRows = rows.map((loc) => {
        const normPath = normalizePath(loc.path);
        const hasHistory = !!(senderStats && senderStats[normPath]);
        const isFav = sender
          ? normalizedFavourites.includes(normPath)
          : !!loc.isSuggested;
        return {
          ...loc,
          originalSuggested: !!loc.isSuggested,
          isSuggested: isFav,
          isSenderSuggested: hasHistory
        };
      });

      // Sort the combined array using the helper function
      const sortedRows = sortLocationsList(mappedRows, sender, senderStats, generalStats);

      setLocations(sortedRows);
      setLocationsLoading(false);

      // Guard localStorage write — skip if the payload would be too large (>2MB)
      try {
        const serialized = JSON.stringify(sortedRows);
        if (serialized.length < 2 * 1024 * 1024) {
          localStorage.setItem("koyomail_locations", serialized);
        }
      } catch (e) {
        console.warn("Could not cache locations in localStorage:", e);
      }

      // ── Connectivity check: fire-and-forget with a per-path timeout ───────
      // Running fs.access() on unreachable UNC paths can hang for 30+ seconds
      // each. We must NOT block the UI waiting for this. Instead we fire it in
      // the background and update state in batches to allow progressive checkmark loading.
      const checkConnectivityInBatches = async (allLocations) => {
        const BATCH_SIZE = 15;
        let currentStatus = {};

        for (let i = 0; i < allLocations.length; i += BATCH_SIZE) {
          // Abort if a newer call has started in the meantime
          if (callId !== loadLocationsIdRef.current) return;

          const batch = allLocations.slice(i, i + BATCH_SIZE);
          try {
            const batchStatus = await checkPathsConnectivity(batch);
            if (callId === loadLocationsIdRef.current) {
              currentStatus = { ...currentStatus, ...batchStatus };
              setConnectivityStatus({ ...currentStatus });
            }
          } catch (err) {
            console.warn("[App] Batch connectivity check failed:", err.message);
          }

          // Small delay between batches to free up the event loop/network connection pool
          await new Promise(resolve => setTimeout(resolve, 80));
        }
      };

      checkConnectivityInBatches(sortedRows);
      // ─────────────────────────────────────────────────────────────────────

    } catch (error) {
      if (callId !== loadLocationsIdRef.current) return; // stale call, ignore
      setLocationsLoading(false);
      console.error("[App] Load failed:", error);
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Load failed: ${errorMsg}`);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // intentionally empty — see emailPayloadRef above

  React.useEffect(() => {
    if (emailPayload?.sender) {
      loadLocations(emailPayload.sender);
    }
  }, [emailPayload?.sender, loadLocations]);

  const handleRefresh = React.useCallback(() => {
    loadLocations();
    
    const isReadFilingMode = initialMode === "file" || !initialMode;
    if (initialMode === "file_multi" || (isReadFilingMode && !Office.context?.mailbox?.item)) {
      try {
        if (Office.context?.mailbox?.getSelectedItemsAsync) {
          Office.context.mailbox.getSelectedItemsAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const items = result.value || [];
              setMultiEmailItems(items);
              setSubject(items.length === 1 ? items[0].subject : `Multiple Emails (${items.length})`);
              if (items.length > 0) {
                setNoItemSelected(false);
                setMessage("");
              } else {
                setNoItemSelected(true);
                setMessage("Please select an email to view or file.");
              }
            }
          });
        }
      } catch (e) {
        console.warn("[App] Failed to refresh selection:", e);
      }
    }
  }, [loadLocations, initialMode]);

  React.useEffect(() => {
    // Reload taskpane if the item changes while it's pinned
    if (typeof Office !== "undefined" && Office.context?.mailbox?.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
        window.location.reload();
      });
    }

    // Only load immediately if we are not in read filing mode (where we must wait for the sender to be resolved)
    const isReadFilingMode = initialMode === "file" || !initialMode;
    if (!isReadFilingMode) {
      loadLocations();
    }

    if (initialMode === "file_multi") {
      try {
        if (Office.context?.mailbox?.getSelectedItemsAsync) {
          Office.context.mailbox.getSelectedItemsAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const items = result.value || [];
              setMultiEmailItems(items);
              setSubject(`Multiple Emails (${items.length})`);
              if (items.length > 0) {
                setNoItemSelected(false);
                setMessage("");
              } else {
                setNoItemSelected(true);
                setMessage("Please select an email to view or file.");
              }
            } else {
              setNoItemSelected(true);
              setMessage("Please select an email to view or file.");
            }
          });
        } else {
          setNoItemSelected(true);
          setMessage("getSelectedItemsAsync is not supported in this client.");
        }
      } catch (e) {
        setNoItemSelected(true);
        setMessage("Error fetching selected items.");
      }
      return;
    }
    if (initialMode === "help" || initialMode === "search" || initialMode === "options" || initialMode === "onsend" || initialMode === "collections" || initialMode === "locations") {
      return;
    }

    const fetchData = async () => {
      try {
        const payload = await buildCurrentEmailPayload();
        if (payload) {
          setEmailPayload(payload);
          setSubject(payload.subject || "");

          if (!payload.sender) {
            console.log("[App] fetchData resolved but no sender found. Loading unsorted locations.");
            loadLocations();
          }

          // Do not show SSO warnings until full payload is available.
          if (!payload.isPartial) {
            setSsoWarning(resolveSsoWarning(payload));
          }
          
          // If the payload is partial, poll for the full enrichment from background
          if (payload.isPartial) {
            console.log("[App] Partial data found, polling for full enrichment...");
            const pollInterval = setInterval(async () => {
              try {
                const enriched = await buildCurrentEmailPayload();
                if (enriched && !enriched.isPartial) {
                  console.log("[App] Full enrichment received (Body & Attachments).");
                  setEmailPayload(enriched);

                  setSsoWarning(resolveSsoWarning(enriched));

                  clearInterval(pollInterval);
                }
              } catch (pollErr) {
                console.warn("[App] Polling enrichment failed:", pollErr.message);
              }
            }, 1000);
            
            // Stop polling after 15 seconds to prevent memory leak
            setTimeout(() => clearInterval(pollInterval), 15000);
          }
        }
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : (typeof err === "object" ? JSON.stringify(err) : String(err));
        console.warn("[App] Initial data gathering failed:", errorMsg);
        if (errorMsg.includes("No mailbox item is currently selected")) {
          // Fallback: check if there are selected items in the list view (reading pane off)
          if (Office.context?.mailbox?.getSelectedItemsAsync) {
            Office.context.mailbox.getSelectedItemsAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
                const items = result.value;
                setMultiEmailItems(items);
                setSubject(items.length === 1 ? items[0].subject : `Multiple Emails (${items.length})`);
                setNoItemSelected(false);
                setMessage("");
              } else {
                setNoItemSelected(true);
                setMessage("Please select an email to view or file.");
              }
            });
          } else {
            setNoItemSelected(true);
            setMessage("Please select an email to view or file.");
          }
        } else {
          setMessage(`Initial load failed: ${errorMsg}`);
        }
        loadLocations();
      }
    };

    fetchData();
  }, [loadLocations, initialMode]);

  React.useEffect(() => {
    if (instantInfo?.sender) return;
    const isReadFilingMode = initialMode === "file" || !initialMode;
    if (!isReadFilingMode) return;

    let attempts = 0;
    const pollInterval = setInterval(() => {
      attempts++;
      try {
        if (typeof Office !== "undefined" && Office.context?.mailbox?.item) {
          const item = Office.context.mailbox.item;
          const subjectVal = typeof item.subject === "string" ? item.subject : "";
          
          if (subjectVal && !subject) {
            setSubject(subjectVal);
          }

          let senderVal = "";
          if (item.from) {
            if (typeof item.from === "object") {
              senderVal = item.from.emailAddress || item.from.displayName || "";
            } else if (typeof item.from === "string") {
              senderVal = item.from;
            }
          }
          if (!senderVal && item.sender) {
            if (typeof item.sender === "object") {
              senderVal = item.sender.emailAddress || item.sender.displayName || "";
            } else if (typeof item.sender === "string") {
              senderVal = item.sender;
            }
          }

          if (senderVal) {
            clearInterval(pollInterval);
            setEmailPayload({
              subject: subjectVal,
              sender: senderVal,
              itemId: item.itemId ? toGraphItemId(item.itemId) : "",
              isPartial: true
            });
            if (subjectVal) {
              setSubject(subjectVal);
            }
          }
        }
      } catch (err) {
        console.warn("[App] Error in mailbox item poll:", err);
      }

      if (attempts >= 30) {
        clearInterval(pollInterval);
        console.log("[App] Mailbox item poll timed out; waiting for fetchData fallback.");
      }
    }, 50);

    return () => clearInterval(pollInterval);
  }, [instantInfo, initialMode, loadLocations, subject]);

  // ── Auto-authentication on load ─────────────────────────────────────────────
  React.useEffect(() => {
    if (graphAuthOk) {
      setSsoWarning("");
    }
  }, [graphAuthOk]);

  React.useEffect(() => {
    if (autoAuthTriggeredRef.current) return;
    autoAuthTriggeredRef.current = true;

    const autoAuthenticate = async () => {
      const AUTH_STARTUP_TIMEOUT_MS = 20000;
      try {
        setGraphAuthStatus("Authenticating...");
        const token = await Promise.race([
          getToken({ interactive: false }),
          new Promise((_, reject) =>
            setTimeout(
              () => reject(new Error("Authentication timed out. Click Sign In to continue.")),
              AUTH_STARTUP_TIMEOUT_MS
            )
          ),
        ]);
        if (token) {
          setGraphAuthOk(true);
          setGraphAuthStatus("Signed in ✓");
          setSsoWarning("");
          return;
        }
      } catch (authErr) {
        console.warn("[App] Silent authentication failed:", authErr?.message || authErr);

        const inIframe = isOutlookIframeHost();
        const wasPreviouslySignedIn = !!localStorage.getItem("koyomail_activeAccountId");

        // New Outlook / filing dialog: try interactive NAA or auth dialog immediately.
        if (inIframe) {
          try {
            setGraphAuthStatus("Signing in...");
            const token = await getToken({ interactive: true });
            if (token) {
              setGraphAuthOk(true);
              setGraphAuthStatus("Signed in ✓");
              setSsoWarning("");
              return;
            }
          } catch (iframeAuthErr) {
            console.warn("[App] Interactive auth failed in iframe host:", iframeAuthErr?.message || iframeAuthErr);
          }
        } else if (wasPreviouslySignedIn) {
          // Classic desktop: reconnect prior MSAL session in-window.
          try {
            setGraphAuthStatus("Reconnecting session...");
            const token = await getToken({ interactive: true });
            if (token) {
              setGraphAuthOk(true);
              setGraphAuthStatus("Signed in ✓");
              setSsoWarning("");
              return;
            }
          } catch (reconnectErr) {
            console.warn("[App] Interactive reconnect failed:", reconnectErr?.message || reconnectErr);
          }
        }
      }

      setGraphAuthOk(false);
      setGraphAuthStatus("Sign in required");
    };

    autoAuthenticate();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const onSelectionChange = (id) => {
    setNarrowSidebarDismissed(false);
    setSelectedIds((prev) => {
      if (isMultiSelect) {
        return prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id];
      } else {
        return prev.includes(id) ? [] : [id];
      }
    });
  };

  const onSaveLocation = async (data) => {
    try {
      if (editingLocation) {
        await updateLocation(editingLocation.id, data);
        setMessage("Location updated.");
      } else {
        await addLocation(data);
        setMessage("Location added.");
      }
      await loadLocations(null, { silent: true });
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Save failed: ${errorMsg}`);
    }
  };
  const onDeleteLocation = async () => {
    if (selectedIds.length === 0) {
      setMessage("Please select at least one location to delete.");
      return;
    }

    try {
      setLoading(true);
      setMessage(`Deleting ${selectedIds.length} location(s)...`);
      for (const id of selectedIds) {
        await deleteLocation(id);
      }
      setSelectedIds([]);
      setMessage("Location(s) deleted successfully.");
      await loadLocations(null, { silent: true });
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      console.error("Delete failed:", error);
      setMessage(`Delete failed: ${errorMsg}`);
    } finally {
      setLoading(false);
    }
  };

  const onFileEmail = async () => {
    if (selectedIds.length === 0) return;

    if (initialMode === "onsend") {
      const paths = selectedIds.map(id => locations.find(loc => loc.id === id)?.folder || locations.find(loc => loc.id === id)?.path);

      // Acquire SSO token silently so the backend can apply the category to the
      // Sent Items copy via Microsoft Graph (the compose item is frozen during ItemSend).
      let ssoToken = null;
      try {
        ssoToken = await getToken({ interactive: false });
      } catch (tokenErr) {
        console.warn("[App] Could not acquire SSO token for On-Send category tagging:", tokenErr.message);
      }

      // Ensure the master category exists with color Preset3 (Yellow) on the client side before notifying parent
      if (koyoOptions.addFiledCategory !== false) {
        const categoryName = koyoOptions.filedCategoryName || "Filed by Koyomail";
        try {
          await ensureMasterCategory(categoryName);
        } catch (catErr) {
          console.warn("[App] Failed to ensure master category:", catErr.message);
        }
      }

      const payloadData = {
        paths,
        subject,
        comment,
        attachmentsOption,
        markReviewed,
        sendLink,
        isOnSend: true,
        ssoToken: ssoToken || null,
        afterFiling: afterFiling || "none",
        addFiledCategory: koyoOptions.addFiledCategory !== false,
        filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
        useUtcTime: koyoOptions.useUtcTime || false,
        assistantCategories: koyoOptions.assistantCategories || ""
      };
      if (Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent("fileEmail:" + JSON.stringify(payloadData));
      }
      return;
    }

    const isReadFilingMode = initialMode === "file" || !initialMode;
    if (initialMode === "file_multi" || (isReadFilingMode && !Office.context?.mailbox?.item && multiEmailItems.length > 0)) {
      setIsFiled(false);
      setLoading(true);
      setMessage("Preparing to file multiple emails...");
      abortControllerRef.current = new AbortController();

      try {
        const selectedLocations = locations.filter((x) => selectedIds.includes(x.id));
        if (selectedLocations.length === 0) {
          throw new Error("Select at least one target location.");
        }
        
        const disconnected = selectedLocations.filter(loc => connectivityStatus[loc.id] === false);
        if (disconnected.length > 0) {
          const paths = disconnected.map(d => d.path.split("\\").pop()).join(", ");
          throw new Error(`Filing failed: Location(s) [${paths}] are disconnected. Please check your network connection.`);
        }

        let graphAccessToken = null;
        let ssoTokenForFiling = null;
        try {
          const tokenResult = await getGraphToken({
            msalInstance: instance,
            interactive: false,
            loginHint: Office?.context?.mailbox?.userProfile?.emailAddress,
          });
          setAuthTier(tokenResult.tier);
          // SSO identity tokens (tier="sso") must NOT be sent as direct Graph tokens.
          // The backend uses OBO exchange for ssoToken; isAccessToken=true for graphAccessToken.
          if (tokenResult.tier === "sso") {
            ssoTokenForFiling = tokenResult.token;
          } else {
            graphAccessToken = tokenResult.token;
          }
        } catch (tokenErr) {
          console.warn("[App] No graph token available for multi-file:", tokenErr?.message);
        }

        if (koyoOptions.addFiledCategory !== false) {
          const categoryName = koyoOptions.filedCategoryName || "Filed by Koyomail";
          try {
            await ensureMasterCategory(categoryName, "Preset3");
          } catch (catErr) {
            console.warn("[App] Failed to ensure master category:", catErr.message);
          }
        }

        let filedCount = 0;
        let skippedCount = 0;
        let draftEmailCreatedOverall = false;
        let allSharingLinks = [];
        let accumulatedErrors = "";

        const items = [...multiEmailItems];
        let completedCount = 0;
        const totalCount = items.length;

        const updateProgress = () => {
          setMessage(`Filing emails (${completedCount}/${totalCount} completed)...`);
        };

        const executeFiling = async () => {
          while (items.length > 0) {
            const item = items.shift();
            if (!item) break;

            const validatedGraphAccessToken = (typeof graphAccessToken === "string" && graphAccessToken.length > 10) 
              ? graphAccessToken 
              : null;
            const validatedSsoToken = (typeof ssoTokenForFiling === "string" && ssoTokenForFiling.length > 10)
              ? ssoTokenForFiling
              : null;

            const payloadData = {
              itemId: toGraphItemId(item.itemId),
              subject: item.subject,
              graphAccessToken: validatedGraphAccessToken,
              ssoToken: validatedSsoToken,
              isPartial: false,
              targetPaths: selectedLocations.map(l => l.folder || l.path),
              comment,
              attachmentsOption,
              markReviewed,
              sendLink,
              skipDraftCreation: true,
              afterFiling: afterFiling || "none",
              addFiledCategory: koyoOptions.addFiledCategory !== false,
              filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
              useUtcTime: koyoOptions.useUtcTime || false,
              assistantCategories: koyoOptions.assistantCategories || "",
              duplicateStrategy: koyoOptions.duplicateStrategy || "rename",
              deleteEmptyFolders: koyoOptions.deleteEmptyFolders || false,
              filedFolderPrefix: koyoOptions.filedFolderPrefix || "*",
              applyReadOnly: koyoOptions.applyReadOnly || false
            };

            try {
              const response = await fileEmail(payloadData, { signal: abortControllerRef.current.signal });
              if (response && response.sharingLinks) {
                allSharingLinks.push(...response.sharingLinks);
              }
              if (response && response.postFilingError) {
                accumulatedErrors += `[${item.subject}] ${response.postFilingError}\n`;
              }
              
              const isFullySkipped = response && response.results && response.results.length > 0 && response.results.every(r => r.status === "skipped");
              if (isFullySkipped) {
                skippedCount++;
              } else {
                filedCount++;
              }

              if (afterFiling && afterFiling !== "none") {
                 if (response?.postFilingError || !validatedGraphAccessToken) {
                   if (afterFiling === "delete") {
                     await deleteItemViaEws(item.itemId);
                   } else if (afterFiling === "archive") {
                     await moveItemViaEws(item.itemId, "archive");
                   }
                 }
              }
            } catch (e) {
              console.error("Failed to file item", item.itemId, e);
              accumulatedErrors += `[${item.subject}] ${e.message}\n`;
            } finally {
              completedCount++;
              updateProgress();
            }
          }
        };

        // Initialize progress
        updateProgress();

        // Run with concurrency limit of 3
        const concurrencyLimit = Math.min(3, totalCount);
        const workers = Array(concurrencyLimit).fill(null).map(() => executeFiling());
        await Promise.all(workers);
        
        let singleDraftCreated = false;
        let singleDraftId = null;
        const validatedGraphAccessTokenForDraft = (typeof graphAccessToken === "string" && graphAccessToken.length > 10) 
          ? graphAccessToken 
          : null;

        if (sendLink && allSharingLinks.length > 0 && validatedGraphAccessTokenForDraft) {
          try {
            setMessage("Creating consolidated draft email in Drafts folder...");
            const draftResponse = await createDraftEmail({
              graphAccessToken: validatedGraphAccessTokenForDraft,
              originalSubject: `Multiple Emails (${filedCount})`,
              comment,
              filedEntries: allSharingLinks,
              emailFont: koyoOptions.emailFont || "Segoe UI",
              fontSize: koyoOptions.fontSize || "11"
            });
            if (draftResponse && draftResponse.success) {
              singleDraftCreated = true;
              singleDraftId = draftResponse.draftId || null;
            }
          } catch (draftErr) {
            console.warn("[App] Consolidated draft creation failed:", draftErr.message);
          }
        }
        let msg = "";
        if (filedCount === 0 && skippedCount === multiEmailItems.length) {
          msg = `All ${skippedCount} emails are already filed.`;
          if (afterFiling !== "none" || markReviewed) {
            msg += " (Post-filing actions skipped).";
          }
        } else if (skippedCount > 0) {
          msg = `Filed ${filedCount} emails. ${skippedCount} emails were already filed and skipped.`;
          if (afterFiling !== "none" || markReviewed) {
            msg += " (Post-filing actions skipped for duplicates).";
          }
        } else {
          msg = `Successfully filed ${filedCount} of ${multiEmailItems.length} emails.`;
        }
        
        if (accumulatedErrors) {
          msg += ` Some post-filing actions failed, check console.`;
          console.warn("Multi-file errors:", accumulatedErrors);
        }
        
        if (sendLink && allSharingLinks.length > 0) {
          if (singleDraftCreated) {
            if (singleDraftId) {
              try {
                Office.context.mailbox.displayMessageForm(singleDraftId);
              } catch (openErr) {
                console.warn("[App] Failed to open consolidated draft compose window:", openErr);
              }
            }
            msg += " A draft email containing all filing links has been created in your Drafts folder.";
          } else {
            openComposeWindow(allSharingLinks, `Multiple Emails (${filedCount})`);
            msg += " Compose window opened with filing links.";
          }
        }

        setMessage(msg);
        setIsFiled(true);

        if (!accumulatedErrors) {
          setTimeout(() => {
            if (Office?.context?.ui?.closeContainer) {
              Office.context.ui.closeContainer();
            } else if (Office.context.ui?.messageParent) {
              Office.context.ui.messageParent("close");
            } else {
              window.close();
            }
          }, 1500);
        }
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        setMessage(`Filing failed: ${errorMsg}`);
      } finally {
        setLoading(false);
        abortControllerRef.current = null;
      }
      return;
    }

    setIsFiled(false);
    setLoading(true);
    setMessage("Preparing to file...");
    abortControllerRef.current = new AbortController();

    try {
      const selectedLocations = locations.filter((x) => selectedIds.includes(x.id));
      if (selectedLocations.length === 0) {
        throw new Error("Select at least one target location.");
      }

      // Check connectivity for all selected locations
      const disconnected = selectedLocations.filter(loc => connectivityStatus[loc.id] === false);
      if (disconnected.length > 0) {
        const paths = disconnected.map(d => d.path.split("\\").pop()).join(", ");
        throw new Error(`Filing failed: Location(s) [${paths}] are disconnected. Please check your network connection.`);
      }

      setMessage("Filing email...");
      let basePayload = emailPayload;
      if (!basePayload || basePayload.isPartial || !basePayload.itemId) {
        try {
          const latestPayload = await buildCurrentEmailPayload(
            basePayload?.isPartial ? { forceRefresh: true } : undefined
          );
          basePayload = latestPayload || basePayload;
        } catch (err) {
          console.log("Using memory payload fallback");
        }
      }
      if (!basePayload) {
        throw new Error("Email content is not ready yet. Please wait a moment.");
      }
      if (basePayload.isPartial) {
        setMessage("Body enrichment is taking longer than expected. Filing with available preview content...");
      }

      const needsGraphPostActions = afterFiling !== "none" || markReviewed || sendLink || (koyoOptions.addFiledCategory !== false);
      const categoryName = koyoOptions.addFiledCategory !== false
        ? (koyoOptions.filedCategoryName || "Filed by Koyomail")
        : null;

      let graphAccessToken = null;
      let ssoTokenForFiling = null;

      const tokenPromise = (basePayload?.itemId && (!basePayload?.ssoToken || needsGraphPostActions))
        ? getGraphToken({
            msalInstance: instance,
            interactive: false,
            loginHint: Office?.context?.mailbox?.userProfile?.emailAddress,
          }).then((tokenResult) => {
            setAuthTier(tokenResult.tier);
            return tokenResult;
          }).catch((tokenErr) => {
            console.warn("[App] No graph token available for filing:", tokenErr?.message || tokenErr);
            return null;
          })
        : Promise.resolve(null);

      const categoryPromise = categoryName
        ? ensureMasterCategory(categoryName, "Preset3").catch((catErr) => {
            console.warn("[App] Failed to ensure master category locally before filing:", catErr.message);
          })
        : Promise.resolve();

      const [tokenResult] = await Promise.all([tokenPromise, categoryPromise]);
      if (tokenResult?.token) {
        if (tokenResult.tier === "sso") {
          ssoTokenForFiling = tokenResult.token;
        } else {
          graphAccessToken = tokenResult.token;
        }
      }

      if (attachmentsOption !== "message") {
        const attList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
        const pendingAttachments = attList.filter((att) => {
          const hasContent = !!att?.base64Content;
          const isMetadataOnly = !!att?.isMetadataOnly;
          const isInline = !!att?.isInline;
          const size = Number(att?.size || 0);
          return (isMetadataOnly || !hasContent) && !isInline && size > 0;
        });
        const hasPendingAttachments = pendingAttachments.length > 0;

        // If no Graph-capable token exists, pending attachment metadata may still be incomplete.
        if (!basePayload?.ssoToken && !graphAccessToken && hasPendingAttachments) {
          try {
            const retryPayload = await buildCurrentEmailPayload({ forceRefresh: true });
            basePayload = retryPayload || basePayload;
          } catch (retryErr) {
            console.warn("[App] Could not retry payload for attachments (likely in dialog):", retryErr.message);
          }

          const retryList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
          const retryPendingAttachments = retryList.filter((att) => {
            const hasContent = !!att?.base64Content;
            const isMetadataOnly = !!att?.isMetadataOnly;
            const isInline = !!att?.isInline;
            const size = Number(att?.size || 0);
            return (isMetadataOnly || !hasContent) && !isInline && size > 0;
          });

          if (!basePayload?.ssoToken && !graphAccessToken && retryPendingAttachments.length > 0) {
            throw new Error("Attachments are still loading. Please wait a few seconds and try again to avoid missing attachments.");
          }
        }
      }
      
      // Filter attachments based on user selection
      let finalAttachments = basePayload.attachments || [];
      if (attachmentsOption === "message") {
        finalAttachments = [];
      } else if (attachmentsOption === "attachments") {
        // Keep as is, but logic favors attachments
      }

      console.log("[App] Filing email with payload:", {
        subject,
        attachmentCount: finalAttachments.length,
        attachmentsOption,
        targetPaths: selectedLocations.map((x) => x.path)
      });

      const validatedGraphAccessToken = (typeof graphAccessToken === "string" && graphAccessToken.length > 10)
        ? graphAccessToken
        : null;
      const validatedSsoToken = (typeof ssoTokenForFiling === "string" && ssoTokenForFiling.length > 10)
        ? ssoTokenForFiling
        : (basePayload?.ssoToken || null);

      const response = await fileEmail({
        ...basePayload,
        graphAccessToken: validatedGraphAccessToken,
        ssoToken: validatedSsoToken,
        masterCategoryEnsured: !!categoryName,
        attachments: finalAttachments,
        subject,
        comment,
        afterFiling,
        markReviewed,
        sendLink,
        attachmentsOption,
        duplicateStrategy: koyoOptions.duplicateStrategy || "rename",
        deleteEmptyFolders: koyoOptions.deleteEmptyFolders || false,
        filedFolderPrefix: koyoOptions.filedFolderPrefix || "*",
        fileReplyingTo: koyoOptions.fileReplyingTo || false,
        targetPaths: selectedLocations.map((x) => x.path),
        applyReadOnly: koyoOptions.applyReadOnly || false,
        useUtcTime: koyoOptions.useUtcTime || false,
        addFiledCategory: koyoOptions.addFiledCategory !== false,
        filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
        assistantCategories: koyoOptions.assistantCategories || "",
        emailFont: koyoOptions.emailFont || "Times New Roman",
        fontSize: koyoOptions.fontSize || "10",
      }, { signal: abortControllerRef.current.signal });

      // Check for skipped status
      const isFullySkipped = response?.results && response.results.length > 0 && response.results.every(r => r.status === "skipped");
      const isPartiallySkipped = response?.results && response.results.some(r => r.status === "skipped") && response.results.some(r => r.status !== "skipped");
      let postFilingHandled = !response?.postFilingError;

      if (response?.postFilingError) {
        // If the error was just about adding the category, we can ignore it if we succeed locally
        setActionError(response.postFilingError);
        setMessage(isFullySkipped ? "This email is already filed, but post-filing action failed." : "Email filed successfully, but post-filing action failed.");

        try {
          const recovery = await recoverPostFilingAfterGraphFailure({
            postFilingError: response.postFilingError,
            itemId: basePayload?.itemId,
            afterFiling,
            markReviewed,
            addFiledCategory: koyoOptions.addFiledCategory !== false,
            filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
          });
          if (recovery.recovered) {
            postFilingHandled = true;
            setActionError("");
            const actionLabel = afterFiling !== "none"
              ? `Post-filing action (${afterFiling}) completed in Outlook.`
              : "Post-filing actions completed in Outlook.";
            setMessage(isFullySkipped
              ? `This email is already filed. ${actionLabel}`
              : `Email filed successfully. ${actionLabel}`);
          }
        } catch (recoveryErr) {
          console.warn("[App] Client post-filing recovery failed:", recoveryErr.message);
          if (isGraphPostFilingDeferralError(response.postFilingError)) {
            setActionError(formatAfterFilingApiError(recoveryErr, "Post-filing action", basePayload?.itemId));
          }
        }
      } else {
        if (isFullySkipped) {
          const skippedActionsMsg = (afterFiling !== "none" || markReviewed) ? " (Post-filing actions skipped)." : "";
          setMessage(`This email is already filed.${skippedActionsMsg}`);
        } else if (isPartiallySkipped) {
          setMessage(`Email filed to new locations (already filed in some).${basePayload?.isPartial ? " Note: Some attachments may be missing." : ""}`);
        } else {
          setMessage(`Email filed successfully.${basePayload?.isPartial ? " Note: Some attachments may be missing." : ""}`);
        }
      }
      
      // Client-side category only when backend post-filing did not complete.
      if (categoryName && !postFilingHandled) {
        try {
           await addCategoryToCurrentEmail(categoryName);
        } catch (e) {
           console.warn("Client-side categorization failed:", e);
        }
      }
      
      // If generate link was requested, draft email AND copy link to clipboard
      if (sendLink && response?.sharingLinks?.length > 0) {
        const linkText = response.sharingLinks.join("\n");

        // Always copy to clipboard so user can Ctrl+V the clickable link anywhere
        let clipboardOk = false;
        try {
          await navigator.clipboard.writeText(linkText);
          clipboardOk = true;
        } catch (clipErr) {
          console.warn("[App] Clipboard write failed:", clipErr);
        }

        if (response.draftEmailCreated) {
          // Backend successfully created a draft email — display it!
          if (response.draftId) {
            try {
              Office.context.mailbox.displayMessageForm(response.draftId);
            } catch (openErr) {
              console.warn("[App] Failed to open server-side draft compose window:", openErr);
            }
          }
          setMessage(clipboardOk
            ? "Email filed successfully. Draft email created & link copied to clipboard."
            : "Email filed successfully. A draft email with the filing link has been created in your Drafts folder.");
        } else {
          // No draft — open compose window as fallback
          openComposeWindow(response.sharingLinks, subject);
          setMessage(clipboardOk
            ? "Email filed successfully. Link copied to clipboard & compose window opened."
            : `Filed link(s): ${response.sharingLinks.join(", ")}`);
        }
      }
      
      // Perform after-filing actions locally ONLY if the backend failed to do it (e.g., due to no token)
      const item = Office.context?.mailbox?.item;
      if (afterFiling !== "none" && !basePayload?.ssoToken && !graphAccessToken) {
        if (item && afterFiling === "delete") {
          setActionError("Automatic local delete was skipped to prevent permanent deletion in this Outlook host.");
          setMessage("Email filed successfully. Please move the email to Deleted Items manually.");
          setIsFiled(true);
          return;
        }

        if (item && afterFiling === "archive") {
          if (item.archiveAsync) {
            item.archiveAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Failed) {
                setMessage("Email filed, but failed to Archive: " + (result.error?.message || "Unknown error"));
              } else {
                setMessage("Email filed and Archived.");
              }
            });
          } else {
            setMessage("Email filed, but 'Archive' action is not supported in this version of Outlook.");
          }
          setIsFiled(true);
          return;
        }
        
        // We are likely in a dialog, message the parent to handle the action
        if (Office.context.ui && Office.context.ui.messageParent) {
          setMessage(`Email filed. Requesting Outlook to ${afterFiling === "delete" ? "transfer email to Deleted Items" : "Archive"}...`);
          Office.context.ui.messageParent(JSON.stringify({ action: "afterFiling", value: afterFiling }));
          
          let secondsPassed = 0;
          while (secondsPassed < 10) {
            await new Promise(resolve => setTimeout(resolve, 1000));
            secondsPassed++;
            const storedError = localStorage.getItem("koyomailActionError");
            if (storedError) {
              const { message: parentError } = JSON.parse(storedError);
              localStorage.removeItem("koyomailActionError");
              setActionError(parentError);
              setMessage("Email filed successfully. Automatic move/archive could not be completed in this Outlook host.");
              setIsFiled(true);
              return;
            }
          }
          setMessage(`Filing complete, but Outlook is taking longer than expected to ${afterFiling === "delete" ? "transfer" : "archive"} the email. You may close this window manually.`);
        } else {
          setMessage("Email filed, but could not request move/archive (parent context not found).");
        }
      } else if (afterFiling !== "none" && postFilingHandled) {
        setMessage(`Email filed and post-filing action completed via Microsoft Graph.`);
      }

      loadLocations(null, { silent: true, lightweight: true });
      setIsFiled(true);

      if ((isReadFilingMode || initialMode === "file_dialog") && postFilingHandled) {
        setTimeout(() => {
          if (isReadFilingMode && Office.context.ui?.closeContainer) {
            Office.context.ui.closeContainer();
          } else if (Office.context.ui?.messageParent) {
            Office.context.ui.messageParent("close");
          } else {
            window.close();
          }
        }, 1500);
      }

    } catch (error) {
      if (error instanceof Error && error.name === "AbortError") {
        console.log("[App] Filing aborted by user.");
        return;
      }
      console.error("[App] Filing failed:", error);
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Filing failed: ${errorMsg}`);
    } finally {
      abortControllerRef.current = null;
      setLoading(false);
    }
  };

  const onFileToPath = async (targetPath) => {
    setIsFiled(false);
    setLoading(true);
    setMessage("");
    abortControllerRef.current = new AbortController();

    if (initialMode === "onsend") {
      let ssoToken = null;
      try {
        ssoToken = await getToken({ interactive: false });
      } catch (tokenErr) {
        console.warn("[App] Could not acquire SSO token for On-Send category tagging:", tokenErr.message);
      }

      if (koyoOptions.addFiledCategory !== false) {
        const categoryName = koyoOptions.filedCategoryName || "Filed by Koyomail";
        try {
          await ensureMasterCategory(categoryName);
        } catch (catErr) {
          console.warn("[App] Failed to ensure master category:", catErr.message);
        }
      }

      const payloadData = {
        paths: [targetPath],
        subject,
        comment,
        attachmentsOption,
        markReviewed,
        sendLink,
        isOnSend: true,
        ssoToken: ssoToken || null,
        afterFiling: afterFiling || "none",
        addFiledCategory: koyoOptions.addFiledCategory !== false,
        filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
        useUtcTime: koyoOptions.useUtcTime || false,
        assistantCategories: koyoOptions.assistantCategories || ""
      };
      if (Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent("fileEmail:" + JSON.stringify(payloadData));
      }
      return;
    }

    const isReadFilingMode = initialMode === "file" || !initialMode;
    const isMultiFlow = initialMode === "file_multi" || (isReadFilingMode && !Office.context?.mailbox?.item && multiEmailItems.length > 0);

    if (isMultiFlow) {
      try {
        let graphAccessToken = null;
        try {
          graphAccessToken = await getToken({ interactive: false });
        } catch (tokenErr) {
          console.warn("[App] No graph token available for multi-file to path:", tokenErr?.message);
        }

        if (koyoOptions.addFiledCategory !== false) {
          const categoryName = koyoOptions.filedCategoryName || "Filed by Koyomail";
          try {
            await ensureMasterCategory(categoryName, "Preset3");
          } catch (catErr) {
            console.warn("[App] Failed to ensure master category:", catErr.message);
          }
        }

        let filedCount = 0;
        let skippedCount = 0;
        let allSharingLinks = [];
        let accumulatedErrors = "";

        const items = [...multiEmailItems];
        let completedCount = 0;
        const totalCount = items.length;

        const updateProgress = () => {
          setMessage(`Filing emails (${completedCount}/${totalCount} completed)...`);
        };

        const executeFiling = async () => {
          while (items.length > 0) {
            const item = items.shift();
            if (!item) break;

            const validatedGraphAccessToken = (typeof graphAccessToken === "string" && graphAccessToken.length > 10) 
              ? graphAccessToken 
              : null;

            const payloadData = {
              itemId: toGraphItemId(item.itemId),
              subject: item.subject,
              graphAccessToken: validatedGraphAccessToken,
              isPartial: false,
              targetPaths: [targetPath],
              comment,
              attachmentsOption,
              markReviewed,
              sendLink,
              skipDraftCreation: true,
              afterFiling: afterFiling || "none",
              addFiledCategory: koyoOptions.addFiledCategory !== false,
              filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
              useUtcTime: koyoOptions.useUtcTime || false,
              assistantCategories: koyoOptions.assistantCategories || "",
              duplicateStrategy: koyoOptions.duplicateStrategy || "rename",
              deleteEmptyFolders: koyoOptions.deleteEmptyFolders || false,
              filedFolderPrefix: koyoOptions.filedFolderPrefix || "*",
              applyReadOnly: koyoOptions.applyReadOnly || false
            };

            try {
              const response = await fileEmail(payloadData, { signal: abortControllerRef.current.signal });
              if (response && response.sharingLinks) {
                allSharingLinks.push(...response.sharingLinks);
              }
              if (response && response.postFilingError) {
                accumulatedErrors += `[${item.subject}] ${response.postFilingError}\n`;
              }
              
              const isFullySkipped = response && response.results && response.results.length > 0 && response.results.every(r => r.status === "skipped");
              if (isFullySkipped) {
                skippedCount++;
              } else {
                filedCount++;
              }

              if (afterFiling && afterFiling !== "none") {
                 if (response?.postFilingError || !validatedGraphAccessToken) {
                   if (afterFiling === "delete") {
                     await deleteItemViaEws(item.itemId);
                   } else if (afterFiling === "archive") {
                     await moveItemViaEws(item.itemId, "archive");
                   }
                 }
              }
            } catch (e) {
              console.error("Failed to file item", item.itemId, e);
              accumulatedErrors += `[${item.subject}] ${e.message}\n`;
            } finally {
              completedCount++;
              updateProgress();
            }
          }
        };

        updateProgress();

        const concurrencyLimit = Math.min(3, totalCount);
        const workers = Array(concurrencyLimit).fill(null).map(() => executeFiling());
        await Promise.all(workers);
        
        let singleDraftCreated = false;
        let singleDraftId = null;
        const validatedGraphAccessTokenForDraft = (typeof graphAccessToken === "string" && graphAccessToken.length > 10) 
          ? graphAccessToken 
          : null;

        if (sendLink && allSharingLinks.length > 0 && validatedGraphAccessTokenForDraft) {
          try {
            setMessage("Creating consolidated draft email in Drafts folder...");
            const draftResponse = await createDraftEmail({
              graphAccessToken: validatedGraphAccessTokenForDraft,
              originalSubject: `Multiple Emails (${filedCount})`,
              comment,
              filedEntries: allSharingLinks,
              emailFont: koyoOptions.emailFont || "Segoe UI",
              fontSize: koyoOptions.fontSize || "11"
            });
            if (draftResponse && draftResponse.success) {
              singleDraftCreated = true;
              singleDraftId = draftResponse.draftId || null;
            }
          } catch (draftErr) {
            console.warn("[App] Consolidated draft creation failed:", draftErr.message);
          }
        }
        let msg = "";
        if (filedCount === 0 && skippedCount === multiEmailItems.length) {
          msg = `All ${skippedCount} emails are already filed.`;
          if (afterFiling !== "none" || markReviewed) {
            msg += " (Post-filing actions skipped).";
          }
        } else if (skippedCount > 0) {
          msg = `Filed ${filedCount} emails. ${skippedCount} emails were already filed and skipped.`;
          if (afterFiling !== "none" || markReviewed) {
            msg += " (Post-filing actions skipped for duplicates).";
          }
        } else {
          msg = `Successfully filed ${filedCount} of ${multiEmailItems.length} emails.`;
        }
        
        if (accumulatedErrors) {
          msg += ` Some post-filing actions failed, check console.`;
          console.warn("Multi-file errors:", accumulatedErrors);
        }
        
        if (sendLink && allSharingLinks.length > 0) {
          if (singleDraftCreated) {
            if (singleDraftId) {
              try {
                Office.context.mailbox.displayMessageForm(singleDraftId);
              } catch (openErr) {
                console.warn("[App] Failed to open consolidated draft compose window:", openErr);
              }
            }
            msg += " A draft email containing all filing links has been created in your Drafts folder.";
          } else {
            openComposeWindow(allSharingLinks, `Multiple Emails (${filedCount})`);
            msg += " Compose window opened with filing links.";
          }
        }

        setMessage(msg);
        setIsFiled(true);

        if (!accumulatedErrors) {
          setTimeout(() => {
            if (Office?.context?.ui?.closeContainer) {
              Office.context.ui.closeContainer();
            } else if (Office.context.ui?.messageParent) {
              Office.context.ui.messageParent("close");
            } else {
              window.close();
            }
          }, 1500);
        }
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        setMessage(`Filing failed: ${errorMsg}`);
      } finally {
        setLoading(false);
        abortControllerRef.current = null;
      }
      return;
    }

    try {
      // Check connectivity for the target path
      const loc = locations.find(x => x.path === targetPath);
      if (loc && connectivityStatus[loc.id] === false) {
        throw new Error(`Filing failed: Location is disconnected. Please check your network connection.`);
      }

      setMessage("Filing email...");
      let basePayload = emailPayload;
      if (!basePayload || basePayload.isPartial || !basePayload.itemId) {
        try {
          const latestPayload = await buildCurrentEmailPayload(
            basePayload?.isPartial ? { forceRefresh: true } : undefined
          );
          basePayload = latestPayload || basePayload;
        } catch (refreshErr) {
          console.warn("[App] Could not refresh payload to path:", refreshErr.message);
        }
      }
      if (!basePayload) {
        throw new Error("Email content is not ready yet. Please wait a moment.");
      }
      if (basePayload.isPartial) {
        setMessage("Body enrichment is taking longer than expected. Filing with available preview content.");
      }

      const needsGraphPostActions = afterFiling !== "none" || markReviewed || sendLink || (koyoOptions.addFiledCategory !== false);
      const categoryName = koyoOptions.addFiledCategory !== false
        ? (koyoOptions.filedCategoryName || "Filed by Koyomail")
        : null;

      let graphAccessToken = null;
      let ssoTokenForFiling = null;

      const tokenPromise = (basePayload?.itemId && (!basePayload?.ssoToken || needsGraphPostActions))
        ? getGraphToken({
            msalInstance: instance,
            interactive: false,
            loginHint: Office?.context?.mailbox?.userProfile?.emailAddress,
          }).then((tokenResult) => {
            setAuthTier(tokenResult.tier);
            return tokenResult;
          }).catch((tokenErr) => {
            console.warn("[App] No graph token available for path filing:", tokenErr?.message || tokenErr);
            return null;
          })
        : Promise.resolve(null);

      const categoryPromise = categoryName
        ? ensureMasterCategory(categoryName, "Preset3").catch(() => {})
        : Promise.resolve();

      const [tokenResult] = await Promise.all([tokenPromise, categoryPromise]);
      if (tokenResult?.token) {
        if (tokenResult.tier === "sso") {
          ssoTokenForFiling = tokenResult.token;
        } else {
          graphAccessToken = tokenResult.token;
        }
      }

      if (attachmentsOption !== "message") {
        const attList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
        const pendingAttachments = attList.filter((att) => {
          const hasContent = !!att?.base64Content;
          const isMetadataOnly = !!att?.isMetadataOnly;
          const isInline = !!att?.isInline;
          const size = Number(att?.size || 0);
          return (isMetadataOnly || !hasContent) && !isInline && size > 0;
        });
        const hasPendingAttachments = pendingAttachments.length > 0;

        if (!basePayload?.ssoToken && !graphAccessToken && hasPendingAttachments) {
          try {
            const retryPayload = await buildCurrentEmailPayload({ forceRefresh: true });
            basePayload = retryPayload || basePayload;
          } catch (retryErr) {
            console.warn("[App] Could not retry payload to path for attachments (likely in dialog):", retryErr.message);
          }

          const retryList = Array.isArray(basePayload.attachments) ? basePayload.attachments : [];
          const retryPendingAttachments = retryList.filter((att) => {
            const hasContent = !!att?.base64Content;
            const isMetadataOnly = !!att?.isMetadataOnly;
            const isInline = !!att?.isInline;
            const size = Number(att?.size || 0);
            return (isMetadataOnly || !hasContent) && !isInline && size > 0;
          });

          if (!basePayload?.ssoToken && !graphAccessToken && retryPendingAttachments.length > 0) {
            throw new Error("Attachments are still loading. Please wait a few seconds and try again to avoid missing attachments.");
          }
        }
      }

      // Filter attachments based on user selection
      let finalAttachments = basePayload.attachments || [];
      if (attachmentsOption === "message") {
        finalAttachments = [];
      }

      const validatedGraphAccessToken = (typeof graphAccessToken === "string" && graphAccessToken.length > 10)
        ? graphAccessToken
        : null;
      const validatedSsoToken = (typeof ssoTokenForFiling === "string" && ssoTokenForFiling.length > 10)
        ? ssoTokenForFiling
        : (basePayload?.ssoToken || null);

      const response = await fileEmail({
        ...basePayload,
        graphAccessToken: validatedGraphAccessToken,
        ssoToken: validatedSsoToken,
        masterCategoryEnsured: !!categoryName,
        attachments: finalAttachments,
        subject,
        comment,
        afterFiling,
        markReviewed,
        sendLink,
        attachmentsOption,
        duplicateStrategy: koyoOptions.duplicateStrategy || "rename",
        targetPaths: [targetPath],
        applyReadOnly: koyoOptions.applyReadOnly || false,
        useUtcTime: koyoOptions.useUtcTime || false,
        deleteEmptyFolders: koyoOptions.deleteEmptyFolders || false,
        filedFolderPrefix: koyoOptions.filedFolderPrefix || "*",
        fileReplyingTo: koyoOptions.fileReplyingTo || false,
        addFiledCategory: koyoOptions.addFiledCategory !== false,
        filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
      }, { signal: abortControllerRef.current.signal });

      // Check for skipped status
      const isFullySkipped = response?.results && response.results.length > 0 && response.results.every(r => r.status === "skipped");
      let postFilingHandled = !response?.postFilingError;

      if (response?.postFilingError) {
        setActionError(response.postFilingError);
        setMessage(isFullySkipped ? "This email is already filed, but post-filing action failed." : "Email filed successfully, but post-filing action failed.");

        try {
          const recovery = await recoverPostFilingAfterGraphFailure({
            postFilingError: response.postFilingError,
            itemId: basePayload?.itemId,
            afterFiling,
            markReviewed,
            addFiledCategory: koyoOptions.addFiledCategory !== false,
            filedCategoryName: koyoOptions.filedCategoryName || "Filed by Koyomail",
          });
          if (recovery.recovered) {
            postFilingHandled = true;
            setActionError("");
            const actionLabel = afterFiling !== "none"
              ? `Post-filing action (${afterFiling}) completed in Outlook.`
              : "Post-filing actions completed in Outlook.";
            setMessage(isFullySkipped
              ? `This email is already filed. ${actionLabel}`
              : `Email filed successfully. ${actionLabel}`);
          }
        } catch (recoveryErr) {
          console.warn("[App] Client post-filing recovery failed:", recoveryErr.message);
          if (isGraphPostFilingDeferralError(response.postFilingError)) {
            setActionError(formatAfterFilingApiError(recoveryErr, "Post-filing action", basePayload?.itemId));
          }
        }
      } else {
        if (isFullySkipped) {
          const skippedActionsMsg = (afterFiling !== "none" || markReviewed) ? " (Post-filing actions skipped)." : "";
          setMessage(`This email is already filed.${skippedActionsMsg}`);
        } else {
          setMessage(`Email filed successfully.${basePayload?.isPartial ? " Note: Some attachments may be missing." : ""}`);
        }
      }

      // If generate link was requested, draft email AND copy link to clipboard
      if (sendLink && response?.sharingLinks?.length > 0) {
        const linkText = response.sharingLinks.join("\n");

        // Always copy to clipboard so user can Ctrl+V the clickable link anywhere
        let clipboardOk = false;
        try {
          await navigator.clipboard.writeText(linkText);
          clipboardOk = true;
        } catch (clipErr) {
          console.warn("[App] Clipboard write failed:", clipErr);
        }

        if (response.draftEmailCreated) {
          setMessage(clipboardOk
            ? "Email filed successfully. Draft email created & link copied to clipboard."
            : "Email filed successfully. A draft email with the filing link has been created in your Drafts folder.");
        } else {
          openComposeWindow(response.sharingLinks, subject);
          setMessage(clipboardOk
            ? "Email filed successfully. Link copied to clipboard & compose window opened."
            : `Filed link(s): ${response.sharingLinks.join(", ")}`);
        }
      }

      // Perform after-filing actions locally ONLY if the backend failed to do it (e.g., due to no token)
      const item = Office.context?.mailbox?.item;
      if (afterFiling !== "none" && !basePayload?.ssoToken && !graphAccessToken) {
        if (item && afterFiling === "delete") {
          setActionError("Automatic local delete was skipped to prevent permanent deletion in this Outlook host.");
          setMessage("Email filed successfully. Please move the email to Deleted Items manually.");
          await loadLocations(null, { silent: true });
          setIsFiled(true);
          return;
        }

        if (item && afterFiling === "archive") {
          if (item.archiveAsync) {
            item.archiveAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Failed) {
                setMessage("Email filed, but failed to Archive: " + (result.error?.message || "Unknown error"));
              } else {
                setMessage("Email filed and Archived.");
              }
            });
          } else {
            setMessage("Email filed, but 'Archive' action is not supported in this version of Outlook.");
          }
          await loadLocations(null, { silent: true });
          setIsFiled(true);
          return;
        }
        
        // We are likely in a dialog, message the parent to handle the action
        if (Office.context.ui && Office.context.ui.messageParent) {
          setMessage(`Email filed. Requesting Outlook to ${afterFiling === "delete" ? "transfer email to Deleted Items" : "Archive"}...`);
          Office.context.ui.messageParent(JSON.stringify({ action: "afterFiling", value: afterFiling }));
          
          let secondsPassed = 0;
          while (secondsPassed < 10) {
            await new Promise(resolve => setTimeout(resolve, 1000));
            secondsPassed++;
            const storedError = localStorage.getItem("koyomailActionError");
            if (storedError) {
              const { message: parentError } = JSON.parse(storedError);
              localStorage.removeItem("koyomailActionError");
              setActionError(parentError);
              setMessage("Email filed successfully. Automatic move/archive could not be completed in this Outlook host.");
              await loadLocations(null, { silent: true });
              setIsFiled(true);
              return;
            }
          }
          setMessage(`Filing complete, but Outlook is taking longer than expected to ${afterFiling === "delete" ? "transfer" : "archive"} the email. You may close this window manually.`);
        } else {
          setMessage("Email filed, but could not request move/archive (parent context not found).");
        }
      } else if (afterFiling !== "none" && postFilingHandled) {
        setMessage(`Email filed and post-filing action completed via Microsoft Graph.`);
      }
      
      loadLocations(null, { silent: true, lightweight: true }); // Refresh to update lastUsedAt
      setIsFiled(true);

      if ((isReadFilingMode || initialMode === "file_dialog") && postFilingHandled) {
        setTimeout(() => {
          if (isReadFilingMode && Office.context.ui?.closeContainer) {
            Office.context.ui.closeContainer();
          } else if (Office.context.ui?.messageParent) {
            Office.context.ui.messageParent("close");
          } else {
            window.close();
          }
        }, 1500);
      }
      
      if (initialMode === "onsend" && Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent("allowSend");
      }
    } catch (error) {
      if (error instanceof Error && error.name === "AbortError") {
        console.log("[App] Filing to path aborted by user.");
        return;
      }
      console.error("[App] Filing to path failed:", error);
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Filing failed: ${errorMsg}`);
    } finally {
      abortControllerRef.current = null;
      setLoading(false);
    }
  };
  
  const handleCancelClick = () => {
    if (loading && abortControllerRef.current) {
      abortControllerRef.current.abort();
      setLoading(false);
      setMessage("Filing cancelled.");
    } else {
      if (Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent("close");
      } else if (Office.context.ui && Office.context.ui.closeContainer) {
        Office.context.ui.closeContainer();
      } else {
        window.close();
      }
    }
  };

  const handleCloseClick = () => {
    if (initialMode === "onsend") {
      if (Office.context.ui && Office.context.ui.messageParent) {
        Office.context.ui.messageParent("allowSend");
      }
      return;
    }

    if (Office.context.ui && Office.context.ui.messageParent) {
      Office.context.ui.messageParent("close");
    } else if (Office.context.ui && Office.context.ui.closeContainer) {
      Office.context.ui.closeContainer();
    } else {
      window.close();
    }
  };

  const onExplore = async () => {
    if (selectedIds.length > 1) {
      setMessage("Please select at most one location to explore.");
      return;
    }

    if (selectedIds.length === 0) {
      try {
        await exploreLocation("");
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
        setMessage(`Explore failed: ${errorMsg}`);
      }
      return;
    }

    const loc = locations.find((x) => x.id === selectedIds[0]);
    if (loc) {
      try {
        await exploreLocation(loc.path);
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
        setMessage(`Explore failed: ${errorMsg}`);
      }
    }
  };

  const onRemoveSuggestion = async () => {
    if (selectedIds.length === 0) {
      setMessage("Please select at least one location to toggle favourite.");
      return;
    }
    try {
      const sender = emailPayloadRef.current?.sender;
      for (const id of selectedIds) {
        await toggleSuggestion(id, sender);
      }
      setMessage("Favourites updated.");
      await loadLocations(null, { silent: true });
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Update failed: ${errorMsg}`);
    }
  };

  const onToggleSuggestion = async (id) => {
    const sender = emailPayloadRef.current?.sender;

    // Optimistically update the UI state to change the star icon and position instantly
    setLocations((prevLocations) => {
      const updated = prevLocations.map((loc) => {
        if (loc.id === id) {
          const newFav = !loc.isSuggested;
          return {
            ...loc,
            isSuggested: newFav,
            isSenderSuggested: newFav ? true : loc.isSenderSuggested
          };
        }
        return loc;
      });

      // Instantly re-sort using the helper function
      return sortLocationsList(
        updated,
        sender,
        senderStatsRef.current,
        generalStatsRef.current
      );
    });

    try {
      await toggleSuggestion(id, sender);
      await loadLocations(sender, { silent: true });
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Toggle failed: ${errorMsg}`);
      // Revert/sync on error
      await loadLocations(sender, { silent: true });
    }
  };

  const onMarkUnused = async () => {
    if (selectedIds.length === 0) {
      setMessage("Please select at least one location to mark as unused.");
      return;
    }
    try {
      for (const id of selectedIds) {
        await markLocationUnused(id);
      }
      setMessage("Locations marked as unused.");
      await loadLocations(null, { silent: true });
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : (typeof error === "object" ? JSON.stringify(error) : String(error));
      setMessage(`Mark failed: ${errorMsg}`);
    }
  };

  console.log("[App] initialMode:", initialMode);
  if (initialMode === "collections" || initialMode === "locations") {
    return <LocationsManagerDialog 
      isOpen={true} 
      onOpenChange={(isOpen) => {
        if (!isOpen) {
          if (Office.context.ui && Office.context.ui.messageParent) {
            Office.context.ui.messageParent("close");
          } else {
            window.close();
          }
        }
      }} 
    />;
  }

  if (initialMode === "help") return <HelpDialog isOpen={true} onOpenChange={() => Office.context.ui?.messageParent?.("close")} />;
  if (initialMode === "options") return <OptionsDialog isOpen={true} initialTab={optionsInitialTab} onOpenChange={() => Office.context.ui?.messageParent?.("close")} />;
  if (initialMode === "search") return (
    <div style={{ position: "fixed", inset: 0, zIndex: 9999, backgroundColor: "#f8f8f8" }}>
      <SearchDialog 
        onClose={() => Office.context.ui?.messageParent?.("close")} 
        onOpenSearchOptions={() => { setOptionsInitialTab("Search"); setIsOptionsOpen(true); }} 
      />
      <OptionsDialog 
        isOpen={isOptionsOpen}
        initialTab={optionsInitialTab}
        onOpenChange={(isOpen) => setIsOptionsOpen(isOpen)}
      />
    </div>
  );

  const hasUnusedSelected = selectedIds.length > 0 && locations.some(l => selectedIds.includes(l.id) && l.isUnused);
  
  const selectedLocs = locations.filter(l => selectedIds.includes(l.id));
  const isCollectionLocation = (loc) => loc.collection && !["Private", "Personal", "Portfolio", "Archive", "Discovered"].includes(loc.collection);
  const hasCollectionSelected = selectedLocs.some(isCollectionLocation);
  const hasDisconnectedSelected = selectedLocs.some(loc => connectivityStatus[loc.id] === false);

  const handleToolbarAdd = () => {
    setEditingLocation(null);
    setIsDialogOpen(true);
  };

  const handleToolbarEdit = () => {
    if (selectedIds.length === 1) {
      const loc = locations.find(l => l.id === selectedIds[0]);
      if (loc) {
        setEditingLocation(loc);
        setIsDialogOpen(true);
      }
    }
  };

  const handleToolbarDelete = async () => {
    if (selectedIds.length === 0) return;
    if (window.confirm("Are you sure you want to delete the selected location(s)?")) {
      try {
        for (const id of selectedIds) {
          await deleteLocation(id);
        }
        setMessage("Location(s) deleted.");
        setSelectedIds([]);
        loadLocations(null, { silent: true });
        localStorage.setItem("koyomail_locations_updated", Date.now().toString());
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        setMessage(`Failed to delete location(s): ${errorMsg}`);
      }
    }
  };

  return (
    <div style={{ height: "100vh", display: "flex", flexDirection: "column", fontFamily: "'Exo 2', 'Segoe UI', sans-serif" }}>
      <Toolbar 
        locations={locations}
        onFileToPath={onFileToPath}
        onExplore={onExplore}
        onRefresh={handleRefresh}
        onRemoveSuggestion={onRemoveSuggestion}
        onMarkUnused={onMarkUnused}
        onToggleMultiSelect={() => {
          const newState = !isMultiSelect;
          setIsMultiSelect(newState);
          setMessage(newState ? "Multi-select enabled: You can now select multiple locations." : "Multi-select disabled.");
        }}
        isMultiSelect={isMultiSelect}
        onHelp={() => setIsHelpOpen(true)}
        onAddLocation={handleToolbarAdd}
        onEditLocation={handleToolbarEdit}
        onDeleteLocation={handleToolbarDelete}
        isAuthOk={graphAuthOk}
        hasUnusedSelected={hasUnusedSelected}
        hasCollectionSelected={hasCollectionSelected}
        hasSingleSelection={selectedIds.length === 1}
        hasSelection={selectedIds.length > 0}
      />

      <div style={{ display: "flex", flexWrap: "nowrap", flexGrow: 1, overflow: "hidden", flexDirection: "column" }}>
        {/* Broken collections warning banner */}
        {brokenCollectionNames.length > 0 && (
          <div style={{
            backgroundColor: "#fff4ce",
            borderBottom: "1px solid #f4cd4a",
            padding: "6px 12px",
            display: "flex",
            alignItems: "center",
            gap: 8,
            fontSize: 12,
            color: "#5d4a00",
            flexShrink: 0
          }}>
            <span style={{ fontWeight: 700, fontSize: 14 }}>⚠</span>
            <span style={{ flex: 1 }}>
              <strong>{brokenCollectionNames.length} collection{brokenCollectionNames.length > 1 ? 's' : ''} could not be found</strong>
              {" "}({brokenCollectionNames.join(", ")}).
              {" "}Please open <strong>Collections</strong> and re-add the missing collection file(s).
            </span>
            <button
              onClick={() => setBrokenCollectionNames([])}
              style={{ background: "none", border: "none", cursor: "pointer", color: "#5d4a00", fontSize: 16, lineHeight: 1, padding: "0 4px", fontWeight: 700 }}
              title="Dismiss"
            >×</button>
          </div>
        )}
        <div style={{ display: "flex", flexWrap: "nowrap", flexGrow: 1, overflow: "hidden", position: "relative" }}>
        {koyoOptions.onlyFileUsingDialog ? (
          <div style={{ flex: "1 1 auto", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24, textAlign: "center", backgroundColor: "#faf9f8" }}>
            <h2 style={{ fontSize: 16, fontWeight: "600", marginBottom: 8, color: "#323130" }}>Sidebar filing is disabled</h2>
            <p style={{ fontSize: 13, color: "#605e5c", marginBottom: 24 }}>You have configured Koyomail to only allow filing via the pop-up dialog.<br/>Use the <strong>File Email</strong> button in the Outlook ribbon to open the filing dialog.</p>
          </div>
        ) : (
          <>
            <div style={{ flex: "1 1 auto", minWidth: 280, padding: 8, height: "100%", boxSizing: "border-box", display: "flex", flexDirection: "column" }}>
              {locationsLoading && locations.length === 0 ? (
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", flexGrow: 1, gap: 8 }}>
                  <Spinner size="medium" label="Loading locations..." />
                </div>
              ) : (
                <LocationTable 
                  locations={locations}
                  isLoading={locationsLoading}
                  selectedIds={selectedIds}
                  onSelectionChange={onSelectionChange}
                  connectivityStatus={connectivityStatus}
                  onToggleSuggestion={onToggleSuggestion}
                  onDoubleClickLocation={koyoOptions.enableDoubleClickFiling && graphAuthOk && !noItemSelected ? (path) => {
                    onFileToPath(path);
                  } : undefined}
                  onAddLocation={() => {
                    setEditingLocation(null);
                    setIsDialogOpen(true);
                  }}
                  sender={emailPayload?.sender}
                />
              )}
            </div>

            {((selectedIds.length > 0 && (!isNarrow || !narrowSidebarDismissed)) || (koyoOptions.alwaysShowFilingOptions && !isNarrow)) && !noItemSelected && (
              <DetailsSidebar 
                subject={subject} setSubject={setSubject}
                comment={comment} setComment={(c) => {
                  setComment(c);
                  if (emailPayload?.itemId) {
                    localStorage.setItem(`koyomail_comment_${emailPayload.itemId}`, c);
                  }
                }}
                afterFiling={afterFiling} setAfterFiling={setAfterFiling}
                markReviewed={markReviewed} setMarkReviewed={setMarkReviewed}
                sendLink={sendLink} setSendLink={setSendLink}
                attachmentsOption={attachmentsOption} setAttachmentsOption={setAttachmentsOption}
                onSaveDefaults={saveDefaults}
                mode={initialMode}
                isNarrow={isNarrow}
                onBack={() => setNarrowSidebarDismissed(true)}
              />
            )}
          </>
        )}
        </div>
      </div>

      <div style={{ padding: "8px 12px", borderTop: "1px solid #edebe9", display: "flex", flexDirection: "column", gap: 4, backgroundColor: "#f3f2f1" }}>
        {(!graphAuthOk || !isNarrow || showNarrowAuthSuccess) && (
          <div style={{ 
            fontSize: 13, 
            color: graphAuthOk ? "#107c10" : "#8a6d00", 
            backgroundColor: graphAuthOk ? "#e8f5e8" : "#fff4ce", 
            padding: "4px 8px", 
            borderRadius: 4,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            minHeight: "24px"
          }}>
            <span>{graphAuthStatus}</span>
            {!graphAuthOk && !graphAuthStatus.includes("✓") && (
              <Button 
                size="small"
                appearance="primary"
                onClick={() => {
                  setGraphAuthStatus("Signing in...");
                  getToken({ interactive: true })
                    .then((token) => {
                      if (!token) {
                        throw new Error("No access token was returned.");
                      }
                      setGraphAuthOk(true);
                      setGraphAuthStatus("Signed in ✓");
                      setSsoWarning("");
                    })
                    .catch(err => {
                      // If it's a redirect error, the page navigates — don't update state
                      if (!err.message?.includes("Redirecting")) {
                        setGraphAuthStatus(`Sign in failed: ${err.message}`);
                        setGraphAuthOk(false);
                      }
                    });
                }}
                style={{ padding: "0 8px", height: "24px", minWidth: "auto" }}
              >
                Sign In
              </Button>
            )}
          </div>
        )}
        
        {ssoWarning && !graphAuthOk && <div style={{ fontSize: 13, color: "#7f6700", backgroundColor: "#fef3cd", padding: "4px 8px", borderRadius: 4 }}>{ssoWarning}</div>}
        {actionError && <div style={{ fontSize: 13, color: "#a4262c", backgroundColor: "#fde7e9", padding: "4px 8px", borderRadius: 4 }}>{actionError}</div>}
        
        {!koyoOptions.onlyFileUsingDialog && (
          <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 8 }}>
            {message && <span style={{ flexGrow: 1, minWidth: 0, fontSize: 13, color: message.includes("failed") ? "#a4262c" : "#107c10", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{message}</span>}
            {loading ? (
              <Button style={{ border: "1px solid #c8c6c4" }} onClick={handleCancelClick}>
                Cancel
              </Button>
            ) : (
              <>
                <Button 
                  appearance="primary" 
                  onClick={onFileEmail} 
                  disabled={selectedIds.length === 0 || !graphAuthOk || hasDisconnectedSelected || noItemSelected}
                  title={hasDisconnectedSelected ? "Cannot file because a selected location is disconnected." : undefined}
                >
                  {initialMode === "onsend" ? "Send & File" : "File"}
                </Button>
                {initialMode === "onsend" && (
                  <>
                    <Button style={{ border: "1px solid #c8c6c4" }} onClick={() => Office.context.ui?.messageParent("allowSend")}>
                      Send Only
                    </Button>
                    <Button style={{ border: "1px solid #c8c6c4" }} onClick={() => Office.context.ui?.messageParent("cancelSend")}>
                      Cancel Send
                    </Button>
                  </>
                )}
                <Button style={{ border: "1px solid #c8c6c4" }} onClick={handleCloseClick}>
                  Close
                </Button>
              </>
            )}
          </div>
        )}
      </div>

      <LocationDialog 
        isOpen={isDialogOpen}
        onOpenChange={setIsDialogOpen}
        onSave={onSaveLocation}
        initialData={editingLocation}
      />

      <HelpDialog 
        isOpen={isHelpOpen}
        onOpenChange={(isOpen) => {
          setIsHelpOpen(isOpen);
          const urlParams = new URLSearchParams(window.location.search);
          if (!isOpen && urlParams.get("mode") === "help") {
            if (Office.context.ui && Office.context.ui.messageParent) {
              Office.context.ui.messageParent("close");
            } else {
              window.close();
            }
          }
        }}
      />

      <OptionsDialog 
        isOpen={isOptionsOpen}
        initialTab={optionsInitialTab}
        onOpenChange={(isOpen) => {
          setIsOptionsOpen(isOpen);
          const urlParams = new URLSearchParams(window.location.search);
          if (!isOpen && urlParams.get("mode") === "options") {
            if (Office.context.ui && Office.context.ui.messageParent) {
              Office.context.ui.messageParent("close");
            } else {
              window.close();
            }
          }
        }}
      />

      {/* Full-screen Search mode — rendered as a dialog from the ribbon */}
      {isSearchOpen && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 9999,
          backgroundColor: "#f8f8f8",
        }}>
          <SearchDialog
            onClose={() => {
              if (Office.context.ui && Office.context.ui.messageParent) {
                Office.context.ui.messageParent("close");
              } else {
                setIsSearchOpen(false);
              }
            }}
            onOpenSearchOptions={() => {
              setOptionsInitialTab("Search");
              setIsOptionsOpen(true);
            }}
          />
        </div>
      )}
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
