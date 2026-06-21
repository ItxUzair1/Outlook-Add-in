import * as React from "react";
import {
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
  Checkbox,
  Input,
  Select,
  Button
} from "@fluentui/react-components";
import { Checkmark16Regular, Star16Filled, Star16Regular, Search16Regular } from "@fluentui/react-icons";

const EmptyState = ({ isSearchFilter, onClearFilters, onAddLocation }) => {
  return (
    <div style={{
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      padding: "32px 16px",
      textAlign: "center",
      backgroundColor: "#faf9f8",
      borderRadius: "8px",
      border: "1px dashed #d1d1d1",
      margin: "12px",
      gap: "12px"
    }}>
      <div style={{
        width: "48px",
        height: "48px",
        borderRadius: "50%",
        background: "linear-gradient(135deg, #e1dfdd 0%, #c8c6c4 100%)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        color: "#605e5c",
        fontSize: "24px"
      }}>
        {isSearchFilter ? "🔍" : "📁"}
      </div>
      <div>
        <h3 style={{ margin: "0 0 4px 0", fontSize: "14px", fontWeight: "600", color: "#323130" }}>
          {isSearchFilter ? "No search results" : "No locations added"}
        </h3>
        <p style={{ margin: 0, fontSize: "12px", color: "#605e5c", maxWidth: "240px" }}>
          {isSearchFilter 
            ? "We couldn't find any locations matching your filters." 
            : "Get started by adding your first filing location."}
        </p>
      </div>
      {isSearchFilter ? (
        <span 
          onClick={onClearFilters} 
          style={{ 
            color: "#0078d4", 
            textDecoration: "underline", 
            cursor: "pointer", 
            fontSize: "13px", 
            fontWeight: "600" 
          }}
        >
          Clear all filters
        </span>
      ) : (
        <Button 
          appearance="primary" 
          size="small" 
          onClick={onAddLocation}
          style={{ marginTop: "4px" }}
        >
          Add a Location
        </Button>
      )}
    </div>
  );
};

/**
 * Formats a path based on the user's preferred path type setting.
 * UNC: \\server\share format for network paths
 * Drive: C:\folder format for local/mapped drive paths
 */
function formatPathByType(rawPath, pathType) {
  if (!rawPath) return "";
  const normalized = String(rawPath);
  // Best-effort conversion for local/admin-share paths. We avoid guessing custom mappings.
  if (pathType === "UNC") {
    const driveMatch = normalized.match(/^([a-zA-Z]):[\\/](.*)$/);
    if (driveMatch) {
      const drive = driveMatch[1].toUpperCase();
      const rest = driveMatch[2].replace(/\//g, "\\");
      return `\\\\localhost\\${drive}$\\${rest}`;
    }
    return normalized;
  }

  // Convert localhost admin shares back to drive format when possible.
  const uncToDrive = normalized.match(/^\\\\localhost\\([a-zA-Z])\$\\(.*)$/i);
  if (uncToDrive) {
    const drive = uncToDrive[1].toUpperCase();
    const rest = uncToDrive[2].replace(/\//g, "\\");
    return `${drive}:\\${rest}`;
  }

  return rawPath;
}

const LocationTable = ({ locations, selectedIds, onSelectionChange, connectivityStatus, onToggleSuggestion, onDoubleClickLocation, onAddLocation }) => {
  const [filterText, setFilterText] = React.useState("");
  const [columnFilter, setColumnFilter] = React.useState("All columns");
  const [locationFilter, setLocationFilter] = React.useState("All locations");
  const [pathType, setPathType] = React.useState("Drive");
  const [includeCollectionName, setIncludeCollectionName] = React.useState(false);
  const [focusedId, setFocusedId] = React.useState(null);

  const [isNarrow, setIsNarrow] = React.useState(() => typeof window !== "undefined" ? window.innerWidth < 550 : false);

  React.useEffect(() => {
    const handleResize = () => {
      setIsNarrow(window.innerWidth < 550);
    };
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  // Listen for options changes to update pathType display
  React.useEffect(() => {
    const loadPathType = () => {
      try {
        const stored = localStorage.getItem('koyomail_options');
        if (stored) {
          const parsed = JSON.parse(stored);
          setPathType(parsed.pathType || "Drive");
          setIncludeCollectionName(!!parsed.includeCollectionName);
        }
      } catch (e) {
        console.error(e);
      }
    };
    loadPathType();

    const handleStorageChange = (e) => {
      if (e.key === "koyomail_options") {
        loadPathType();
      }
    };

    window.addEventListener('koyomail_options_updated', loadPathType);
    window.addEventListener('storage', handleStorageChange);
    return () => {
      window.removeEventListener('koyomail_options_updated', loadPathType);
      window.removeEventListener('storage', handleStorageChange);
    };
  }, []);

  const filtered = locations.filter((item) => {
    const text = filterText.toLowerCase();
    const desc = String(item.description || "").toLowerCase();
    const path = String(item.path || "").toLowerCase();
    const coll = String(item.collection || "").toLowerCase();

    // 1. Column-specific Text Filter
    let matchesText = true;
    if (text) {
      switch (columnFilter) {
        case "Description":
          matchesText = desc.includes(text);
          break;
        case "Collection":
          matchesText = coll.includes(text);
          break;
        case "Location":
          matchesText = path.includes(text);
          break;
        default: // "All columns"
          matchesText = desc.includes(text) || path.includes(text) || coll.includes(text);
          break;
      }
    }

    // 2. Category Filter
    let matchesCategory = true;
    if (locationFilter === "Suggested") {
      matchesCategory = item.isSuggested;
    } else if (locationFilter === "Private") {
      matchesCategory = item.collection === "Private" || item.collection === "Personal";
    }

    return matchesText && matchesCategory;
  });

  const handleClearFilters = () => {
    setFilterText("");
    setColumnFilter("All columns");
    setLocationFilter("All locations");
  };

  const handleKeyDown = (e, index) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      const nextRow = document.querySelector(`[data-row-index="${index + 1}"]`);
      if (nextRow) nextRow.focus();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      const prevRow = document.querySelector(`[data-row-index="${index - 1}"]`);
      if (prevRow) prevRow.focus();
    } else if (e.key === " ") {
      e.preventDefault();
      onSelectionChange(filtered[index].id);
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (onDoubleClickLocation) {
        onDoubleClickLocation(filtered[index].path);
      }
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* Filter Bar */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, padding: "8px", borderBottom: "1px solid #edebe9", backgroundColor: "#fff" }}>
        <Input
          size="small"
          placeholder="Filter locations"
          contentBefore={<Search16Regular />}
          value={filterText}
          onChange={(e) => setFilterText(e.target.value)}
          style={{ flex: "1 1 150px", minWidth: 120 }}
        />
        <Select size="small" value={columnFilter} onChange={(e) => setColumnFilter(e.target.value)} style={{ flex: "1 1 100px", minWidth: 90 }}>
          <option>All columns</option>
          <option>Description</option>
          <option>Collection</option>
          <option>Location</option>
        </Select>
        <Select size="small" value={locationFilter} onChange={(e) => setLocationFilter(e.target.value)} style={{ flex: "1 1 100px", minWidth: 90 }}>
          <option>All locations</option>
          <option>Suggested</option>
          <option>Private</option>
        </Select>
      </div>

      {/* Table / List Container */}
      <div style={{ overflowY: "auto", overflowX: isNarrow ? "hidden" : "auto", flexGrow: 1 }}>
        {isNarrow ? (
          <div style={{ display: "flex", flexDirection: "column" }}>
            {filtered.map((item, index) => (
              <div 
                key={item.id}
                data-row-index={index}
                tabIndex={0}
                onFocus={() => setFocusedId(item.id)}
                onBlur={() => setFocusedId(null)}
                onKeyDown={(e) => handleKeyDown(e, index)}
                onClick={() => onSelectionChange(item.id)}
                onDoubleClick={() => onDoubleClickLocation && onDoubleClickLocation(item.path)}
                style={{
                  display: "flex",
                  flexDirection: "column",
                  padding: "10px 12px",
                  borderBottom: "1px solid #edebe9",
                  backgroundColor: selectedIds.includes(item.id) ? "#f3f2f1" : focusedId === item.id ? "#faf9f8" : "#fff",
                  outline: focusedId === item.id ? "2px solid #0078d4" : "none",
                  outlineOffset: "-2px",
                  cursor: "pointer",
                  userSelect: "none",
                  position: "relative",
                  transition: "background-color 0.15s ease",
                  opacity: item.isUnused ? 0.7 : 1
                }}
              >
                {/* Row 1: Checkbox, Description, Icons */}
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1, minWidth: 0 }}>
                    <Checkbox
                      size="small"
                      checked={selectedIds.includes(item.id)}
                      onChange={(e) => {
                        if (e && e.stopPropagation) e.stopPropagation();
                      }}
                      style={{ pointerEvents: "none" }}
                    />
                    <span style={{ 
                      fontWeight: "600", 
                      fontSize: "13px", 
                      color: item.isUnused ? "#a4262c" : "#323130",
                      textDecoration: item.isUnused ? "line-through" : "none",
                      whiteSpace: "nowrap",
                      overflow: "hidden",
                      textOverflow: "ellipsis"
                    }}>
                      {item.description}
                    </span>
                  </div>
                  
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    {connectivityStatus[item.id] && (
                      <Checkmark16Regular style={{ color: "#107c10" }} title="Connected" />
                    )}
                    <div onClick={(e) => { e.stopPropagation(); onToggleSuggestion(item.id); }} style={{ cursor: "pointer", display: "flex", alignItems: "center" }}>
                      {item.isSuggested ? (
                        <Star16Filled 
                          style={{ color: "#ffb900" }} 
                          title={
                            item.isSenderSuggested
                              ? (item.originalSuggested ? "Suggested for this sender & marked as favourite" : "Suggested for this sender")
                              : "Favourite location"
                          } 
                        />
                      ) : (
                        <Star16Regular style={{ color: "#c8c6c4" }} title="Set as favourite" />
                      )}
                    </div>
                  </div>
                </div>

                {/* Row 2: Collection & Path */}
                <div style={{ 
                  display: "flex", 
                  alignItems: "center", 
                  gap: 8, 
                  paddingLeft: 32, 
                  marginTop: 4, 
                  fontSize: "11px", 
                  color: "#605e5c",
                  textDecoration: item.isUnused ? "line-through" : "none"
                }}>
                  <span style={{ 
                    backgroundColor: "#edebe9", 
                    padding: "1px 6px", 
                    borderRadius: 4, 
                    fontSize: "10px",
                    fontWeight: "600",
                    color: "#323130",
                    flexShrink: 0
                  }}>
                    {item.collection}
                  </span>
                  <span style={{ 
                    whiteSpace: "nowrap", 
                    overflow: "hidden", 
                    textOverflow: "ellipsis", 
                    flex: 1 
                  }}>
                    {includeCollectionName && item.collection && (
                      <span style={{ fontWeight: "600", marginRight: "6px", color: "#323130" }}>[{item.collection}]</span>
                    )}
                    {formatPathByType(item.path, pathType)}
                  </span>
                </div>
              </div>
            ))}
            {filtered.length === 0 && (
              <EmptyState 
                isSearchFilter={locations.length > 0} 
                onClearFilters={handleClearFilters} 
                onAddLocation={onAddLocation} 
              />
            )}
          </div>
        ) : (
          <Table size="extra-small" style={{ minWidth: 600 }}>
            <TableHeader>
              <TableRow>
                <TableHeaderCell style={{ width: 24 }}></TableHeaderCell>
                <TableHeaderCell style={{ width: 40 }}>Online</TableHeaderCell>
                <TableHeaderCell style={{ width: 40 }}>Favorites</TableHeaderCell>
                <TableHeaderCell style={{ width: 80 }}>Collection</TableHeaderCell>
                <TableHeaderCell style={{ width: 200 }}>Description</TableHeaderCell>
                <TableHeaderCell style={{ minWidth: 200, width: "100%" }}>Location</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {filtered.map((item, index) => (
                <TableRow 
                  key={item.id} 
                  data-row-index={index}
                  tabIndex={0}
                  onFocus={() => setFocusedId(item.id)}
                  onBlur={() => setFocusedId(null)}
                  onKeyDown={(e) => handleKeyDown(e, index)}
                  selected={selectedIds.includes(item.id)}
                  onDoubleClick={() => onDoubleClickLocation && onDoubleClickLocation(item.path)}
                  onClick={() => onSelectionChange(item.id)}
                  style={{ 
                    cursor: "pointer", 
                    color: item.isUnused ? "#a4262c" : "inherit",
                    textDecoration: item.isUnused ? "line-through" : "none",
                    opacity: item.isUnused ? 0.7 : 1,
                    backgroundColor: selectedIds.includes(item.id) ? "#f3f2f1" : focusedId === item.id ? "#faf9f8" : "transparent",
                    outline: focusedId === item.id ? "2px solid #0078d4" : "none",
                    outlineOffset: "-2px"
                  }}
                >
                  <TableCell style={{ width: 24 }}>
                    <Checkbox
                      size="small"
                      checked={selectedIds.includes(item.id)}
                      onChange={(e) => {
                        if (e && e.stopPropagation) e.stopPropagation();
                      }}
                      style={{ pointerEvents: "none" }}
                    />
                  </TableCell>
                  <TableCell style={{ width: 40 }}>
                    {connectivityStatus[item.id] && (
                      <Checkmark16Regular style={{ color: "#107c10" }} title="Connected" />
                    )}
                  </TableCell>
                  <TableCell style={{ width: 40 }}>
                    <div onClick={(e) => { e.stopPropagation(); onToggleSuggestion(item.id); }} style={{ cursor: "pointer" }}>
                      {item.isSuggested ? (
                        <Star16Filled 
                          style={{ color: "#ffb900" }} 
                          title={
                            item.isSenderSuggested
                              ? (item.originalSuggested ? "Suggested for this sender & marked as favourite" : "Suggested for this sender")
                              : "Favourite location"
                          } 
                        />
                      ) : (
                        <Star16Regular style={{ color: "#c8c6c4" }} title="Set as favourite" />
                      )}
                    </div>
                  </TableCell>
                  <TableCell style={{ width: 80, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {item.collection}
                  </TableCell>
                  <TableCell style={{ width: 200, overflow: "hidden" }}>
                    <TableCellLayout weight="semibold" style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", ...(item.isUnused ? { textDecoration: "line-through", color: "#a4262c" } : {}) }}>
                      {item.description}
                    </TableCellLayout>
                  </TableCell>
                  <TableCell style={{ minWidth: 200, width: "100%", overflow: "hidden" }}>
                    <TableCellLayout size="small" style={{ color: "#605e5c", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", ...(item.isUnused ? { textDecoration: "line-through" } : {}) }}>
                      {includeCollectionName && item.collection && (
                        <span style={{ fontWeight: "600", marginRight: "6px", color: "#323130" }}>[{item.collection}]</span>
                      )}
                      {formatPathByType(item.path, pathType)}
                    </TableCellLayout>
                  </TableCell>
                </TableRow>
              ))}
              {filtered.length === 0 && (
                <TableRow style={{ backgroundColor: "transparent" }}>
                  <TableCell colSpan={6} style={{ padding: 24, border: "none" }}>
                    <EmptyState 
                      isSearchFilter={locations.length > 0} 
                      onClearFilters={handleClearFilters} 
                      onAddLocation={onAddLocation} 
                    />
                  </TableCell>
                </TableRow>
              )}
            </TableBody>
          </Table>
        )}
      </div>
    </div>
  );
};

export default LocationTable;
