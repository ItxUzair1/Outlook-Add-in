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
} from "@fluentui/react-components";
import { Checkmark16Regular, Star16Filled, Star16Regular, Search16Regular } from "@fluentui/react-icons";

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

const LocationTable = ({ locations, selectedIds, onSelectionChange, connectivityStatus, onToggleSuggestion, onDoubleClickLocation }) => {
  const [filterText, setFilterText] = React.useState("");
  const [columnFilter, setColumnFilter] = React.useState("All columns");
  const [locationFilter, setLocationFilter] = React.useState("All locations");
  const [pathType, setPathType] = React.useState("Drive");
  const [includeCollectionName, setIncludeCollectionName] = React.useState(false);

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
    const desc = (item.description || "").toLowerCase();
    const path = (item.path || "").toLowerCase();
    const coll = (item.collection || "").toLowerCase();

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

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%", overflow: "hidden" }}>
      {/* Filter Bar */}
      <div style={{ display: "flex", gap: 6, padding: "4px", borderBottom: "1px solid #edebe9", backgroundColor: "#fff" }}>
        <Input
          size="small"
          placeholder="Filter locations"
          contentBefore={<Search16Regular />}
          value={filterText}
          onChange={(e) => setFilterText(e.target.value)}
          style={{ width: 160 }}
        />
        <Select size="small" value={columnFilter} onChange={(e) => setColumnFilter(e.target.value)} style={{ width: 110 }}>
          <option>All columns</option>
          <option>Description</option>
          <option>Collection</option>
          <option>Location</option>
        </Select>
        <Select size="small" value={locationFilter} onChange={(e) => setLocationFilter(e.target.value)} style={{ width: 110 }}>
          <option>All locations</option>
          <option>Suggested</option>
          <option>Private</option>
        </Select>
      </div>

      {/* Table */}
      <div style={{ overflowY: "auto", overflowX: "auto", flexGrow: 1 }}>
        <Table size="extra-small" style={{ minWidth: 600 }}>
          <TableHeader>
            <TableRow>
              <TableHeaderCell style={{ width: 24 }}></TableHeaderCell>
              <TableHeaderCell style={{ width: 40 }}>Online</TableHeaderCell>
              <TableHeaderCell style={{ width: 40 }}>Favorites</TableHeaderCell>
              <TableHeaderCell style={{ width: 80 }}>Collection</TableHeaderCell>
              <TableHeaderCell style={{ minWidth: 150 }}>Description</TableHeaderCell>
              <TableHeaderCell style={{ minWidth: 300 }}>Location</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {filtered.map((item) => (
              <TableRow 
                key={item.id} 
                selected={selectedIds.includes(item.id)}
                onDoubleClick={() => onDoubleClickLocation && onDoubleClickLocation(item.path)}
                onClick={() => onSelectionChange(item.id)}
                style={{ 
                  cursor: "pointer", 
                  color: item.isUnused ? "#a4262c" : "inherit",
                  textDecoration: item.isUnused ? "line-through" : "none",
                  opacity: item.isUnused ? 0.7 : 1 
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
                      <Star16Filled style={{ color: "#ffb900" }} title="Suggested" />
                    ) : (
                      <Star16Regular style={{ color: "#c8c6c4" }} />
                    )}
                  </div>
                </TableCell>
                <TableCell style={{ width: 80, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {item.collection}
                </TableCell>
                <TableCell style={{ minWidth: 150, overflow: "hidden" }}>
                  <TableCellLayout weight="semibold" style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", ...(item.isUnused ? { textDecoration: "line-through", color: "#a4262c" } : {}) }}>
                    {item.description}
                  </TableCellLayout>
                </TableCell>
                <TableCell style={{ minWidth: 300, overflow: "hidden" }}>
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
              <TableRow>
                <TableCell colSpan={6} style={{ color: "#605e5c", fontSize: 12, padding: 12 }}>No locations found.</TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </div>
    </div>
  );
};

export default LocationTable;
