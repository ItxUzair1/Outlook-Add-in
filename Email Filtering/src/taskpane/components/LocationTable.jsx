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

const LocationTable = ({ locations, selectedIds, onSelectionChange, connectivityStatus, onToggleSuggestion }) => {
  const [filterText, setFilterText] = React.useState("");
  const [columnFilter, setColumnFilter] = React.useState("All columns");
  const [locationFilter, setLocationFilter] = React.useState("All locations");

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
        case "Path":
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
    } else if (locationFilter === "Personal") {
      matchesCategory = item.collection === "Personal";
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
          <option>Path</option>
        </Select>
        <Select size="small" value={locationFilter} onChange={(e) => setLocationFilter(e.target.value)} style={{ width: 110 }}>
          <option>All locations</option>
          <option>Suggested</option>
          <option>Personal</option>
        </Select>
      </div>

      {/* Table */}
      <div style={{ overflowY: "auto", overflowX: "auto", flexGrow: 1 }}>
        <Table size="extra-small" style={{ minWidth: 600 }}>
          <TableHeader>
            <TableRow>
              <TableHeaderCell style={{ width: 24 }}></TableHeaderCell>
              <TableHeaderCell style={{ width: 40 }}>Status</TableHeaderCell>
              <TableHeaderCell style={{ width: 40 }}>Rank</TableHeaderCell>
              <TableHeaderCell style={{ width: 80 }}>Collection</TableHeaderCell>
              <TableHeaderCell style={{ minWidth: 150 }}>Description</TableHeaderCell>
              <TableHeaderCell style={{ minWidth: 300 }}>Path</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {filtered.map((item) => (
              <TableRow key={item.id} selected={selectedIds.includes(item.id)}>
                <TableCell>
                  <Checkbox
                    size="small"
                    checked={selectedIds.includes(item.id)}
                    onChange={() => onSelectionChange(item.id)}
                  />
                </TableCell>
                <TableCell>
                  {connectivityStatus[item.id] && (
                    <Checkmark16Regular style={{ color: "#107c10" }} title="Connected" />
                  )}
                </TableCell>
                <TableCell>
                  <div onClick={(e) => { e.stopPropagation(); onToggleSuggestion(item.id); }} style={{ cursor: "pointer" }}>
                    {item.isSuggested ? (
                      <Star16Filled style={{ color: "#ffb900" }} title="Suggested" />
                    ) : (
                      <Star16Regular style={{ color: "#c8c6c4" }} />
                    )}
                  </div>
                </TableCell>
                <TableCell>{item.collection}</TableCell>
                <TableCell>
                  <TableCellLayout weight="semibold">{item.description || item.path}</TableCellLayout>
                </TableCell>
                <TableCell>
                  <TableCellLayout size="small" style={{ color: "#605e5c" }}>{item.path}</TableCellLayout>
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
