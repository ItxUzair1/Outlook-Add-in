import { Router } from "express";
import { getSearchIndex, saveSearchIndex } from "../../storage/repositories.js";
import { exec } from "child_process";
import fs from "fs/promises";
import path from "path";

const router = Router();

/**
 * GET /api/search?dateRange=&from=&to=&cc=&subject=&body=&hasAttachments=&location=&keywords=&resultKind=
 * resultKind: all (default) | files — files = index row whose filePath is not .eml/.msg (e.g. saved attachments).
 * Searches the filed email index with optional filters.
 */
router.get("/", async (req, res, next) => {
  try {
    const index = await getSearchIndex();

    const {
      dateRange,   // e.g. "1m", "3m", "6m", "1y", "all"
      from,
      to,
      cc,
      subject,
      keywords,    // matches subject, sender, recipients, path, body; + comment if including=true
      location,    // filed location path keyword
      hasAttachments, // "true" / "false"
      body,        // search within indexed body
      resultKind,  // "all" | "files"
    } = req.query;

    let results = [...index];

    // ── Date range filter ────────────────────────────────────────────────────
    if (dateRange && dateRange !== "all") {
      const now = new Date();
      const cutoff = new Date(now);
      switch (dateRange) {
        case "1m":  cutoff.setMonth(now.getMonth() - 1); break;
        case "3m":  cutoff.setMonth(now.getMonth() - 3); break;
        case "6m":  cutoff.setMonth(now.getMonth() - 6); break;
        case "1y":  cutoff.setFullYear(now.getFullYear() - 1); break;
      }
      results = results.filter(r => r.sentAt && new Date(r.sentAt) >= cutoff);
    }

    // ── From filter ─────────────────────────────────────────────────────────
    if (from && from.trim()) {
      const q = from.trim().toLowerCase();
      results = results.filter(r =>
        (r.sender || "").toLowerCase().includes(q)
      );
    }

    // ── To filter ───────────────────────────────────────────────────────────
    if (to && to.trim()) {
      const q = to.trim().toLowerCase();
      results = results.filter(r =>
        Array.isArray(r.recipients)
          ? r.recipients.some(addr => addr.toLowerCase().includes(q))
          : (r.recipients || "").toLowerCase().includes(q)
      );
    }

    // ── CC filter ───────────────────────────────────────────────────────────
    if (cc && cc.trim()) {
      const q = cc.trim().toLowerCase();
      results = results.filter(r =>
        Array.isArray(r.cc)
          ? r.cc.some(addr => addr.toLowerCase().includes(q))
          : (r.cc || "").toLowerCase().includes(q)
      );
    }

    // ── Subject filter ───────────────────────────────────────────────────────
    if (subject && subject.trim()) {
      const q = subject.trim().toLowerCase();
      results = results.filter(r =>
        (r.subject || "").toLowerCase().includes(q)
      );
    }

    // ── Location / filed path filter ─────────────────────────────────────────
    if (location && location.trim()) {
      const q = location.trim().toLowerCase();
      results = results.filter(r =>
        (r.filePath || "").toLowerCase().includes(q)
      );
    }

    // ── Attachments filter ───────────────────────────────────────────────────
    if (hasAttachments === "true") {
      results = results.filter(r => r.hasAttachments === true);
    } else if (hasAttachments === "false") {
      results = results.filter(r => !r.hasAttachments);
    }

    // ── Result kind: email message vs other filed file ───────────────────────
    if (resultKind === "files") {
      results = results.filter((r) => {
        const fp = (r.filePath || "").toLowerCase();
        return fp && !fp.endsWith(".eml") && !fp.endsWith(".msg");
      });
    }

    // ── Body filter ──────────────────────────────────────────────────────────
    if (body && body.trim()) {
      const q = body.trim().toLowerCase();
      results = results.filter(r =>
        (r.body || "").toLowerCase().includes(q)
      );
    }

    // ── General keywords (subject, sender, recipients, path, body; + comment if including) ───
    if (keywords && keywords.trim()) {
      const q = keywords.trim().toLowerCase();
      const includingValue = req.query.including === "true";

      results = results.filter(r => {
        const recipients = Array.isArray(r.recipients) ? r.recipients.join(" ") : (r.recipients || "");
        let match = 
          (r.subject || "").toLowerCase().includes(q) ||
          (r.sender || "").toLowerCase().includes(q) ||
          recipients.toLowerCase().includes(q) ||
          (r.filePath || "").toLowerCase().includes(q) ||
          (r.body || "").toLowerCase().includes(q);
        
        // If "including" is on, also match the comment field
        if (includingValue && (r.comment || "").toLowerCase().includes(q)) {
            match = true;
        }
        
        return match;
      });
    }

    // Sort by sentAt descending
    results.sort((a, b) => new Date(b.sentAt || b.filedAt || 0) - new Date(a.sentAt || a.filedAt || 0));

    res.json({ count: results.length, results });
  } catch (e) {
    next(e);
  }
});

/**
 * GET /api/search/browse-folder
 * Opens a native Windows Folder Picker dialog via PowerShell and returns the selected path.
 */
router.get("/browse-folder", (req, res, next) => {
  const psScript = `
Add-Type -AssemblyName System.windows.forms;
$f = New-Object System.Windows.Forms.FolderBrowserDialog;
$f.Description = 'Select Destination Folder';
$f.ShowNewFolderButton = $true;
if($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
    Write-Output $f.SelectedPath 
}
  `;

  const encoded = Buffer.from(psScript, "utf16le").toString("base64");
  
  exec(`powershell -Sta -NoProfile -EncodedCommand ${encoded}`, (error, stdout, stderr) => {
    if (error) {
      console.error(`[searchRoutes] Folder picker failed: ${error.message}`);
      return res.status(500).json({ error: "Failed to open folder picker", details: error.message });
    }
    const selectedPath = stdout.trim();
    res.json({ path: selectedPath }); // Will be empty if user cancelled
  });
});

/**
 * POST /api/search/open
 * Opens the file in its default OS application (Outlook).
 */
router.post("/open", async (req, res, next) => {
  try {
    const { filePath } = req.body;
    if (!filePath) return res.status(400).json({ error: "filePath is required" });

    // Verify file exists first
    try {
      await fs.access(filePath);
    } catch (err) {
      console.warn(`[searchRoutes] Open attempt failed: File not found at ${filePath}`);
      return res.status(404).json({ error: "File not found at original location", code: "ENOENT" });
    }

    // Use 'start' command on Windows to launch default app
    // Surround with double quotes to handle spaces in paths
    exec(`start "" "${filePath}"`, (error) => {
      if (error) {
          console.error(`[searchRoutes] Failed to open file: ${error.message}`);
          return res.status(500).json({ error: `Could not open file: ${error.message}` });
      }
      res.json({ status: "success" });
    });
  } catch (e) {
    next(e);
  }
});

/**
 * POST /api/search/open-folder
 * Opens the containing folder of the search result.
 */
router.post("/open-folder", async (req, res, next) => {
  try {
    const { filePath } = req.body;
    if (!filePath) return res.status(400).json({ error: "filePath is required" });

    // Extract directory from file path
    const dirPath = path.dirname(filePath);

    // Verify directory exists first
    try {
      await fs.access(dirPath);
    } catch (err) {
      console.warn(`[searchRoutes] Open Folder attempt failed: Directory not found at ${dirPath}`);
      return res.status(404).json({ error: "Folder not found at original location", code: "ENOENT" });
    }

    // Use 'start' command to launch Explorer at that directory
    exec(`start "" "${dirPath}"`, (error) => {
      if (error) {
          console.error(`[searchRoutes] Failed to open folder: ${error.message}`);
          return res.status(500).json({ error: `Could not open folder: ${error.message}` });
      }
      res.json({ status: "success" });
    });
  } catch (e) {
    next(e);
  }
});

/**
 * DELETE /api/search/:id
 * Deletes the physical file and removes its entry from the index.
 */
router.delete("/:id", async (req, res, next) => {
  try {
    const { id } = req.params;
    const index = await getSearchIndex();
    const itemIdx = index.findIndex(x => x.id === id);

    if (itemIdx === -1) {
        return res.status(404).json({ error: "Item not found in index" });
    }

    const item = index[itemIdx];

    // Delete the file from the filesystem
    try {
      if (item.filePath) {
          await fs.unlink(item.filePath);
      }
    } catch (err) {
      console.warn(`[searchRoutes] Failed to delete physical file: ${err.message}`);
      // Continue even if physical file is missing, to keep index clean
    }

    // Update the search index
    const updatedIndex = index.filter(x => x.id !== id);
    await saveSearchIndex(updatedIndex);

    res.json({ status: "deleted" });
  } catch (e) {
    next(e);
  }
});

/**
 * POST /api/search/move
 * Moves the physical file to a new destination directory and updates the search index.
 */
router.post("/move", async (req, res, next) => {
  try {
    const { id, destinationDir } = req.body;
    if (!id || !destinationDir) {
      return res.status(400).json({ error: "id and destinationDir are required" });
    }

    const index = await getSearchIndex();
    const itemIdx = index.findIndex(x => x.id === id);

    if (itemIdx === -1) {
      return res.status(404).json({ error: "Item not found in index" });
    }

    const item = index[itemIdx];
    if (!item.filePath) {
      return res.status(400).json({ error: "Item does not have a physical file path" });
    }

    // Verify source exists
    try {
      await fs.access(item.filePath);
    } catch {
      return res.status(404).json({ error: "Original file not found on disk" });
    }

    // Verify destination directory exists
    try {
      const destStat = await fs.stat(destinationDir);
      if (!destStat.isDirectory()) {
         return res.status(400).json({ error: "Destination must be a directory" });
      }
    } catch {
      return res.status(400).json({ error: "Destination directory does not exist or is inaccessible" });
    }

    // Move file
    const fileName = path.basename(item.filePath);
    const newFilePath = path.join(destinationDir, fileName);

    // Prevent overwriting
    try {
      await fs.access(newFilePath);
      return res.status(400).json({ error: "A file with that name already exists at the destination" });
    } catch {
      // Good, file doesn't exist
    }

    try {
      await fs.rename(item.filePath, newFilePath);
    } catch (renameErr) {
      if (renameErr.code === 'EXDEV') {
        // Cross-device link not permitted, use copy + unlink fallback
        await fs.copyFile(item.filePath, newFilePath);
        await fs.unlink(item.filePath);
      } else {
        throw renameErr;
      }
    }

    // Update index
    item.filePath = newFilePath;
    item.parentFolder = destinationDir; // Update parent folder attribute if tracked
    await saveSearchIndex(index);

    res.json({ status: "moved", newPath: newFilePath });
  } catch (e) {
    next(e);
  }
});

/**
 * POST /api/search/sync
 * Scans the entire index and removes entries where the physical file is missing.
 */
router.post("/sync", async (req, res, next) => {
  try {
    const index = await getSearchIndex();
    
    // Check existence for all files in parallel (with some concurrency control implicitly via Promise.all)
    const checks = await Promise.all(index.map(async (item) => {
      try {
        if (!item.filePath) return { id: item.id, exists: false };
        await fs.access(item.filePath);
        return { id: item.id, exists: true };
      } catch (err) {
        return { id: item.id, exists: false };
      }
    }));

    const missingIds = new Set(checks.filter(c => !c.exists).map(c => c.id));
    
    // Prune missing and then De-duplicate based on internetMessageId and filePath
    const seen = new Set();
    const updatedIndex = [];
    
    for (const item of index) {
      if (missingIds.has(item.id)) continue;
      
      const key = `${item.internetMessageId}-${item.filePath}`;
      if (item.internetMessageId && seen.has(key)) continue;
      
      if (item.internetMessageId) seen.add(key);
      updatedIndex.push(item);
    }

    if (updatedIndex.length !== index.length) {
      await saveSearchIndex(updatedIndex);
    }

    res.json({ 
      status: "synced", 
      removedCount: index.length - updatedIndex.length,
      totalCount: updatedIndex.length 
    });
  } catch (e) {
    next(e);
  }
});

export default router;
