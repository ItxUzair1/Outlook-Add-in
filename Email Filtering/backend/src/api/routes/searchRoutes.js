import { Router } from "express";
import { getSearchIndex, saveSearchIndex, getLocations } from "../../storage/repositories.js";
import { exec } from "child_process";
import fs from "fs/promises";
import path from "path";
import os from "os";

const router = Router();

/**
 * GET /api/search?dateRange=&from=&to=&cc=&subject=&body=&hasAttachments=&location=&keywords=&resultKind=&searchScope=
 * resultKind: all (default) | files — files = index row whose filePath is not .eml/.msg (e.g. saved attachments).
 * searchScope: locations_i_use (default) | all_locations — restricts results to user's configured locations or searches all.
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
      searchScope, // "locations_i_use" | "all_locations"
    } = req.query;

    let results = [...index];

    // ── Search scope filter (locations I use vs all) ────────────────────────
    if (!searchScope || searchScope === "locations_i_use") {
      // Only include results whose filePath matches one of the user's configured locations.
      // If there are no configured locations, return no results for this scope.
      const locations = await getLocations();
      if (locations.length === 0) {
        results = [];
      } else {
        const locationPaths = locations.map(loc => (loc.path || "").toLowerCase().replace(/\\/g, "/"));
        results = results.filter(r => {
          const fp = (r.filePath || "").toLowerCase().replace(/\\/g, "/");
          return locationPaths.some(lp => lp && fp.startsWith(lp));
        });
      }
    }
    // When searchScope === "all_locations", no location-based filtering is applied

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
 * Opens a modern Windows File Explorer-style folder picker using the
 * IFileOpenDialog COM interface (same dialog engine as File Explorer).
 * Shows all drives including mapped network drives.
 */
router.get("/browse-folder", async (req, res, next) => {
  const timestamp = Date.now();
  const ps1Path = path.join(os.tmpdir(), `koyobrowse_${timestamp}.ps1`);

  // PowerShell script that uses C# COM interop to show the modern folder picker.
  // IFileOpenDialog with FOS_PICKFOLDERS gives the full File Explorer dialog
  // including mapped network drives, Quick Access, search, etc.
  const psScript = `Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IShellItem {
    void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
    void GetParent(out IShellItem ppsi);
    void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    int Compare(IShellItem psi, uint hint, out int piOrder);
}

[ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IFileOpenDialog {
    [PreserveSig] int Show(IntPtr hwndOwner);
    void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
    void SetFileTypeIndex(uint iFileType);
    void GetFileTypeIndex(out uint piFileType);
    void Advise(IntPtr pfde, out uint pdwCookie);
    void Unadvise(uint dwCookie);
    void SetOptions(uint fos);
    void GetOptions(out uint pfos);
    void SetDefaultFolder(IShellItem psi);
    void SetFolder(IShellItem psi);
    void GetFolder(out IShellItem ppsi);
    void GetCurrentSelection(out IShellItem ppsi);
    void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
    void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
    void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
    void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    void GetResult(out IShellItem ppsi);
    void AddPlace(IShellItem psi, int fdap);
    void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
    void Close(int hr);
    void SetClientGuid(ref Guid guid);
    void ClearClientData();
    void SetFilter(IntPtr pFilter);
    void GetResults(out IntPtr ppenum);
    void GetSelectedItems(out IntPtr ppsai);
}

public class FolderPicker {
    [DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    public static string Pick() {
        // Send Alt key to allow this process to set foreground window
        keybd_event(0x12, 0, 0, UIntPtr.Zero);
        keybd_event(0x12, 0, 2, UIntPtr.Zero);

        var clsid = new Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7");
        Type t = Type.GetTypeFromCLSID(clsid);
        IFileOpenDialog dialog = (IFileOpenDialog)Activator.CreateInstance(t);
        try {
            dialog.SetTitle("Select Destination Folder");
            dialog.SetOkButtonLabel("Select Folder");
            uint opts;
            dialog.GetOptions(out opts);
            // FOS_PICKFOLDERS (0x20) = folder selection mode
            // FOS_FORCEFILESYSTEM (0x40) = only file-system items
            dialog.SetOptions(opts | 0x20 | 0x40);
            int hr = dialog.Show(IntPtr.Zero);
            if (hr != 0) return string.Empty;
            IShellItem item;
            dialog.GetResult(out item);
            string folderPath;
            item.GetDisplayName(0x80058000, out folderPath); // SIGDN_FILESYSPATH
            return folderPath ?? string.Empty;
        } catch {
            return string.Empty;
        } finally {
            Marshal.FinalReleaseComObject(dialog);
        }
    }
}
'@

$result = [FolderPicker]::Pick()
if ($result) { Write-Output $result }
`;

  try {
    await fs.writeFile(ps1Path, psScript);

    exec(
      `powershell -sta -ExecutionPolicy Bypass -File "${ps1Path}"`,
      { timeout: 120000 },
      async (error, stdout, stderr) => {
        // Clean up temp file
        try { await fs.unlink(ps1Path); } catch (e) {}

        if (error && error.killed) {
          return res.status(500).json({ error: "Folder picker timed out" });
        }
        if (error && !stdout.trim()) {
          console.error(`[searchRoutes] Folder picker failed: ${stderr || error.message}`);
          return res.status(500).json({ error: "Failed to open folder picker", details: stderr || error.message });
        }
        const selectedPath = stdout.trim();
        res.json({ path: selectedPath }); // Will be empty if user cancelled
      }
    );
  } catch (err) {
    console.error(`[searchRoutes] Failed to create ps1 temp file: ${err.message}`);
    return res.status(500).json({ error: "Failed to open folder picker", details: err.message });
  }
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
