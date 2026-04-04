import { Router } from "express";
import { getSearchIndex, saveSearchIndex } from "../../storage/repositories.js";
import { exec } from "child_process";
import fs from "fs/promises";

const router = Router();

/**
 * GET /api/search?dateRange=&from=&to=&cc=&subject=&body=&hasAttachments=&location=&keywords=
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
      keywords,    // general keyword search across subject/sender/recipients
      location,    // filed location path keyword
      hasAttachments, // "true" / "false"
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

    // ── General keywords filter (subject + sender + recipients + filePath) ───
    if (keywords && keywords.trim()) {
      const q = keywords.trim().toLowerCase();
      const includingValue = req.query.including === "true";

      results = results.filter(r => {
        const recipients = Array.isArray(r.recipients) ? r.recipients.join(" ") : (r.recipients || "");
        let match = 
          (r.subject || "").toLowerCase().includes(q) ||
          (r.sender || "").toLowerCase().includes(q) ||
          recipients.toLowerCase().includes(q) ||
          (r.filePath || "").toLowerCase().includes(q);
        
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
