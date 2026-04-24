import { Router } from "express";
import path from "path";
import { config } from "../../config/index.js";
import { readJson, writeJson } from "../../storage/jsonStore.js";

const router = Router();
const prefsPath = path.join(config.dataDir, "preferences.json");

const DEFAULT_PREFS = {
  enableSearching: true,
  searchScope: "locations_i_use",
  disableDelete: false,
  disableMoveTo: false,
  discoverLocations: false,
  applyReadOnly: false,
  pathType: "UNC",
  duplicateStrategy: "rename",
  defaultAttachments: "all",
};

/**
 * GET /api/preferences
 * Returns current user preferences (merged with defaults for any missing keys).
 */
router.get("/", async (_req, res, next) => {
  try {
    const stored = await readJson(prefsPath, {});
    res.json({ ...DEFAULT_PREFS, ...stored });
  } catch (e) {
    next(e);
  }
});

/**
 * PUT /api/preferences
 * Updates user preferences. Accepts partial updates (merges with existing).
 */
router.put("/", async (req, res, next) => {
  try {
    const stored = await readJson(prefsPath, {});
    const updated = { ...stored, ...req.body };
    await writeJson(prefsPath, updated);
    res.json({ ...DEFAULT_PREFS, ...updated });
  } catch (e) {
    next(e);
  }
});

export default router;
