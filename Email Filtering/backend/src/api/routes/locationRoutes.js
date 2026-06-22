import { Router } from "express";
import {
  createLocation,
  listLocations,
  listSuggestedLocations,
  removeLocation,
  updateLocation,
  checkConnectivity,
  exploreLocation,
  removeSuggestion,
  toggleSuggestion,
  markUnused,
  discoverLocations,
  checkPathsConnectivity,
  getSenderHistoryStats,
  getSenderFavourites,
  getGeneralHistoryStats,
} from "../../services/locationService.js";

const router = Router();

router.get("/", async (req, res, next) => {
  try {
    const { sender } = req.query;
    res.json(await listLocations(sender));
  } catch (e) {
    next(e);
  }
});

router.get("/sender-history", async (req, res, next) => {
  try {
    const { sender } = req.query;
    const [history, favourites, generalHistory] = await Promise.all([
      getSenderHistoryStats(sender),
      getSenderFavourites(sender),
      getGeneralHistoryStats()
    ]);
    res.json({ history, favourites, generalHistory });
  } catch (e) {
    next(e);
  }
});

router.get("/suggested", async (_req, res, next) => {
  try {
    res.json(await listSuggestedLocations(10));
  } catch (e) {
    next(e);
  }
});

router.get("/status", async (_req, res, next) => {
  try {
    res.json(await checkConnectivity());
  } catch (e) {
    next(e);
  }
});

router.post("/status/check", async (req, res, next) => {
  try {
    if (!Array.isArray(req.body.paths)) {
      return res.status(400).json({ message: "paths must be an array" });
    }
    res.json(await checkPathsConnectivity(req.body.paths));
  } catch (e) {
    next(e);
  }
});

router.post("/", async (req, res, next) => {
  try {
    if (!req.body.path) {
      return res.status(400).json({ message: "path is required" });
    }

    const created = await createLocation(req.body);
    return res.status(201).json(created);
  } catch (e) {
    return next(e);
  }
});

/**
 * POST /api/locations/discover
 * Scans the search index for unique filing directories and auto-adds them as locations.
 */
router.post("/discover", async (_req, res, next) => {
  try {
    const result = await discoverLocations();
    res.json(result);
  } catch (e) {
    next(e);
  }
});

router.put("/:id", async (req, res, next) => {
  try {
    const updated = await updateLocation(req.params.id, req.body);
    if (!updated) {
      return res.status(404).json({ message: "Location not found" });
    }

    return res.json(updated);
  } catch (e) {
    return next(e);
  }
});

router.post("/explore", async (req, res, next) => {
  try {
    const pathToExplore = req.body.path || "";
    await exploreLocation(pathToExplore);
    res.status(204).send();
  } catch (e) {
    next(e);
  }
});

router.post("/:id/remove-suggestion", async (req, res, next) => {
  try {
    const { sender } = req.query;
    const updated = await removeSuggestion(req.params.id, sender);
    if (!updated) {
      return res.status(404).json({ message: "Location not found" });
    }
    return res.json(updated);
  } catch (e) {
    return next(e);
  }
});

router.post("/:id/toggle-suggestion", async (req, res, next) => {
  try {
    const { sender } = req.query;
    const updated = await toggleSuggestion(req.params.id, sender);
    if (!updated) {
      return res.status(404).json({ message: "Location not found" });
    }
    return res.json(updated);
  } catch (e) {
    return next(e);
  }
});

router.post("/:id/mark-unused", async (req, res, next) => {
  try {
    const updated = await markUnused(req.params.id);
    if (!updated) {
      return res.status(404).json({ message: "Location not found" });
    }
    return res.json(updated);
  } catch (e) {
    return next(e);
  }
});

router.delete("/:id", async (req, res, next) => {
  try {
    const ok = await removeLocation(req.params.id);
    if (!ok) {
      return res.status(404).json({ message: "Location not found" });
    }

    return res.status(204).send();
  } catch (e) {
    return next(e);
  }
});

export default router;
