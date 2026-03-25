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
} from "../../services/locationService.js";

const router = Router();

router.get("/", async (_req, res, next) => {
  try {
    res.json(await listLocations());
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
    if (!req.body.path) {
      return res.status(400).json({ message: "path is required" });
    }
    await exploreLocation(req.body.path);
    res.status(204).send();
  } catch (e) {
    next(e);
  }
});

router.post("/:id/remove-suggestion", async (req, res, next) => {
  try {
    const updated = await removeSuggestion(req.params.id);
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
    const updated = await toggleSuggestion(req.params.id);
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
