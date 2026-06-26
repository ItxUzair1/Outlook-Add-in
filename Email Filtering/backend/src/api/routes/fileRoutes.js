import { Router } from "express";
import { fileEmail, createConsolidatedDraft, applyPostFilingActions } from "../../services/fileService.js";

const router = Router();

router.post("/email", async (req, res, next) => {
  try {
    const { subject, targetPaths } = req.body || {};

    if (!subject) {
      return res.status(400).json({ message: "subject is required" });
    }

    if (!Array.isArray(targetPaths) || targetPaths.length === 0) {
      return res.status(400).json({ message: "targetPaths must be a non-empty array" });
    }

    const result = await fileEmail(req.body);
    return res.status(201).json(result);
  } catch (e) {
    return next(e);
  }
});

router.post("/draft", async (req, res, next) => {
  try {
    const result = await createConsolidatedDraft(req.body);
    return res.status(201).json(result);
  } catch (e) {
    return next(e);
  }
});

router.post("/post-filing", async (req, res, next) => {
  try {
    const { itemId } = req.body || {};
    if (!itemId) {
      return res.status(400).json({ message: "itemId is required" });
    }
    const result = await applyPostFilingActions(req.body);
    return res.status(200).json(result);
  } catch (e) {
    return next(e);
  }
});

export default router;
