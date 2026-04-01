import { Router } from "express";
import { fileEmail } from "../../services/fileService.js";

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



export default router;
