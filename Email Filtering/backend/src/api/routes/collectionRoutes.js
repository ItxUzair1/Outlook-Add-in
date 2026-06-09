import { Router } from "express";
import { loadCollectionFile, saveCollectionFile } from "../../services/collectionService.js";

const router = Router();

router.post("/load", async (req, res, next) => {
  try {
    const { filePath } = req.body;
    if (!filePath) {
      return res.status(400).json({ message: "filePath is required" });
    }

    const locations = await loadCollectionFile(filePath);
    return res.json({ locations });
  } catch (e) {
    return next(e);
  }
});

router.post("/save", async (req, res, next) => {
  try {
    const { filePath, locations } = req.body;
    if (!filePath || !Array.isArray(locations)) {
      return res.status(400).json({ message: "filePath and locations array are required" });
    }

    await saveCollectionFile(filePath, locations);
    return res.status(204).send();
  } catch (e) {
    return next(e);
  }
});

export default router;
