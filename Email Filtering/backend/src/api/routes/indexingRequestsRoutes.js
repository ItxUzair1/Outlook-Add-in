import express from "express";
import { createIndexingRequest } from "../../storage/indexingRequestsStore.js";

const router = express.Router();

router.post("/", async (req, res, next) => {
  try {
    const { projectName, userEmail } = req.body;
    if (!projectName || !userEmail) {
      return res.status(400).json({ error: "Missing projectName or userEmail" });
    }

    const id = await createIndexingRequest(projectName, userEmail);
    res.json({ success: true, id });
  } catch (error) {
    console.error("[indexingRequestsRoutes] Error creating request:", error);
    next(error);
  }
});

export default router;
