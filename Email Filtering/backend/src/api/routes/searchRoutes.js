import { Router } from "express";

const router = Router();

router.get("/", (_req, res) => {
  res.status(501).json({
    message: "Search is planned for a later milestone. Endpoint intentionally not implemented in Milestone 2.",
  });
});

export default router;
