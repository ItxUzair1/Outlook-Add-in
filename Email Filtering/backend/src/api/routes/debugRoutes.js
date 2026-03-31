import { Router } from "express";

const router = Router();

/**
 * POST /api/debug/auth-log
 * Receives auth diagnostic logs from the frontend (authManager.js)
 * and prints them to the backend terminal — useful when DevTools
 * is not available (e.g. New Outlook).
 */
router.post("/auth-log", (req, res) => {
  const { level = "info", message, data } = req.body || {};

  const prefix = {
    info:  "ℹ️  [AUTH]",
    warn:  "⚠️  [AUTH]",
    error: "❌ [AUTH]",
    ok:    "✅ [AUTH]",
  }[level] ?? "   [AUTH]";

  const extras = data ? `\n         ${JSON.stringify(data, null, 2).replace(/\n/g, "\n         ")}` : "";
  console.log(`${prefix} ${message}${extras}`);

  res.status(204).send();
});

export default router;
