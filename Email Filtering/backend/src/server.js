import dns from "dns";
dns.setDefaultResultOrder("ipv4first");
process.env.UV_THREADPOOL_SIZE = 64;
import express from "express";
import cors from "cors";
import helmet from "helmet";
import morgan from "morgan";
import https from "https";
import path from "path";
import fs from "fs";
import os from "os";
import devCerts from "office-addin-dev-certs";
import fetch from "node-fetch";

import { config } from "./config/index.js";
import healthRoutes from "./api/routes/healthRoutes.js";
import locationRoutes from "./api/routes/locationRoutes.js";
import fileRoutes from "./api/routes/fileRoutes.js";
import searchRoutes from "./api/routes/searchRoutes.js";
import preferencesRoutes from "./api/routes/preferencesRoutes.js";
import debugRoutes from "./api/routes/debugRoutes.js";
import collectionRoutes from "./api/routes/collectionRoutes.js";
import indexingRequestsRoutes from "./api/routes/indexingRequestsRoutes.js";
import { warmupAnalyticsConnection } from "./storage/analyticsStore.js";

const app = express();

app.use(helmet());
app.use(
  cors({
    origin(origin, cb) {
      // In Agent mode, allow any origin since requests are authenticated by API token
      if (!origin || config.allowOrigins.includes(origin) || process.env.AGENT_MODE === "true") {
        cb(null, true);
        return;
      }

      console.warn(`[CORS] Rejected request from origin: ${origin}`);
      cb(new Error("Origin not allowed by CORS"));
    },
  })
);
// 60 MB limit: a 21 MB combined attachment set becomes ~28 MB when base64-encoded
// in the JSON filing payload. The previous 25 MB cap was silently rejecting large emails.
app.use(express.json({ limit: "60mb" }));
app.use(morgan("dev"));

app.get("/", (_req, res) => {
  res.json({ service: "email-filing-backend", status: "running" });
});

// ─── Agent security middleware (only active when AGENT_MODE=true) ────────────
// When running as the on-site corporate server agent, every request must carry
// the x-koyomail-token header matching AGENT_API_TOKEN in .env.
// Health and root endpoints are always public so the add-in can probe connectivity.
if (process.env.AGENT_MODE === "true" && process.env.AGENT_API_TOKEN) {
  const AGENT_TOKEN = process.env.AGENT_API_TOKEN;
  const PUBLIC_PATHS = ["/", "/api/health"];
  app.use((req, res, next) => {
    if (PUBLIC_PATHS.some((p) => req.path === p || req.path.startsWith("/api/health"))) {
      return next();
    }
    const token = req.headers["x-koyomail-token"] || req.query._token;
    if (!token || token !== AGENT_TOKEN) {
      return res.status(401).json({ error: "Unauthorized — invalid or missing agent token" });
    }
    next();
  });
  console.log("[server] Agent mode active — API token authentication enabled");
}

app.use("/api/health", healthRoutes);
app.use("/api/locations", locationRoutes);
app.use("/api/file", fileRoutes);
app.use("/api/search", searchRoutes);
app.use("/api/preferences", preferencesRoutes);
app.use("/api/debug", debugRoutes);
app.use("/api/collections", collectionRoutes);
app.use("/api/indexing-requests", indexingRequestsRoutes);

app.use((error, _req, res, _next) => {
  const status = error.status || 500;
  res.status(status).json({
    message: error.message || "Internal server error",
  });
});

// Validate critical configuration on startup
function validateStartupConfig() {
  const errors = [];
  const warnings = [];

  // Check Azure SSO/Graph credentials
  if (!config.azureClientId) {
    warnings.push("AZURE_CLIENT_ID is missing from .env - SSO and Microsoft Graph features will be unavailable");
  }
  if (!config.azureClientSecret) {
    warnings.push("AZURE_CLIENT_SECRET is missing from .env - SSO and Microsoft Graph features will be unavailable");
  }
  if (!config.azureTenantId) {
    warnings.push("AZURE_TENANT_ID is not set in .env - using 'common' tenant endpoint (slower) instead of your specific tenant");
  }

  // Check file storage
  if (!config.fileStorageRoot) {
    errors.push("FILE_STORAGE_ROOT is missing from .env - file storage is not configured");
  }

  if (errors.length > 0) {
    console.error("\n❌ STARTUP CONFIGURATION ERRORS:");
    errors.forEach(err => console.error(`   • ${err}`));
    console.error("\nPlease fix these issues and restart the server.\n");
  }

  if (warnings.length > 0) {
    console.warn("\n  STARTUP CONFIGURATION WARNINGS:");
    warnings.forEach(warn => console.warn(`   • ${warn}`));
    console.warn("");
  }

  return errors.length === 0;
}

if (process.argv.includes('--install-certs-only')) {
  devCerts.getHttpsServerOptions().then(() => {
    console.log("Certificates successfully installed or already trusted.");
    process.exit(0);
  }).catch(err => {
    console.log("Certificates generated. (Internal installation skipped due to pkg environment).");
    process.exit(0);
  });
} else {
  // Validate config before starting server
  if (!validateStartupConfig()) {
    process.exit(1);
  }

  // Start HTTPS server using dev certs
  try {
    const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
    const options = {
      key: fs.readFileSync(path.join(certDir, "localhost.key")),
      cert: fs.readFileSync(path.join(certDir, "localhost.crt"))
    };
    
    const server = https.createServer(options, app);
    
    // Increase timeouts to 30 minutes to prevent long-running directory scans from failing
    server.timeout = 30 * 60 * 1000;
    server.keepAliveTimeout = 30 * 60 * 1000;

    server.listen(config.port, () => {
      console.log(`✓ Backend listening securely on HTTPS port ${config.port}`);
      console.log(`✓ Azure SSO: ${config.azureClientId ? "CONFIGURED" : "DISABLED"}`);
      console.log(`✓ File Storage: ${config.fileStorageRoot || "NOT CONFIGURED"}`);

      // Pre-warm MongoDB connection so the first search analytics write succeeds
      warmupAnalyticsConnection();

    });
  } catch (err) {
    console.error("Failed to start HTTPS server (missing or invalid certificates):", err);
    process.exit(1);
  }
}
