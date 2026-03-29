import express from "express";
import cors from "cors";
import helmet from "helmet";
import morgan from "morgan";
import { config } from "./config/index.js";
import healthRoutes from "./api/routes/healthRoutes.js";
import locationRoutes from "./api/routes/locationRoutes.js";
import fileRoutes from "./api/routes/fileRoutes.js";
import searchRoutes from "./api/routes/searchRoutes.js";
import preferencesRoutes from "./api/routes/preferencesRoutes.js";

const app = express();

app.use(helmet());
app.use(
  cors({
    origin(origin, cb) {
      if (!origin || config.allowOrigins.includes(origin)) {
        cb(null, true);
        return;
      }

      cb(new Error("Origin not allowed by CORS"));
    },
  })
);
app.use(express.json({ limit: "25mb" }));
app.use(morgan("dev"));

app.get("/", (_req, res) => {
  res.json({ service: "email-filing-backend", status: "running" });
});

app.use("/api/health", healthRoutes);
app.use("/api/locations", locationRoutes);
app.use("/api/file", fileRoutes);
app.use("/api/search", searchRoutes);
app.use("/api/preferences", preferencesRoutes);

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
    errors.push("AZURE_CLIENT_ID is missing from .env - SSO and Microsoft Graph features will be unavailable");
  }
  if (!config.azureClientSecret) {
    errors.push("AZURE_CLIENT_SECRET is missing from .env - SSO and Microsoft Graph features will be unavailable");
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
    console.warn("\n⚠️  STARTUP CONFIGURATION WARNINGS:");
    warnings.forEach(warn => console.warn(`   • ${warn}`));
    console.warn("");
  }

  return errors.length === 0;
}

// Validate config before starting server
if (!validateStartupConfig()) {
  process.exit(1);
}

app.listen(config.port, () => {
  console.log(`✓ Backend listening on port ${config.port}`);
  console.log(`✓ Azure SSO: ${config.azureClientId ? "CONFIGURED" : "DISABLED"}`);
  console.log(`✓ File Storage: ${config.fileStorageRoot || "NOT CONFIGURED"}`);
});
