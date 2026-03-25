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

app.listen(config.port, () => {
  console.log(`Backend listening on port ${config.port}`);
});
