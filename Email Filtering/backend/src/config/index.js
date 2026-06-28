import path from "path";
import dotenv from "dotenv";
import os from "os";
import fs from "fs";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Explicitly load from backend root
dotenv.config({ path: path.join(__dirname, "../../.env") });

function resolvePath(input, fallback) {
  const value = input || fallback;
  return path.isAbsolute(value) ? value : path.resolve(process.cwd(), value);
}

// Store user data in a persistent user profile directory to avoid deletion on app update
const defaultDataDir = path.join(os.homedir(), ".koyomail", "data");
const envDataDir = process.env.DATA_DIR;
// If DATA_DIR is not set, or is set to default relative path values "./data" or "data", use defaultDataDir
const targetDataDir = (!envDataDir || envDataDir === "./data" || envDataDir === "data") 
  ? defaultDataDir 
  : resolvePath(envDataDir, defaultDataDir);

// Perform one-time migration from old legacy directories if they exist
function migrateLegacyData(newDir) {
  const oldDirs = [
    path.resolve(process.cwd(), "./data"),
    path.resolve(process.cwd(), "../data"), // if process Cwd is backend, check root's data folder
    path.resolve(process.cwd(), "./backend/data") // if process Cwd is root, check backend's data folder
  ];

  for (const oldDir of oldDirs) {
    if (path.resolve(oldDir) === path.resolve(newDir)) {
      continue;
    }

    if (fs.existsSync(oldDir)) {
      try {
        if (!fs.existsSync(newDir)) {
          fs.mkdirSync(newDir, { recursive: true });
        }
        const files = fs.readdirSync(oldDir);
        for (const file of files) {
          const oldFile = path.join(oldDir, file);
          const newFile = path.join(newDir, file);
          if (fs.statSync(oldFile).isFile()) {
            if (!fs.existsSync(newFile)) {
              fs.copyFileSync(oldFile, newFile);
              console.log(`[Migration] Copied ${file} from legacy data directory ${oldDir} to ${newFile}`);
            }
          }
        }
      } catch (err) {
        console.error(`[Migration] Error migrating legacy data files from ${oldDir}:`, err);
      }
    }
  }
}

migrateLegacyData(targetDataDir);

export const config = {
  port: Number(process.env.PORT || 4000),
  allowOrigins: (process.env.ALLOW_ORIGINS || "https://localhost:3000,http://localhost:3000")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean),
  dataDir: targetDataDir,
  fileStorageRoot: resolvePath(process.env.FILE_STORAGE_ROOT, "./file-storage"),
  msgStrategy: String(process.env.MSG_STRATEGY || (os.platform() === "win32" ? "outlook-com" : "pseudo")).trim(),
  strictMsgRequired: true,
  azureClientId: process.env.AZURE_CLIENT_ID,
  azureTenantId: process.env.AZURE_TENANT_ID,
  azureClientSecret: process.env.AZURE_CLIENT_SECRET,
  // The scopes needed by the backend via OBO or direct Graph access
  graphScopes: (process.env.GRAPH_SCOPES || "https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access").split(" "),
};
