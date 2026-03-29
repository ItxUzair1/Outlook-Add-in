import path from "path";
import dotenv from "dotenv";
import os from "os";

dotenv.config();

function resolvePath(input, fallback) {
  const value = input || fallback;
  return path.isAbsolute(value) ? value : path.resolve(process.cwd(), value);
}

export const config = {
  port: Number(process.env.PORT || 4000),
  allowOrigins: (process.env.ALLOW_ORIGINS || "https://localhost:3000,http://localhost:3000")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean),
  dataDir: resolvePath(process.env.DATA_DIR, "./data"),
  fileStorageRoot: resolvePath(process.env.FILE_STORAGE_ROOT, "./file-storage"),
  msgStrategy: String(process.env.MSG_STRATEGY || (os.platform() === "win32" ? "outlook-com" : "pseudo")).trim(),
  strictMsgRequired: true,
  azureClientId: process.env.AZURE_CLIENT_ID,
  azureTenantId: process.env.AZURE_TENANT_ID,
  azureClientSecret: process.env.AZURE_CLIENT_SECRET,
  graphScopes: (process.env.GRAPH_SCOPES || "https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access").split(" "),
};
