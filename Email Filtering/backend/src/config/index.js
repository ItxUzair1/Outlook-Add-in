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
};
