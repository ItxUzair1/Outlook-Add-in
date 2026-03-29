import { fileEmail } from "./backend/src/services/fileService.js";
import fs from "fs/promises";
import path from "path";

async function test() {
  const payload = {
    subject: "Test Attachment Filing",
    sender: "test@example.com",
    to: ["recipient@example.com"],
    sentAt: new Date().toISOString(),
    bodyPreview: "This is a test body for attachment filing. It contains a $ dollar sign to test PowerShell escaping.",
    attachments: [
      {
        name: "test_image.txt",
        base64Content: Buffer.from("Hello world from $ attachment").toString("base64"),
      }
    ],
    targetPaths: ["test-output"],
    msgStrategy: "outlook-com"
  };

  try {
    console.log("Starting filing test (Strategy: outlook-com)...");
    const result = await fileEmail(payload);
    console.log("Filing result status:", result.results[0].status);
    
    // The path is relative to fileStorageRoot in the config
    const outputDir = path.resolve("file-storage", "test-output");
    console.log("Checking output directory:", outputDir);
    const files = await fs.readdir(outputDir);
    console.log("Files in output dir:", files);
    
    const msgFile = files.find(f => f.endsWith(".msg"));
    if (msgFile) {
      const stats = await fs.stat(path.join(outputDir, msgFile));
      console.log(`Generated MSG file: ${msgFile}, Size: ${stats.size} bytes`);
    }

  } catch (error) {
    console.error("Test failed:", error);
  }
}

test();
