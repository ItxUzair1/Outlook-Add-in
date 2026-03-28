import { fileEmail } from "../c:/Users/muham/Desktop/Email Filtering/Email Filtering/backend/src/services/fileService.js";
import fs from "fs/promises";
import path from "path";

async function test() {
  const payload = {
    subject: "Test Attachment Filing",
    sender: "test@example.com",
    to: ["recipient@example.com"],
    sentAt: new Date().toISOString(),
    bodyPreview: "This is a test body for attachment filing.",
    attachments: [
      {
        name: "test_image.txt",
        base64Content: Buffer.from("Hello world from attachment").toString("base64"),
      }
    ],
    targetPaths: ["./test-output"]
  };

  try {
    console.log("Starting filing test...");
    const result = await fileEmail(payload);
    console.log("Filing result:", JSON.stringify(result, null, 2));
    
    const outputDir = path.resolve("test-output");
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
