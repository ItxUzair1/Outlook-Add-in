import fs from "fs/promises";
import path from "path";
import { listLocations, createLocation } from "../src/services/locationService.js";
import { fileEmail } from "../src/services/fileService.js";
import { config } from "../src/config/index.js";

async function ensureDryRunLocation() {
  const targetPath = path.join(config.fileStorageRoot, "_dry-run", "ProjectA", "Emails");
  const current = await listLocations();
  const existing = current.find((x) => x.path === targetPath);

  if (existing) {
    return existing;
  }

  return createLocation({
    type: "network",
    path: targetPath,
    description: "Dry Run Project A",
    collection: "Projects",
    isDefault: false,
  });
}

async function run() {
  console.log("[dry-run] starting");
  const location = await ensureDryRunLocation();

  const sampleAttachment = Buffer.from("Dry run attachment file", "utf-8").toString("base64");
  const sampleMime = Buffer.from(
    [
      "From: dryrun@example.com",
      "To: qa@example.com",
      "Subject: Dry run message",
      "Date: Tue, 25 Mar 2026 10:00:00 +0000",
      "MIME-Version: 1.0",
      'Content-Type: text/plain; charset="utf-8"',
      "",
      "This is a backend dry-run MIME payload.",
    ].join("\r\n"),
    "utf-8"
  ).toString("base64");

  const response = await fileEmail({
    internetMessageId: `dry-run-${Date.now()}`,
    subject: "Milestone2 Dry Run",
    sender: "dryrun@example.com",
    to: ["qa@example.com"],
    cc: [],
    sentAt: new Date().toISOString(),
    bodyPreview: "Milestone2 dry run payload",
    mimeBase64: sampleMime,
    msgStrategy: config.msgStrategy,
    duplicateStrategy: "rename",
    targetPaths: [location.path],
    attachments: [
      {
        id: "att-1",
        name: "dry-run.txt",
        base64Content: sampleAttachment,
      },
    ],
  });

  const first = response.results?.[0];
  if (!first) {
    throw new Error("No file result returned from filing operation.");
  }

  await fs.access(first.msgPath);
  if (first.attachments?.[0]) {
    await fs.access(first.attachments[0]);
  }

  console.log("[dry-run] success");
  console.log(`[dry-run] msg: ${first.msgPath}`);
  if (first.attachments?.[0]) {
    console.log(`[dry-run] attachment: ${first.attachments[0]}`);
  }
  console.log(`[dry-run] mode: ${first.msgWriteMode}`);
}

run().catch((error) => {
  console.error("[dry-run] failed", error.message);
  process.exit(1);
});
