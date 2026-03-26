import fs from "fs/promises";
import path from "path";
import os from "os";
import { execFile } from "child_process";
import { promisify } from "util";
import { v4 as uuidv4 } from "uuid";
import { buildMsgFileName, sanitizeFileName } from "../utils/fileName.js";
import { config } from "../config/index.js";
import { getSearchIndex, saveSearchIndex } from "../storage/repositories.js";
import { markUsedByPaths } from "./locationService.js";

const execFileAsync = promisify(execFile);

function resolveTarget(targetPath) {
  if (path.isAbsolute(targetPath)) {
    return targetPath;
  }

  return path.join(config.fileStorageRoot, targetPath);
}

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function uniqueFilePath(basePath) {
  const ext = path.extname(basePath);
  const head = basePath.slice(0, -ext.length);
  let i = 1;
  while (await exists(`${head}(${i})${ext}`)) {
    i += 1;
  }
  return `${head}(${i})${ext}`;
}

function buildPseudoMsg(payload) {
  const data = {
    internetMessageId: payload.internetMessageId || "",
    subject: payload.subject || "",
    sender: payload.sender || "",
    to: payload.to || [],
    cc: payload.cc || [],
    sentAt: payload.sentAt || "",
    bodyPreview: payload.bodyPreview || "",
    note: "Milestone 2 placeholder MSG payload. Replace with real MSG conversion in later iteration.",
  };

  return Buffer.from(JSON.stringify(data, null, 2), "utf-8");
}

function buildMimeMessage(payload) {
  const subject = payload.subject || "No Subject";
  const from = payload.sender || "unknown@example.com";
  const to = Array.isArray(payload.to) && payload.to.length > 0 ? payload.to.join(", ") : "undisclosed-recipients:;";
  const cc = Array.isArray(payload.cc) && payload.cc.length > 0 ? payload.cc.join(", ") : "";
  const date = payload.sentAt || new Date().toUTCString();
  const body = payload.bodyPreview || "";

  const lines = [
    `From: ${from}`,
    `To: ${to}`,
    cc ? `Cc: ${cc}` : "",
    `Subject: ${subject}`,
    `Date: ${date}`,
    "MIME-Version: 1.0",
    'Content-Type: text/plain; charset="utf-8"',
    "",
    body,
  ].filter(Boolean);

  return Buffer.from(lines.join("\r\n"), "utf-8");
}

async function convertPayloadToMsgWithOutlook(payload, targetMsgPath) {
  const tempDir = path.join(os.tmpdir(), "email-filing-msg");
  await fs.mkdir(tempDir, { recursive: true });

  const attachmentLines = [];
  if (Array.isArray(payload.attachments)) {
    for (const att of payload.attachments) {
      if (att.name && att.base64Content) {
        const tempAttPath = path.join(tempDir, `${uuidv4()}_${sanitizeFileName(att.name)}`);
        await fs.writeFile(tempAttPath, Buffer.from(att.base64Content, "base64"));
        attachmentLines.push(`$item.Attachments.Add('${tempAttPath.replace(/'/g, "''")}') | Out-Null`);
      }
    }
  }

  const escapedSubject = (payload.subject || "").replace(/'/g, "''");
  const escapedBody = (payload.bodyPreview || "").replace(/'/g, "''");
  const escapedMsg = targetMsgPath.replace(/'/g, "''");
  
  const toList = Array.isArray(payload.to) ? payload.to.join("; ") : "";
  const ccList = Array.isArray(payload.cc) ? payload.cc.join("; ") : "";

  const psScript = [
    "$ErrorActionPreference = 'Stop'",
    "$outlook = New-Object -ComObject Outlook.Application",
    "$item = $outlook.CreateItem(0)",
    `$item.Subject = '${escapedSubject}'`,
    `$item.Body = '${escapedBody}'`,
    toList ? `$item.To = '${toList.replace(/'/g, "''")}'` : "",
    ccList ? `$item.CC = '${ccList.replace(/'/g, "''")}'` : "",
    ...attachmentLines,
    // Add a note about the conversion
    "$item.Body = '--- Filed via Mail Manager ---' + [char]13 + [char]10 + $item.Body",
    `$item.SaveAs('${escapedMsg}', 9)`,
    "$item.Close(1)", // 1 = olDiscard (it's already saved)
  ].filter(Boolean).join("; ");

  try {
    await execFileAsync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", psScript], {
      windowsHide: true,
      timeout: 30000,
    });
  } finally {
    // Cleanup temp attachments would be good here, but they are in OS temp
  }
}

async function writeMsgByStrategy(msgPath, payload) {
  const strategy = (payload.msgStrategy || config.msgStrategy || "pseudo").toLowerCase();

  if (strategy === "outlook-com") {
    try {
      await convertPayloadToMsgWithOutlook(payload, msgPath);
      return "outlook-com";
    } catch (error) {
      if (config.strictMsgRequired) {
        throw new Error(`Strict MSG generation failed (Outlook COM): ${error.message}`);
      }
    }
  }

  await fs.writeFile(msgPath, buildPseudoMsg(payload));
  return "pseudo";
}

async function writeAttachments(baseFolder, attachments) {
  if (!Array.isArray(attachments) || attachments.length === 0) {
    return [];
  }

  const attachmentDir = path.join(baseFolder, "Attachments");
  await fs.mkdir(attachmentDir, { recursive: true });

  const saved = [];
  for (const att of attachments) {
    if (!att || !att.name || !att.base64Content) {
      continue;
    }

    const safeName = sanitizeFileName(att.name);
    const filePath = path.join(attachmentDir, safeName);
    await fs.writeFile(filePath, Buffer.from(att.base64Content, "base64"));
    saved.push(filePath);
  }

  return saved;
}

export async function fileEmail(payload) {
  const targets = Array.isArray(payload.targetPaths) ? payload.targetPaths : [];
  const duplicateStrategy = payload.duplicateStrategy || "rename";
  const msgName = buildMsgFileName(payload.subject, payload.sentAt);
  const filedAt = new Date().toISOString();

  const perTarget = [];

  for (const target of targets) {
    const folder = resolveTarget(target);
    await fs.mkdir(folder, { recursive: true });

    let msgPath = path.join(folder, msgName);
    const alreadyThere = await exists(msgPath);

    if (alreadyThere && duplicateStrategy === "skip") {
      perTarget.push({ targetPath: folder, msgPath, status: "skipped", attachments: [] });
      continue;
    }

    if (alreadyThere && duplicateStrategy === "rename") {
      msgPath = await uniqueFilePath(msgPath);
    }

    const msgWriteMode = await writeMsgByStrategy(msgPath, payload);
    const attachmentPaths = await writeAttachments(folder, payload.attachments);

    perTarget.push({
      targetPath: folder,
      msgPath,
      msgWriteMode,
      status: alreadyThere && duplicateStrategy === "overwrite" ? "overwritten" : "saved",
      attachments: attachmentPaths,
    });
  }

  const successful = perTarget.filter((x) => x.status === "saved" || x.status === "overwritten");
  if (successful.length > 0) {
    await markUsedByPaths(successful.map((x) => x.targetPath));

    const existingIndex = await getSearchIndex();
    const rows = successful.map((x) => ({
      id: `${payload.internetMessageId || payload.subject}-${x.msgPath}`,
      internetMessageId: payload.internetMessageId || null,
      subject: payload.subject || "",
      sender: payload.sender || "",
      recipients: payload.to || [],
      cc: payload.cc || [],
      sentAt: payload.sentAt || filedAt,
      filedAt,
      hasAttachments: Array.isArray(payload.attachments) && payload.attachments.length > 0,
      filePath: x.msgPath,
      comment: payload.comment || "",
      markReviewed: !!payload.markReviewed,
      sendLink: !!payload.sendLink,
    }));

    await saveSearchIndex([...rows, ...existingIndex]);
  }

  return {
    fileName: msgName,
    filedAt,
    results: perTarget,
  };
}
