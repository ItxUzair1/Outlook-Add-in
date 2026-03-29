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
import * as graphService from "./graphService.js";

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

async function convertPayloadToMsgWithOutlook(payload, targetMsgPath) {
  const tempDir = path.join(os.tmpdir(), "email-filing-msg");
  await fs.mkdir(tempDir, { recursive: true });

  const escapedMsg = targetMsgPath.replace(/'/g, "''");

  // Prefer MIME-based conversion because it preserves received-message semantics
  // (read mode) and inline attachment CID mappings.
  if (payload.rawMimeBase64) {
    const tempEmlPath = path.join(tempDir, `${uuidv4()}_source.eml`);
    await fs.writeFile(tempEmlPath, Buffer.from(payload.rawMimeBase64, "base64"));

    const psMIME = [
      "$ErrorActionPreference = 'Stop'",
      "$outlook = New-Object -ComObject Outlook.Application",
      "$ns = $outlook.GetNamespace('MAPI')",
      `$emlPath = '${tempEmlPath.replace(/'/g, "''")}'`,
      "if (-not (Test-Path -LiteralPath $emlPath)) { throw ('MIME source file not found: ' + $emlPath) }",
      "$resolvedEml = (Resolve-Path -LiteralPath $emlPath).Path",
      "$item = $null",
      "try { $item = $ns.OpenSharedItem($resolvedEml) } catch {",
      "  $fileUrl = 'file:///' + (($resolvedEml -replace '\\\\','/') -replace ' ','%20')",
      "  $item = $ns.OpenSharedItem($fileUrl)",
      "}",
      `$item.SaveAs('${escapedMsg}', 9)`, // 9 = olMsgUnicode
      "$item.Close(1)",
    ].join("; ");

    try {
      const { stderr } = await execFileAsync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", psMIME], {
        windowsHide: true,
        timeout: 45000,
      });
      if (stderr) console.warn("[fileService] PowerShell stderr (MIME path):", stderr);
      return;
    } catch (error) {
      // Bubble up so writeMsgByStrategy can switch to .eml fidelity fallback.
      throw new Error(`MIME-based conversion failed: ${error.message}`);
    }
  }

  const attachmentLines = [];
  if (Array.isArray(payload.attachments)) {
    for (const att of payload.attachments) {
      if (att && att.name && (att.base64Content || att.base64Content === "")) {
        const tempAttPath = path.join(tempDir, `${uuidv4()}_${sanitizeFileName(att.name)}`);
        const buffer = Buffer.from(att.base64Content || "", "base64");
        await fs.writeFile(tempAttPath, buffer);
        attachmentLines.push(`$item.Attachments.Add('${tempAttPath.replace(/'/g, "''")}', 1) | Out-Null`);
      }
    }
  }

  // Write body to a temporary file to avoid PowerShell escaping/interpolation issues
  const tempBodyPath = path.join(tempDir, `${uuidv4()}_body.html`);
  const bodyContent = payload.isHtml
    ? `<html><body>${payload.body || ""}</body></html>`
    : `<html><body><pre>${(payload.body || payload.bodyPreview || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</pre></body></html>`;
  
  await fs.writeFile(tempBodyPath, bodyContent, "utf-8");

  const escapedSubject = (payload.subject || "No Subject").replace(/'/g, "''");
  const sentAtStamp = new Date(payload.sentAt || Date.now()).toISOString().replace(/'/g, "''");
  const toList = Array.isArray(payload.to) ? payload.to.join("; ") : "";
  const ccList = Array.isArray(payload.cc) ? payload.cc.join("; ") : "";

  console.log(`[fileService] Converting email to MSG (Outlook COM). Attachments: ${payload.attachments?.length || 0}`);
  
  const psScript = [
    "$ErrorActionPreference = 'Stop'",
    "$outlook = New-Object -ComObject Outlook.Application",
    "$item = $outlook.CreateItem(0)",
    `$item.Subject = '${escapedSubject}'`,
    toList ? `$item.To = '${toList.replace(/'/g, "''")}'` : "",
    ccList ? `$item.CC = '${ccList.replace(/'/g, "''")}'` : "",
    "$item.BodyFormat = 2", // olFormatHTML
    `$html = Get-Content '${tempBodyPath.replace(/'/g, "''")}' -Raw -Encoding UTF8`,
    "$item.HTMLBody = $html",
    ...attachmentLines,
    // Best-effort MAPI flags to make saved .msg open in read mode instead of compose.
    "try {",
    "  $pa = $item.PropertyAccessor",
    "  $flagsTag = 'http://schemas.microsoft.com/mapi/proptag/0x0E070003'",
    "  $submitTag = 'http://schemas.microsoft.com/mapi/proptag/0x00390040'",
    "  $deliveryTag = 'http://schemas.microsoft.com/mapi/proptag/0x0E060040'",
    "  $flags = 0",
    "  try { $flags = [int]$pa.GetProperty($flagsTag) } catch { $flags = 0 }",
    "  $flags = $flags -band (-bnot 8)",
    "  $pa.SetProperty($flagsTag, $flags)",
    `  $stamp = [DateTime]::Parse('${sentAtStamp}')`,
    "  $pa.SetProperty($submitTag, $stamp)",
    "  $pa.SetProperty($deliveryTag, $stamp)",
    "  $item.UnRead = $false",
    "} catch { Write-Host ('MSG read-mode hint failed: ' + $_.Exception.Message) }",
    "$item.Save()",
    `$item.SaveAs('${escapedMsg}', 9)`, // 9 = olMsg
    "$item.Close(1)", // 1 = olDiscard
  ].filter(Boolean).join("; ");

  try {
    const { stderr } = await execFileAsync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", psScript], {
      windowsHide: true,
      timeout: 30000,
    });
    if (stderr) console.warn("[fileService] PowerShell stderr:", stderr);
  } catch (error) {
    console.error("[fileService] PowerShell conversion failed:", error.message);
    throw error;
  }
}

function toExtension(filePath, extensionWithoutDot) {
  const ext = path.extname(filePath);
  const head = ext ? filePath.slice(0, -ext.length) : filePath;
  return `${head}.${extensionWithoutDot}`;
}

async function writeMsgByStrategy(msgPath, payload) {
  const strategy = (payload.msgStrategy || config.msgStrategy || "pseudo").toLowerCase();

  if (strategy === "outlook-com") {
    try {
      await convertPayloadToMsgWithOutlook(payload, msgPath);
      return { mode: "outlook-com", path: msgPath };
    } catch (error) {
      // Host-specific fallback: preserve message fidelity by saving Graph MIME as .eml
      // when Outlook cannot convert MIME to .msg via OpenSharedItem.
      if (payload.rawMimeBase64) {
        const emlPath = toExtension(msgPath, "eml");
        await fs.writeFile(emlPath, Buffer.from(payload.rawMimeBase64, "base64"));
        console.warn("[fileService] MIME->MSG conversion unavailable on this host. Saved as .eml fidelity fallback.");
        return { mode: "eml-mime", path: emlPath };
      }

      if (config.strictMsgRequired) {
        throw new Error(`Strict MSG generation failed (Outlook COM): ${error.message}`);
      }
    }
  }

  await fs.writeFile(msgPath, buildPseudoMsg(payload));
  return { mode: "pseudo", path: msgPath };
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
  let finalPayload = { ...payload };
  let postFilingError = null;
  const normalizedAccessToken = typeof payload.graphAccessToken === "string"
    ? payload.graphAccessToken.trim()
    : "";
  const normalizedSsoToken = typeof payload.ssoToken === "string"
    ? payload.ssoToken.trim()
    : "";

  // SSO-first policy: prefer SSO/OBO token when available, then fallback to direct MSAL access token.
  let graphAuthToken = normalizedSsoToken || normalizedAccessToken || null;
  let graphAuthOptions = { isAccessToken: !normalizedSsoToken && !!normalizedAccessToken };
  
  // Safe fallback: if we have a manual access token and it's long enough, always consider it a fallback.
  const fallbackGraphAuthToken = (normalizedAccessToken && normalizedAccessToken.length > 10) 
    ? normalizedAccessToken 
    : null;
  const fallbackGraphAuthOptions = { isAccessToken: true };

  const isGraphAuthFailure = (error) => {
    const msg = String(error?.message || error || "").toLowerCase();
    return (
      msg.includes("invalidauthenticationtoken") ||
      msg.includes("access token is empty") ||
      msg.includes("graph token exchange failed") ||
      msg.includes("no authentication token was provided") ||
      msg.includes("401")
    );
  };

  const withGraphAuthFallback = async (operation) => {
    if (!graphAuthToken) {
      throw new Error("No authentication token available for Graph operation.");
    }

    try {
      return await operation(graphAuthToken, graphAuthOptions);
    } catch (primaryError) {
      if (!fallbackGraphAuthToken || !isGraphAuthFailure(primaryError)) {
        console.warn(`[fileService] Graph primary auth failed (OBO/SSO), but no valid fallback available: ${primaryError.message}`);
        throw primaryError;
      }

      console.log(`[fileService] Graph primary auth failed: ${primaryError.message}. Attempting fallback to manual access token (${fallbackGraphAuthToken.length} chars).`);
      const fallbackResult = await operation(fallbackGraphAuthToken, fallbackGraphAuthOptions);
      
      // If fallback succeeded, update the token state for subsequent operations in this request.
      graphAuthToken = fallbackGraphAuthToken;
      graphAuthOptions = fallbackGraphAuthOptions;
      return fallbackResult;
    }
  };

  // If we have a Graph-capable token and itemId, enrich from Microsoft Graph API.
  // This ensures attachments are not 0 bytes even if the frontend failed to gather them.
  if (graphAuthToken && payload.itemId) {
    try {
      console.log(`[fileService] Enriching payload via Microsoft Graph for item: ${payload.itemId}`);
      
      const [msgData, attachments] = await withGraphAuthFallback((token, options) =>
        Promise.race([
          Promise.all([
            graphService.fetchEmailMessage(token, payload.itemId, options),
            graphService.fetchAttachments(token, payload.itemId, options),
          ]),
          new Promise((_, reject) => setTimeout(() => reject(new Error("Graph API timeout")), 25000))
        ])
      );

      finalPayload.subject = msgData.subject || finalPayload.subject;
      finalPayload.body = msgData.body?.content || finalPayload.body;
      finalPayload.isHtml = msgData.body?.contentType === "html";
      finalPayload.attachments = attachments; // Use original attachments from Graph
      finalPayload.sender = msgData.from?.emailAddress?.address || finalPayload.sender;
      finalPayload.sentAt = msgData.sentDateTime || finalPayload.sentAt;

      try {
        finalPayload.rawMimeBase64 = await withGraphAuthFallback((token, options) =>
          graphService.fetchMimeMessage(token, payload.itemId, options)
        );
        console.log("[fileService] MIME stream fetched successfully for MSG fidelity.");
      } catch (mimeError) {
        finalPayload.rawMimeBase64 = null;
        console.warn("[fileService] MIME fetch failed; using compose fallback conversion:", mimeError.message);
      }
      
      console.log("=========================================================");
      console.log(`[fileService] GRAPH API ENRICHMENT SUCCESS`);
      console.log(`[fileService] ItemId: ${payload.itemId}`);
      console.log(`[fileService] Subject: "${msgData.subject}"`);
      console.log(`[fileService] Body present: ${!!msgData.body?.content}`);
      console.log(`[fileService] Body content-type: ${msgData.body?.contentType}`);
      console.log(`[fileService] Body length: ${msgData.body?.content?.length || 0} characters`);
      
      console.log(`[fileService] Attachments found: ${attachments.length}`);
      if (attachments.length > 0) {
        attachments.forEach((att, idx) => {
          console.log(`   - Attachment ${idx + 1}: ${att.name} (Size: ${att.size || 0} bytes)`);
          // Note: we don't log the base64 content entirely because it would overflow the terminal, but we confirm its presence
          console.log(`     -> Base64 content present: ${!!att.base64Content}, Length: ${att.base64Content?.length || 0}`);
        });
      }
      console.log("=========================================================");
    } catch (error) {
      console.error("[fileService] Graph enrichment failed, falling back to local payload:", error.message);
      // Fallback: stay with original payload
    }
  } else {
    console.log(`[fileService] Skipping Graph enrichment. Graph Token: ${!!graphAuthToken}, ItemId: ${!!payload.itemId}`);
    if (Array.isArray(payload.attachments) && payload.attachments.length > 0) {
      console.log(`[fileService] Using ${payload.attachments.length} attachments from frontend.`);
    }
  }

  // Guarantee a non-empty body in the saved MSG, even if host APIs return no content.
  const hasBody = typeof finalPayload.body === "string" && finalPayload.body.trim().length > 0;
  const hasPreview = typeof finalPayload.bodyPreview === "string" && finalPayload.bodyPreview.trim().length > 0;
  if (!hasBody && !hasPreview) {
    finalPayload.bodyPreview = `[Mail Manager] Message body could not be retrieved in this Outlook host. Subject: ${finalPayload.subject || "No Subject"}`;
  }

  const targets = Array.isArray(finalPayload.targetPaths) ? finalPayload.targetPaths : [];
  const duplicateStrategy = finalPayload.duplicateStrategy || "rename";
  const attachmentsOption = (finalPayload.attachmentsOption || "all").toLowerCase();
  const shouldSaveMessage = attachmentsOption !== "attachments";
  const shouldSaveAttachments = attachmentsOption !== "message";
  const msgName = buildMsgFileName(finalPayload.subject, finalPayload.sentAt);
  const filedAt = new Date().toISOString();

  const perTarget = [];

  const appendPostFilingError = (message) => {
    if (!message) return;
    postFilingError = postFilingError ? `${postFilingError} | ${message}` : message;
  };

  for (const target of targets) {
    const folder = resolveTarget(target);
    await fs.mkdir(folder, { recursive: true });

    let msgPath = null;
    let msgWriteMode = null;
    let alreadyThere = false;

    if (shouldSaveMessage) {
      msgPath = path.join(folder, msgName);
      alreadyThere = await exists(msgPath);

      if (alreadyThere && duplicateStrategy === "skip") {
        perTarget.push({ targetPath: folder, msgPath, status: "skipped", attachments: [] });
        continue;
      }

      if (alreadyThere && duplicateStrategy === "rename") {
        msgPath = await uniqueFilePath(msgPath);
      }

      // Force no embedded attachments for "message" mode.
      const msgPayload = shouldSaveAttachments ? finalPayload : { ...finalPayload, attachments: [] };
      const writeResult = await writeMsgByStrategy(msgPath, msgPayload);
      msgWriteMode = writeResult.mode;
      msgPath = writeResult.path;
    }

    const attachmentPaths = shouldSaveAttachments
      ? await writeAttachments(folder, finalPayload.attachments)
      : [];

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
      id: `${finalPayload.internetMessageId || finalPayload.subject}-${x.msgPath || x.targetPath}`,
      internetMessageId: finalPayload.internetMessageId || null,
      subject: finalPayload.subject || "",
      sender: finalPayload.sender || "",
      recipients: finalPayload.to || [],
      cc: finalPayload.cc || [],
      sentAt: finalPayload.sentAt || filedAt,
      filedAt,
      hasAttachments: Array.isArray(x.attachments) && x.attachments.length > 0,
      filePath: x.msgPath || x.attachments[0] || x.targetPath,
      comment: finalPayload.comment || "",
      markReviewed: !!finalPayload.markReviewed,
      sendLink: !!finalPayload.sendLink,
    }));

    await saveSearchIndex([...rows, ...existingIndex]);

    // Optional post-filing actions driven by checkboxes.
    if (graphAuthToken && finalPayload.itemId && finalPayload.markReviewed) {
      try {
        await withGraphAuthFallback((token, options) =>
          graphService.markEmailReviewed(token, finalPayload.itemId, options)
        );
      } catch (err) {
        appendPostFilingError(`[FS-POST-FAIL] Mark as reviewed: ${err.message}`);
        console.error("[fileService] [FS-POST-FAIL]", err.message);
      }
    }

    if (graphAuthToken && finalPayload.sendLink) {
      try {
        const filedEntries = successful
          .map((entry) => entry.msgPath || entry.targetPath)
          .filter(Boolean);

        await withGraphAuthFallback((token, options) =>
          graphService.sendFilingLinkEmail(
            token,
            {
              originalSubject: finalPayload.subject,
              filedAt,
              comment: finalPayload.comment,
              filedEntries,
            },
            options
          )
        );
      } catch (err) {
        appendPostFilingError(`[FS-LINK-FAIL] Send link: ${err.message}`);
        console.error("[fileService] [FS-LINK-FAIL]", err.message);
      }
    }

    // Handle post-filing move/archive in backend only for SSO flows.
    // MSAL fallback flows are handled by the frontend to avoid duplicate operations.
    if (payload.ssoToken && finalPayload.itemId && finalPayload.afterFiling && finalPayload.afterFiling !== "none") {
      try {
        console.log(`[fileService] Performing post-filing action: ${finalPayload.afterFiling}`);
        if (finalPayload.afterFiling === "delete") {
          await graphService.deleteEmail(graphAuthToken, finalPayload.itemId, graphAuthOptions);
        } else if (finalPayload.afterFiling === "archive") {
          await graphService.archiveEmail(graphAuthToken, finalPayload.itemId, graphAuthOptions);
        }
        console.log(`[fileService] Post-filing action ${finalPayload.afterFiling} completed successfully.`);
      } catch (err) {
        appendPostFilingError(`Post-filing action (${finalPayload.afterFiling}) failed: ${err.message}`);
        console.error("[fileService]", err.message);
      }
    }
  }

  const firstSavedPath = perTarget.find((x) => x.msgPath)?.msgPath || null;

  return {
    fileName: firstSavedPath ? path.basename(firstSavedPath) : msgName,
    filedAt,
    results: perTarget,
    postFilingError,
  };
}
