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

function buildEmlFile(payload) {
  if (payload.rawMimeBase64 && (!payload.attachments || payload.attachments.length > 0)) {
    // If we have raw MIME and we are NOT explicitly stripping attachments, use it directly!
    return Buffer.from(payload.rawMimeBase64, "base64");
  }

  const boundary = "----=_NextPart_Boundary_" + Date.now();
  
  let eml = [];
  eml.push(`From: ${payload.sender || ""}`);
  eml.push(`To: ${(payload.to || []).join(", ")}`);
  eml.push(`Cc: ${(payload.cc || []).join(", ")}`);
  eml.push(`Subject: ${payload.subject || ""}`);
  eml.push(`Date: ${payload.sentAt ? new Date(payload.sentAt).toUTCString() : new Date().toUTCString()}`);
  eml.push(`MIME-Version: 1.0`);

  if (payload.attachments && payload.attachments.length > 0) {
    eml.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
    eml.push(``);
    eml.push(`--${boundary}`);
    eml.push(`Content-Type: ${payload.isHtml ? 'text/html' : 'text/plain'}; charset="utf-8"`);
    eml.push(``);
    eml.push(payload.body || payload.bodyPreview || "");
    eml.push(``);

    for (const att of payload.attachments) {
      if (!att.name || !att.base64Content) continue;
      eml.push(`--${boundary}`);
      eml.push(`Content-Type: application/octet-stream; name="${att.name}"`);
      eml.push(`Content-Transfer-Encoding: base64`);
      eml.push(`Content-Disposition: attachment; filename="${att.name}"`);
      eml.push(``);
      // Chunk base64 to 76 chars
      const b64 = att.base64Content || "";
      for (let i = 0; i < b64.length; i += 76) {
        eml.push(b64.substring(i, i + 76));
      }
      eml.push(``);
    }
    eml.push(`--${boundary}--`);
  } else {
    eml.push(`Content-Type: ${payload.isHtml ? 'text/html' : 'text/plain'}; charset="utf-8"`);
    eml.push(``);
    eml.push(payload.body || payload.bodyPreview || "");
  }

  return Buffer.from(eml.join("\r\n"), "utf-8");
}

async function writeEmlByStrategy(emlPath, payload) {
  const buffer = buildEmlFile(payload);
  await fs.writeFile(emlPath, buffer);
  return { mode: "eml", path: emlPath };
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
    finalPayload.bodyPreview = `[Koyomail] Message body could not be retrieved in this Outlook host. Subject: ${finalPayload.subject || "No Subject"}`;
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

      // Force no embedded attachments and fallback to basic MSG build for "message" mode.
      const msgPayload = shouldSaveAttachments ? finalPayload : { ...finalPayload, attachments: [], rawMimeBase64: null };
      const writeResult = await writeEmlByStrategy(msgPath, msgPayload);
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
