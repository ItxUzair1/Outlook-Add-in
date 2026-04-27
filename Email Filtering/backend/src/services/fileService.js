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

  // Track whether Graph enrichment succeeded so we know if post-filing Graph actions can work.
  let graphEnrichmentSucceeded = false;

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

      // CRITICAL: Use the Graph-returned message ID for all subsequent operations.
      // The frontend may send an EWS-format ID (AQMk...) that doesn't work with Graph API
      // write operations. The ID returned by Graph in the response is guaranteed to be valid.
      if (msgData.id && msgData.id !== payload.itemId) {
        console.log(`[fileService] Replacing frontend item ID with Graph-verified ID:`);
        console.log(`   Frontend ID: ${payload.itemId.substring(0, 40)}...`);
        console.log(`   Graph ID:    ${msgData.id.substring(0, 40)}...`);
        finalPayload.itemId = msgData.id;
      }

      graphEnrichmentSucceeded = true;

      finalPayload.subject = msgData.subject || finalPayload.subject;
      finalPayload.body = msgData.body?.content || finalPayload.body;
      finalPayload.isHtml = msgData.body?.contentType === "html";
      finalPayload.attachments = attachments; // Use original attachments from Graph
      finalPayload.sender = msgData.from?.emailAddress?.address || finalPayload.sender;
      finalPayload.sentAt = msgData.sentDateTime || finalPayload.sentAt;

      try {
        finalPayload.rawMimeBase64 = await withGraphAuthFallback((token, options) =>
          graphService.fetchMimeMessage(token, finalPayload.itemId, options)
        );
        console.log("[fileService] MIME stream fetched successfully for MSG fidelity.");
      } catch (mimeError) {
        finalPayload.rawMimeBase64 = null;
        console.warn("[fileService] MIME fetch failed; using compose fallback conversion:", mimeError.message);
      }
      
      console.log("=========================================================");
      console.log(`[fileService] GRAPH API ENRICHMENT SUCCESS`);
      console.log(`[fileService] ItemId (verified): ${finalPayload.itemId}`);
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

      let parentMessagePayload = null;
      if (payload.fileReplyingTo && payload.conversationId) {
        try {
          console.log(`[fileService] "File replying to" enabled. Searching thread: ${payload.conversationId}`);
          const parentMsg = await withGraphAuthFallback((token, options) =>
            graphService.fetchParentMessageInThread(token, payload.conversationId, payload.itemId, options)
          );
          
          if (parentMsg && parentMsg.id) {
            console.log(`[fileService] Found parent message: ${parentMsg.id} - ${parentMsg.subject}`);
            
            const parentAttachments = await withGraphAuthFallback((token, options) =>
              graphService.fetchAttachments(token, parentMsg.id, options)
            );
            
            let parentMime = null;
            try {
              parentMime = await withGraphAuthFallback((token, options) =>
                graphService.fetchMimeMessage(token, parentMsg.id, options)
              );
            } catch (err) {}
            
            parentMessagePayload = {
              ...payload,
              fileReplyingTo: false, // Prevent infinite recursion
              itemId: parentMsg.id,
              internetMessageId: parentMsg.internetMessageId || parentMsg.id,
              subject: parentMsg.subject,
              sender: parentMsg.from?.emailAddress?.address || "",
              to: parentMsg.toRecipients?.map(x => x.emailAddress?.address) || [],
              cc: parentMsg.ccRecipients?.map(x => x.emailAddress?.address) || [],
              sentAt: parentMsg.sentDateTime,
              body: parentMsg.body?.content || "",
              isHtml: parentMsg.body?.contentType === "html",
              attachments: parentAttachments,
              rawMimeBase64: parentMime,
            };
          }
        } catch (parentErr) {
          console.warn("[fileService] Failed to fetch parent message in thread:", parentErr.message);
        }
      }

      if (parentMessagePayload) {
        try {
          console.log(`[fileService] Initiating concurrent filing for parent message: ${parentMessagePayload.subject}`);
          // Don't await here to proceed with the main payload filing quickly
          fileEmail(parentMessagePayload).catch(err => {
            console.error(`[fileService] Background parent filing failed: ${err.message}`);
          });
        } catch (err) {
          console.warn("[fileService] Failed to kick off parent message filing:", err.message);
        }
      }
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
  const useUtc = !!finalPayload.useUtcTime;
  const filedAt = useUtc ? new Date().toISOString() : new Date().toLocaleString("en-US", { timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone });

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

      // Apply read-only attribute if option is enabled
      if (finalPayload.applyReadOnly && msgPath) {
        try {
          await fs.chmod(msgPath, 0o444);
        } catch (roErr) {
          console.warn(`[fileService] Failed to set read-only on ${msgPath}: ${roErr.message}`);
        }
      }
    }

    const attachmentPaths = shouldSaveAttachments
      ? await writeAttachments(folder, finalPayload.attachments)
      : [];

    // Apply read-only attribute to saved attachments if option is enabled
    if (finalPayload.applyReadOnly && attachmentPaths.length > 0) {
      for (const attPath of attachmentPaths) {
        try {
          await fs.chmod(attPath, 0o444);
        } catch (roErr) {
          console.warn(`[fileService] Failed to set read-only on ${attPath}: ${roErr.message}`);
        }
      }
    }

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
      // Append timestamp to ensure ID uniqueness even if filing the same email to the same path multiple times.
      id: `${finalPayload.internetMessageId || finalPayload.subject}-${x.msgPath || x.targetPath}-${Date.now()}`,
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
      body: finalPayload.body || finalPayload.bodyPreview || "",
      sendLink: !!finalPayload.sendLink,
    }));

    // Simple deduplication: don't add if the EXACT same filePath and internetMessageId already exist in the last few minutes 
    // or just rely on the unique ID for UI, but let's prevent spamming the index.
    const filteredRows = rows.filter(newRow => 
        !existingIndex.some(oldRow => 
            oldRow.filePath === newRow.filePath && 
            oldRow.internetMessageId === newRow.internetMessageId &&
            newRow.internetMessageId !== null // Only deduplicate if we have a real ID
        )
    );

    if (filteredRows.length > 0) {
        await saveSearchIndex([...filteredRows, ...existingIndex]);
    }

    // Optional post-filing actions driven by checkboxes.
    // These ONLY work if Graph enrichment succeeded (which gives us a verified item ID).
    if (!graphEnrichmentSucceeded && (finalPayload.markReviewed || finalPayload.addFiledCategory || 
        (finalPayload.afterFiling && finalPayload.afterFiling !== "none"))) {
      console.warn("[fileService] Graph enrichment failed earlier — skipping post-filing Graph actions (category, move, archive, etc.).");
      appendPostFilingError("Post-filing actions skipped: could not verify email ID with Microsoft Graph. The email was saved to disk successfully.");
    }

    if (graphEnrichmentSucceeded) {
      // 1. Mark as reviewed
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

      // 2. Add "Filed" category BEFORE any move/archive (moving changes the item ID)
      if (graphAuthToken && finalPayload.itemId && finalPayload.addFiledCategory) {
        try {
          const categoryName = finalPayload.filedCategoryName || 'Filed';
          await withGraphAuthFallback((token, options) =>
            graphService.addCategoryToEmail(token, finalPayload.itemId, categoryName, options)
          );
          console.log(`[fileService] Successfully added "${categoryName}" category.`);
        } catch (err) {
          appendPostFilingError(`Add filed category failed: ${err.message}`);
          console.error("[fileService] [FS-POST-FAIL] Add category:", err.message);
        }
      }

      // 3. Add additional categories if provided (Assistant Categories)
      if (graphAuthToken && finalPayload.itemId && finalPayload.assistantCategories) {
        const extraCats = finalPayload.assistantCategories.split(',').map(c => c.trim()).filter(Boolean);
        for (const cat of extraCats) {
          try {
            await withGraphAuthFallback((token, options) =>
              graphService.addCategoryToEmail(token, finalPayload.itemId, cat, options)
            );
          } catch (err) {
            console.warn(`[fileService] Failed to add extra category ${cat}:`, err.message);
          }
        }
      }

      // 4. Add filed date to subject (non-moving action, do before move)
      if (graphAuthToken && finalPayload.itemId && finalPayload.afterFiling === "add_date") {
        try {
          const dateStr = useUtc 
            ? new Date().toISOString().replace('T', ' ').substring(0, 19) + ' UTC'
            : new Date().toLocaleString();
          const newSubject = `[Filed ${dateStr}] ${finalPayload.subject || ''}`;
          await graphService.updateEmailSubject(graphAuthToken, finalPayload.itemId, newSubject, graphAuthOptions);
          console.log(`[fileService] Successfully updated subject to: ${newSubject}`);
        } catch (dateErr) {
          appendPostFilingError(`Add date to subject failed: ${dateErr.message}`);
        }
      }

      // 5. Handle move/archive LAST (these change the item ID, so no Graph calls after this)
      if (graphAuthToken && finalPayload.itemId && finalPayload.afterFiling && 
          finalPayload.afterFiling !== "none" && finalPayload.afterFiling !== "add_date") {
        try {
          if (finalPayload.afterFiling === "delete" || finalPayload.afterFiling === "move_deleted") {
            await graphService.deleteEmail(graphAuthToken, finalPayload.itemId, graphAuthOptions);
          } else if (finalPayload.afterFiling === "archive") {
            await graphService.archiveEmail(graphAuthToken, finalPayload.itemId, graphAuthOptions);
          } else if (finalPayload.afterFiling === "move_filed_items") {
            const folderId = await graphService.getOrCreateMailFolder(graphAuthToken, 'inbox', 'Filed Items', graphAuthOptions);
            await graphService.moveEmail(graphAuthToken, finalPayload.itemId, folderId, graphAuthOptions);
          } else if (finalPayload.afterFiling === "move_filed_folders") {
            const prefix = finalPayload.filedFolderPrefix || '*';
            const locationName = targets.length > 0 ? targets[0].split(/[\\/]/).filter(Boolean).pop() : 'Filed';
            const folderName = `${prefix} ${locationName}`.trim();
            const folderId = await graphService.getOrCreateMailFolder(graphAuthToken, 'inbox', folderName, graphAuthOptions);
            await graphService.moveEmail(graphAuthToken, finalPayload.itemId, folderId, graphAuthOptions);
            
            if (finalPayload.deleteEmptyFolders) {
              await graphService.cleanupEmptyFolders(graphAuthToken, 'inbox', prefix, graphAuthOptions);
            }
          }
        } catch (err) {
          appendPostFilingError(`Post-filing action (${finalPayload.afterFiling}) failed: ${err.message}`);
        }
      }
    } // end graphEnrichmentSucceeded


    const firstSavedPath = successful.length > 0 ? (successful[0].msgPath || null) : null;

    // Debug: log all paths so we can trace why sharing links may be empty
    console.log(`[fileService] sendLink=${finalPayload.sendLink}, successful=${successful.length}`);
    successful.forEach((entry, i) => {
      console.log(`[fileService]   entry[${i}] msgPath="${entry.msgPath}" targetPath="${entry.targetPath}"`);
    });

    const sharingLinks = (finalPayload.sendLink && successful.length > 0)
      ? successful
          .map((entry) => entry.msgPath || entry.targetPath)
          .filter(p => !!p)
      : [];
    
    console.log(`[fileService] sharingLinks generated: ${JSON.stringify(sharingLinks)}`);

    return {
      fileName: firstSavedPath ? path.basename(firstSavedPath) : msgName,
      filedAt,
      results: perTarget,
      postFilingError,
      sharingLinks,
    };
  }

  const firstSavedPath = perTarget.find((x) => x.msgPath)?.msgPath || null;

  return {
    fileName: firstSavedPath ? path.basename(firstSavedPath) : msgName,
    filedAt,
    results: perTarget,
    postFilingError,
  };
}
