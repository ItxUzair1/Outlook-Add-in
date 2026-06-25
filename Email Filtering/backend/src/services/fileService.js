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
  if (payload.rawMimeBase64) {
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

  // ── Safety net: auto-detect SSO identity tokens routed to the wrong field ──
  // Office SSO tokens (from Office.auth.getAccessToken) have aud=api://{clientId},
  // NOT aud=https://graph.microsoft.com. If such a token reaches us as graphAccessToken,
  // calling Graph API directly returns 401. Detect this and re-route via OBO.
  const isLikelySsoToken = (token) => {
    if (!token || token.length < 20) return false;
    try {
      const parts = token.split(".");
      if (parts.length !== 3) return false;
      const raw = parts[1].replace(/-/g, "+").replace(/_/g, "/");
      const padded = raw + "=".repeat((4 - raw.length % 4) % 4);
      const decoded = JSON.parse(Buffer.from(padded, "base64").toString("utf-8"));
      const aud = String(decoded.aud || "");
      // Graph tokens have aud containing "graph.microsoft.com" or the Graph app GUID
      return !aud.includes("graph.microsoft.com") && !aud.includes("00000003-0000-0000-c000-000000000000");
    } catch {
      return false;
    }
  };

  // If the frontend sent an SSO token in graphAccessToken (token type mismatch),
  // re-route it to the ssoToken path so the backend performs the OBO exchange.
  let effectiveSsoToken = normalizedSsoToken;
  let effectiveAccessToken = normalizedAccessToken;
  if (!normalizedSsoToken && normalizedAccessToken && isLikelySsoToken(normalizedAccessToken)) {
    console.warn("[fileService] Detected SSO identity token in graphAccessToken field — re-routing via OBO exchange.");
    effectiveSsoToken = normalizedAccessToken;
    effectiveAccessToken = "";
  }

  // SSO-first policy: prefer SSO/OBO token when available, then fallback to direct MSAL access token.
  let graphAuthToken = effectiveSsoToken || effectiveAccessToken || null;
  let graphAuthOptions = { isAccessToken: !effectiveSsoToken && !!effectiveAccessToken };
  
  // Safe fallback: if we have a manual access token and it's long enough, always consider it a fallback.
  const fallbackGraphAuthToken = (effectiveAccessToken && effectiveAccessToken.length > 10) 
    ? effectiveAccessToken 
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

  const attachmentsOption = (finalPayload.attachmentsOption || "all").toLowerCase();
  const shouldSaveMessage = attachmentsOption !== "attachments";
  const shouldEmbedAttachments = attachmentsOption !== "message";
  const shouldWriteSeparateAttachments = attachmentsOption === "attachments";

  // Track whether Graph enrichment succeeded so we know if post-filing Graph actions can work.
  let graphEnrichmentSucceeded = false;

  // If we have a Graph-capable token and itemId, enrich from Microsoft Graph API.
  // This ensures attachments are not 0 bytes even if the frontend failed to gather them.
  if (graphAuthToken && payload.itemId) {
    try {
      console.log(`[fileService] Enriching payload via Microsoft Graph for item: ${payload.itemId}`);
      
      const [msgData, attachments, mimeBase64] = await withGraphAuthFallback((token, options) =>
        Promise.race([
          Promise.all([
            graphService.fetchEmailMessage(token, payload.itemId, options),
            shouldWriteSeparateAttachments 
              ? graphService.fetchAttachments(token, payload.itemId, options) 
              : Promise.resolve(payload.attachments || []),
            (shouldEmbedAttachments && shouldSaveMessage) 
              ? graphService.fetchMimeMessage(token, payload.itemId, options).catch(e => {
                  console.warn("[fileService] MIME fetch failed; using compose fallback conversion:", e.message);
                  return null;
                })
              : Promise.resolve(null)
          ]),
          new Promise((_, reject) => setTimeout(() => reject(new Error("Graph API timeout")), 90000))
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
      finalPayload.hasAttachments = msgData.hasAttachments; // Save real graph attachment state
      finalPayload.sender = msgData.from?.emailAddress?.address || finalPayload.sender;
      finalPayload.to = msgData.toRecipients?.map(x => x.emailAddress?.address).filter(Boolean) || finalPayload.to;
      finalPayload.cc = msgData.ccRecipients?.map(x => x.emailAddress?.address).filter(Boolean) || finalPayload.cc;
      finalPayload.sentAt = msgData.sentDateTime || finalPayload.sentAt;
      finalPayload.rawMimeBase64 = mimeBase64;
      
      if (mimeBase64) {
        console.log("[fileService] MIME stream fetched successfully for MSG fidelity.");
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
      console.error("================== GRAPH ENRICHMENT FAILED ==================");
      console.error("[fileService] Graph enrichment failed — falling back to local payload.");
      console.error(`[fileService] Error: ${error.message}`);
      console.error(`[fileService] Token type in use: isAccessToken=${graphAuthOptions.isAccessToken}, hasToken=${!!graphAuthToken}`);
      console.error("==============================================================");
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
  const msgName = buildMsgFileName(finalPayload.subject, finalPayload.sentAt, finalPayload.sender);
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
      const msgPayload = shouldEmbedAttachments ? finalPayload : { ...finalPayload, attachments: [], rawMimeBase64: null };
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

    const attachmentPaths = shouldWriteSeparateAttachments
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
      hasAttachments: finalPayload.hasAttachments !== undefined ? finalPayload.hasAttachments : (Array.isArray(finalPayload.attachments) && finalPayload.attachments.length > 0),
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
      // Small pacing delay between Graph actions to avoid hitting rate limits (HTTP 429)
      // when multiple emails are filed in rapid succession.
      const graphPause = () => new Promise(r => setTimeout(r, 250));

      // 1. Mark as reviewed
      if (graphAuthToken && finalPayload.itemId && finalPayload.markReviewed) {
        try {
          await withGraphAuthFallback((token, options) =>
            graphService.markEmailReviewed(token, finalPayload.itemId, options)
          );
          
          // Append reviewed indicator to the subject
          const reviewedSubject = `[Reviewed] ${finalPayload.subject || ''}`;
          await graphPause();
          await graphService.updateEmailSubject(graphAuthToken, finalPayload.itemId, reviewedSubject, graphAuthOptions);
          console.log(`[fileService] Marked as read and updated subject to: ${reviewedSubject}`);
        } catch (err) {
          appendPostFilingError(`[FS-POST-FAIL] Mark as reviewed: ${err.message}`);
          console.error("[fileService] [FS-POST-FAIL]", err.message);
        }
      }

      // 2. Add "Filed" category BEFORE any move/archive (moving changes the item ID)
      if (graphAuthToken && finalPayload.itemId && finalPayload.addFiledCategory) {
        try {
          await graphPause();
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
            await graphPause();
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
          await graphPause();
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
          await graphPause();
          if (finalPayload.afterFiling === "delete" || finalPayload.afterFiling === "move_deleted") {
            await graphService.deleteEmail(graphAuthToken, finalPayload.itemId, graphAuthOptions);
          } else if (finalPayload.afterFiling === "archive") {
            await graphService.archiveEmail(graphAuthToken, finalPayload.itemId, graphAuthOptions);
          } else if (finalPayload.afterFiling === "move_filed_items") {
            const folderId = await graphService.getOrCreateMailFolder(graphAuthToken, 'inbox', 'Filed Items', graphAuthOptions);
            await graphPause();
            await graphService.moveEmail(graphAuthToken, finalPayload.itemId, folderId, graphAuthOptions);
          } else if (finalPayload.afterFiling === "move_filed_folders") {
            const prefix = finalPayload.filedFolderPrefix || '*';
            const locationName = targets.length > 0 ? targets[0].split(/[\\\/]/).filter(Boolean).pop() : 'Filed';
            const folderName = `${prefix} ${locationName}`.trim();
            const folderId = await graphService.getOrCreateMailFolder(graphAuthToken, 'inbox', folderName, graphAuthOptions);
            await graphPause();
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


    // ── On-Send: tag the Sent Items copy via Graph (delayed) ────────────────
    // When isOnSend is true the compose item was frozen during ItemSend so
    // OfficeJS item.categories calls fail with code 5000.  We instead wait a
    // few seconds for Outlook to deliver the message to Sent Items, then
    // locate it by subject and apply the category / subject prefix via Graph.
    if (finalPayload.isOnSend && finalPayload.ssoToken && finalPayload.subject) {
      const onSendSubject = finalPayload.subject;
      const onSendToken = finalPayload.ssoToken;

      // Read options that were embedded in the payload or fall back to defaults.
      const addCat = finalPayload.addFiledCategory !== false;
      const catName = finalPayload.filedCategoryName || "Filed by mailmanager (koyomail)";
      const markReviewedFlag = !!finalPayload.markReviewed;
      const afterFilingAction = finalPayload.afterFiling || "none";
      const useUtcTime = !!finalPayload.useUtcTime;

      const DELAY_MS = 7000; // 7 s — enough time for Outlook to place the message in Sent Items
      console.log(`[fileService] On-Send: scheduling background Sent-Items tagging in ${DELAY_MS}ms for subject: "${onSendSubject}"`);

      setTimeout(async () => {
        try {
          console.log(`[fileService] On-Send background task running — searching Sent Items for: "${onSendSubject}"`);

          // ── Resolve which token mode to use ─────────────────────────────────
          // The token from the dialog can be:
          //   a) A direct Graph access token (from NAA / MSAL in New Outlook)
          //      → must be used with { isAccessToken: true }
          //   b) An Office SSO identity token
          //      → can be exchanged via OBO by the backend
          //
          // We try direct access token first (most common in New Outlook).
          // If that fails with an auth error we retry via OBO.
          // ────────────────────────────────────────────────────────────────────
          const isAuthError = (err) => {
            const msg = String(err?.message || "").toLowerCase();
            return msg.includes("401") || msg.includes("unauthorized") ||
                   msg.includes("invalidauthenticationtoken") || msg.includes("access token");
          };

          let resolvedToken = onSendToken;
          let resolvedOptions = { isAccessToken: true }; // Try as direct Graph token first

          console.log(`[fileService] On-Send: attempting with direct Graph access token (isAccessToken=true)...`);

          // Helper to search with auto-retry on first-attempt empty result
          const searchWithRetry = async (token, opts) => {
            let msg = await graphService.searchSentMessage(token, onSendSubject, opts);
            if (!msg) {
              console.warn(`[fileService] On-Send: message not found yet — retrying in 8s`);
              await new Promise(r => setTimeout(r, 8000));
              msg = await graphService.searchSentMessage(token, onSendSubject, opts);
            }
            return msg;
          };

          let sentMsgToUse = null;
          try {
            sentMsgToUse = await searchWithRetry(resolvedToken, resolvedOptions);
          } catch (directErr) {
            // Direct token failed — may be an Office SSO identity token; try OBO
            console.warn(`[fileService] On-Send: direct token attempt failed (${directErr.message}). Retrying via OBO flow...`);
            resolvedOptions = {}; // OBO mode (isAccessToken: false)
            sentMsgToUse = await searchWithRetry(resolvedToken, resolvedOptions);
          }

          if (!sentMsgToUse) {
            console.warn(`[fileService] On-Send: could not find sent message with subject "${onSendSubject}" — giving up.`);
            return;
          }

          console.log(`[fileService] On-Send: found sent message id=${sentMsgToUse.id}`);

          // 1. Apply the "Filed" category
          if (addCat) {
            try {
              await graphService.addCategoryToEmail(resolvedToken, sentMsgToUse.id, catName, resolvedOptions);
              console.log(`[fileService] On-Send: applied category "${catName}" to sent message.`);
            } catch (catErr) {
              console.warn(`[fileService] On-Send: failed to add category: ${catErr.message}`);
            }
          }

          // 2. Apply subject prefix if markReviewed or add_date was requested
          let subjectPrefix = "";
          if (markReviewedFlag) subjectPrefix += "[Reviewed] ";
          if (afterFilingAction === "add_date") {
            const dateStr = useUtcTime
              ? new Date().toISOString().replace("T", " ").substring(0, 19) + " UTC"
              : new Date().toLocaleString();
            subjectPrefix += `[Filed ${dateStr}] `;
          }
          if (subjectPrefix && !onSendSubject.startsWith(subjectPrefix.trim())) {
            try {
              const newSubject = subjectPrefix + onSendSubject;
              await graphService.updateEmailSubject(resolvedToken, sentMsgToUse.id, newSubject, resolvedOptions);
              console.log(`[fileService] On-Send: updated sent message subject to "${newSubject}".`);
            } catch (subErr) {
              console.warn(`[fileService] On-Send: failed to update subject: ${subErr.message}`);
            }
          }

          // 3. Apply the after-filing action (delete, archive, move_filed_items, move_filed_folders)
          if (afterFilingAction && afterFilingAction !== "none" && afterFilingAction !== "add_date") {
            try {
              console.log(`[fileService] On-Send: processing after-filing action "${afterFilingAction}" on sent message id=${sentMsgToUse.id}...`);
              if (afterFilingAction === "delete" || afterFilingAction === "move_deleted") {
                await graphService.deleteEmail(resolvedToken, sentMsgToUse.id, resolvedOptions);
                console.log(`[fileService] On-Send: successfully moved sent message to Deleted Items.`);
              } else if (afterFilingAction === "archive") {
                await graphService.archiveEmail(resolvedToken, sentMsgToUse.id, resolvedOptions);
                console.log(`[fileService] On-Send: successfully moved sent message to Archive.`);
              } else if (afterFilingAction === "move_filed_items") {
                const folderId = await graphService.getOrCreateMailFolder(resolvedToken, 'inbox', 'Filed Items', resolvedOptions);
                await graphService.moveEmail(resolvedToken, sentMsgToUse.id, folderId, resolvedOptions);
                console.log(`[fileService] On-Send: successfully moved sent message to Filed Items folder.`);
              } else if (afterFilingAction === "move_filed_folders") {
                const prefix = finalPayload.filedFolderPrefix || '*';
                const locationName = targets.length > 0 ? targets[0].split(/[\\/]/).filter(Boolean).pop() : 'Filed';
                const folderName = `${prefix} ${locationName}`.trim();
                const folderId = await graphService.getOrCreateMailFolder(resolvedToken, 'inbox', folderName, resolvedOptions);
                await graphService.moveEmail(resolvedToken, sentMsgToUse.id, folderId, resolvedOptions);
                console.log(`[fileService] On-Send: successfully moved sent message to "${folderName}" folder.`);
                
                if (finalPayload.deleteEmptyFolders) {
                  await graphService.cleanupEmptyFolders(resolvedToken, 'inbox', prefix, resolvedOptions);
                }
              }
            } catch (actionErr) {
              console.warn(`[fileService] On-Send: after-filing action "${afterFilingAction}" failed: ${actionErr.message}`);
            }
          }

          console.log(`[fileService] On-Send background tagging complete.`);
        } catch (bgErr) {
          console.error(`[fileService] On-Send background tagging failed: ${bgErr.message}`);
        }
      }, DELAY_MS);
    }
    // ────────────────────────────────────────────────────────────────────────

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

    // If "Generate email link" was requested and we have a Graph token, create a draft email
    // with the links so the user can add recipients and send it.
    let draftEmailCreated = false;
    let draftId = null;
    let webLink = null;
    if (finalPayload.sendLink && !finalPayload.skipDraftCreation && sharingLinks.length > 0 && graphAuthToken && graphEnrichmentSucceeded) {
      try {
        console.log(`[fileService] Creating draft email with filing links...`);
        const draftResult = await graphService.createDraftLinkEmail(graphAuthToken, {
          filedEntries: sharingLinks,
          originalSubject: finalPayload.subject,
          comment: finalPayload.comment,
          fontFamily: finalPayload.emailFont || "Segoe UI",
          fontSize: finalPayload.fontSize ? `${finalPayload.fontSize}pt` : "11pt",
        }, graphAuthOptions);
        draftEmailCreated = true;
        draftId = draftResult?.draftId || null;
        webLink = draftResult?.webLink || null;
        console.log(`[fileService] Draft email with filing links created successfully. ID: ${draftId}`);
      } catch (draftErr) {
        console.warn(`[fileService] Failed to create draft email with filing links: ${draftErr.message}`);
        appendPostFilingError(`Generate email link: Could not create draft email — ${draftErr.message}. Links: ${sharingLinks.join(", ")}`);
      }
    }

    return {
      fileName: firstSavedPath ? path.basename(firstSavedPath) : msgName,
      filedAt,
      results: perTarget,
      postFilingError,
      sharingLinks,
      draftEmailCreated,
      draftId,
      webLink,
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

export async function createConsolidatedDraft(payload) {
  const { graphAccessToken, ssoToken, filedEntries, originalSubject, comment, emailFont, fontSize } = payload;
  
  const normalizedAccessToken = typeof graphAccessToken === "string" ? graphAccessToken.trim() : "";
  const normalizedSsoToken = typeof ssoToken === "string" ? ssoToken.trim() : "";

  const graphAuthToken = normalizedSsoToken || normalizedAccessToken || null;
  const graphAuthOptions = { isAccessToken: !normalizedSsoToken && !!normalizedAccessToken };
  
  if (!graphAuthToken) {
    throw new Error("No authentication token available for creating draft email.");
  }

  return await graphService.createDraftLinkEmail(graphAuthToken, {
    filedEntries,
    originalSubject: originalSubject || "Multiple Emails",
    comment,
    fontFamily: emailFont || "Segoe UI",
    fontSize: fontSize ? `${fontSize}pt` : "11pt"
  }, graphAuthOptions);
}
