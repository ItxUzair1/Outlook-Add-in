/* global Office */
import { fileEmail, applyPostFilingActions } from "./backendApi.js";
import { applyPostFilingByItemId } from "../utils/afterFilingUtils.js";
import {
  applyCategoryForFilingOutcome,
  applyPostActionFailureCategoryByItemId,
  wasEmailSavedToDisk,
  DEFAULT_FAILURE_CATEGORY_NAME,
  DEFAULT_POST_ACTION_FAILURE_CATEGORY_NAME,
  shouldApplyPostActionFailureCategory,
} from "../utils/filingCategoryUtils.js";
import { openFilingLinksCompose } from "../utils/filingCompose.js";

let queueChain = Promise.resolve();
let pendingCount = 0;

/**
 * Sequential filing queue — if email B is filed while email A is still processing,
 * B waits until A completes (FIFO).
 */
export function enqueueFilingJob(job) {
  const wasQueued = pendingCount > 0;
  pendingCount += 1;

  if (wasQueued && job.meta) {
    job.meta.wasQueued = true;
  }

  const run = queueChain
    .then(() => processFilingJob(job))
    .catch((err) => {
      console.error("[filingQueue] Unhandled job error:", err);
    })
    .finally(() => {
      pendingCount = Math.max(0, pendingCount - 1);
    });

  queueChain = run;
  return run;
}

export function getPendingFilingCount() {
  return pendingCount;
}

function resolveTargetFolderName(payload) {
  const paths = Array.isArray(payload?.targetPaths) ? payload.targetPaths : [];
  if (paths.length === 0) return "Filed";
  return String(paths[0]).split(/[\\/]/).filter(Boolean).pop() || "Filed";
}

function needsPostFilingActions(payload, meta) {
  const afterFiling = meta.afterFiling || payload.afterFiling || "none";
  const markReviewed = meta.markReviewed ?? payload.markReviewed ?? false;
  const addFiledCategory = meta.addFiledCategory ?? payload.addFiledCategory ?? false;
  return (
    (afterFiling && afterFiling !== "none") ||
    markReviewed ||
    addFiledCategory
  );
}

/**
 * Background post-filing always targets the filed email's item ID (never mailbox.item).
 */
async function ensureBackgroundPostFiling(payload, meta, response, { skipCategory = false } = {}) {
  const itemId = payload.ewsItemId || payload.itemId;
  if (!itemId || !needsPostFilingActions(payload, meta)) {
    return { handled: !response?.postFilingError, detail: null };
  }

  const afterFiling = meta.afterFiling || payload.afterFiling || "none";
  const addFiledCategory = meta.addFiledCategory ?? payload.addFiledCategory ?? false;
  const filedCategoryName = meta.filedCategoryName || payload.filedCategoryName || "Filed by Koyomail";
  const graphFailed = !!response?.postFilingError;
  const needsMove = afterFiling && afterFiling !== "none" && afterFiling !== "add_date";

  const clientResult = await applyPostFilingByItemId({
    itemId,
    afterFiling: needsMove || graphFailed ? afterFiling : "none",
    addFiledCategory: !skipCategory && addFiledCategory,
    filedCategoryName,
    targetFolderName: resolveTargetFolderName(payload),
    filedFolderPrefix: payload.filedFolderPrefix || "*",
  });

  if (clientResult.handled) {
    return { handled: true, detail: clientResult };
  }

  if ((graphFailed || needsMove) && (payload.graphAccessToken || payload.ssoToken)) {
    try {
      await applyPostFilingActions({
        itemId: payload.itemId,
        graphAccessToken: payload.graphAccessToken,
        ssoToken: payload.ssoToken,
        afterFiling,
        markReviewed: meta.markReviewed ?? payload.markReviewed,
        addFiledCategory,
        filedCategoryName,
        masterCategoryEnsured: payload.masterCategoryEnsured,
        assistantCategories: payload.assistantCategories,
        useUtcTime: payload.useUtcTime,
        filedFolderPrefix: payload.filedFolderPrefix,
        deleteEmptyFolders: payload.deleteEmptyFolders,
        targetPaths: payload.targetPaths,
        subject: payload.subject || meta.subject,
      });
      return { handled: true, detail: { completed: ["graph-retry"] } };
    } catch (retryErr) {
      const detail = clientResult.failures?.length
        ? `${clientResult.failures.join("; ")} | Graph retry: ${retryErr.message}`
        : retryErr.message;
      return { handled: false, detail };
    }
  }

  if (!graphFailed && !needsMove) {
    return { handled: true, detail: null };
  }

  return {
    handled: false,
    detail: clientResult.failures?.join("; ") || response?.postFilingError || "Post-filing action could not be completed.",
  };
}

async function processFilingJob(job) {
  const { payload, meta = {} } = job;
  const subject = meta.subject || payload.subject || "Email";
  const itemId = payload.ewsItemId || payload.itemId;
  const addFiledCategory = meta.addFiledCategory ?? payload.addFiledCategory ?? false;
  const filedCategoryName = meta.filedCategoryName || payload.filedCategoryName || "Filed by Koyomail";
  const failureCategoryName = payload.failureCategoryName || DEFAULT_FAILURE_CATEGORY_NAME;

  if (meta.wasQueued) {
    console.log(`[filingQueue] Queued filing started: "${subject}"`);
  }

  try {
    const response = await fileEmail(payload);

    if (!wasEmailSavedToDisk(response)) {
      await applyCategoryForFilingOutcome({
        itemId,
        response,
        addFiledCategory: true,
        filedCategoryName,
        failureCategoryName,
        filingError: "not_saved",
      });
      console.error(`[filingQueue] Filing failed for "${subject}": email was not saved to disk.`);
      return;
    }

    // Filed successfully — apply success category first (before post-actions).
    if (addFiledCategory && itemId) {
      await applyCategoryForFilingOutcome({
        itemId,
        response,
        addFiledCategory: true,
        filedCategoryName,
        failureCategoryName,
      });
    }

    const postFilingOutcome = await ensureBackgroundPostFiling(payload, meta, response, {
      skipCategory: true,
    });
    const postFilingHandled = postFilingOutcome.handled;
    const afterFiling = meta.afterFiling || payload.afterFiling || "none";
    const markReviewed = meta.markReviewed ?? payload.markReviewed ?? false;
    const postActionFailureCategoryName =
      payload.postActionFailureCategoryName || DEFAULT_POST_ACTION_FAILURE_CATEGORY_NAME;

    if (
      shouldApplyPostActionFailureCategory({
        saved: true,
        postFilingHandled,
        afterFiling,
        markReviewed,
      }) &&
      itemId
    ) {
      await applyPostActionFailureCategoryByItemId(itemId, postActionFailureCategoryName);
      const errDetail = postFilingOutcome.detail ||
        response?.postFilingError ||
        "Archive/delete/move could not be completed automatically.";
      console.error(
        `[filingQueue] Post-action failed for "${subject}" (email was filed):`,
        errDetail
      );
    }

    const sendLink = meta.sendLink ?? payload.sendLink;

    if (sendLink && response?.sharingLinks?.length > 0) {
      const linkText = response.sharingLinks.join("\n");
      try {
        await navigator.clipboard.writeText(linkText);
      } catch {
        /* clipboard optional */
      }

      if (response.draftEmailCreated && response.draftId) {
        try {
          Office.context.mailbox.displayMessageForm(response.draftId);
        } catch (openErr) {
          console.warn("[filingQueue] Could not open draft:", openErr.message);
        }
      } else {
        openFilingLinksCompose(response.sharingLinks, subject, {
          emailFont: meta.emailFont || payload.emailFont,
          fontSize: meta.fontSize || payload.fontSize,
        });
      }
    }

    console.log(`[filingQueue] Filed successfully: "${subject}"`);
  } catch (err) {
    const errMsg = err instanceof Error ? err.message : String(err);
    await applyCategoryForFilingOutcome({
      itemId,
      response: null,
      addFiledCategory: true,
      filedCategoryName,
      failureCategoryName,
      filingError: errMsg,
    });
    console.error(`[filingQueue] Filing failed for "${subject}":`, errMsg);
    throw err;
  }
}
