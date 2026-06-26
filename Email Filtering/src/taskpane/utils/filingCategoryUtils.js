import { addCategoryViaEws } from "./afterFilingUtils.js";

/** Outlook category presets — dark green (success), dark orange (failure). */
export const SUCCESS_CATEGORY_COLOR = "Preset19";
export const FAILURE_CATEGORY_COLOR = "Preset16";
export const GRAPH_SUCCESS_COLOR = "preset19";
export const GRAPH_FAILURE_COLOR = "preset16";

export const DEFAULT_FAILURE_CATEGORY_NAME = "Filing failed - Koyomail";
export const DEFAULT_POST_ACTION_FAILURE_CATEGORY_NAME = "Filed but post-action failed - Koyomail";
export const POST_ACTION_FAILURE_MESSAGE = "Email filed successfully, but post-action failed.";

export function wasEmailSavedToDisk(response) {
  const results = response?.results;
  if (!Array.isArray(results) || results.length === 0) {
    return false;
  }
  return results.some(
    (r) => r.status === "saved" || r.status === "overwritten" || r.status === "skipped"
  );
}

export async function applySuccessCategoryByItemId(itemId, categoryName) {
  if (!itemId || !categoryName) {
    return false;
  }
  try {
    await addCategoryViaEws(itemId, categoryName);
    return true;
  } catch (err) {
    console.warn("[filingCategory] Could not apply success category:", err.message);
    return false;
  }
}

export async function applyPostActionFailureCategoryByItemId(
  itemId,
  categoryName = DEFAULT_POST_ACTION_FAILURE_CATEGORY_NAME
) {
  if (!itemId || !categoryName) {
    return false;
  }
  try {
    await addCategoryViaEws(itemId, categoryName);
    return true;
  } catch (err) {
    console.warn("[filingCategory] Could not apply post-action failure category:", err.message);
    return false;
  }
}

export function hadConfiguredPostFilingAction(afterFiling, markReviewed) {
  return (afterFiling && afterFiling !== "none") || !!markReviewed;
}

export function shouldApplyPostActionFailureCategory({
  saved,
  postFilingHandled,
  afterFiling,
  markReviewed,
}) {
  return !!saved && !postFilingHandled && hadConfiguredPostFilingAction(afterFiling, markReviewed);
}

export async function applyFailureCategoryByItemId(
  itemId,
  categoryName = DEFAULT_FAILURE_CATEGORY_NAME
) {
  if (!itemId || !categoryName) {
    return false;
  }
  try {
    await addCategoryViaEws(itemId, categoryName);
    return true;
  } catch (err) {
    console.warn("[filingCategory] Could not apply failure category:", err.message);
    return false;
  }
}

/**
 * Filed to disk → success category (post-action failure category applied separately).
 * Filing error → failure category only.
 */
export async function applyCategoryForFilingOutcome({
  itemId,
  response,
  addFiledCategory,
  filedCategoryName = "Filed by Koyomail",
  failureCategoryName = DEFAULT_FAILURE_CATEGORY_NAME,
  filingError = null,
}) {
  if (!itemId) {
    return { applied: "none" };
  }

  const saved = !filingError && wasEmailSavedToDisk(response);

  if (saved) {
    if (addFiledCategory) {
      await applySuccessCategoryByItemId(itemId, filedCategoryName);
      return { applied: "success" };
    }
    return { applied: "none" };
  }

  await applyFailureCategoryByItemId(itemId, failureCategoryName);
  return { applied: "failure" };
}
