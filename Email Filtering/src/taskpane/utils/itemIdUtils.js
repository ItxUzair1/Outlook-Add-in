/* global Office */

/**
 * Consolidated item ID conversion utilities.
 * Converts between different Outlook item ID formats.
 */

/**
 * Converts an item ID to REST format when possible.
 * This is needed when calling REST API endpoints or when mixing EWS and REST contexts.
 * @param {string} itemId - The item ID to convert
 * @returns {string} The converted REST ID, or the original ID if conversion fails
 */
export function toRestItemId(itemId) {
  try {
    const mailbox = Office?.context?.mailbox;
    if (mailbox?.convertToRestId && Office?.MailboxEnums?.RestVersion?.v2_0) {
      return mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    }
  } catch (error) {
    console.warn("[itemIdUtils] convertToRestId failed:", error.message);
  }
  return itemId;
}

/**
 * Converts an item ID to EWS format when possible.
 * This is needed for Exchange Web Services API calls.
 * @param {string} itemId - The item ID to convert
 * @returns {string} The converted EWS ID, or the original ID if conversion fails
 */
export function toEwsItemId(itemId) {
  try {
    const mailbox = Office?.context?.mailbox;
    if (mailbox?.convertToEwsId && Office?.MailboxEnums?.RestVersion?.v2_0) {
      return mailbox.convertToEwsId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    }
  } catch (error) {
    console.warn("[itemIdUtils] convertToEwsId failed:", error.message);
  }
  return itemId;
}

/**
 * Alias: Converts an item ID to Graph-compatible format (REST format).
 * @param {string} itemId - The item ID to convert
 * @returns {string} The converted ID
 */
export function toGraphItemId(itemId) {
  return toRestItemId(itemId);
}
