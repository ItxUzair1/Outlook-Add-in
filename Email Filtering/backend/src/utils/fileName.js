const INVALID_CHARS_REGEX = /[<>:"/\\|?*\x00-\x1f]/g;

export function sanitizeFileName(value) {
  const normalized = String(value || "No Subject").replace(INVALID_CHARS_REGEX, "_").replace(/\s+/g, " ").trim();
  const cleaned = normalized.replace(/[.\s]+$/g, "");
  return cleaned || "No Subject";
}

export function buildMsgFileName(subject, sentAt, sender) {
  const date = sentAt ? new Date(sentAt) : new Date();

  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");

  // Extract the display name from the sender (strip <email@address> part if present)
  let senderPart = "";
  if (sender) {
    const nameMatch = String(sender).match(/^([^<]+)<[^>]+>/);
    const rawName = nameMatch ? nameMatch[1].trim() : String(sender).trim();
    // Remove any email-only strings (no @ symbol means it's a real display name)
    if (rawName && !rawName.includes("@")) {
      senderPart = `${sanitizeFileName(rawName)}_`;
    } else if (rawName && rawName.includes("@")) {
      // It's just an email address — use the part before @
      senderPart = `${sanitizeFileName(rawName.split("@")[0])}_`;
    }
  }

  return `${yyyy}${mm}${dd}_${hh}${mi}${ss}_${senderPart}${sanitizeFileName(subject)}.eml`;
}
