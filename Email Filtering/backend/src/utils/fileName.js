const INVALID_CHARS_REGEX = /[<>:"/\\|?*\x00-\x1f]/g;

export function sanitizeFileName(value) {
  const normalized = String(value || "No Subject").replace(INVALID_CHARS_REGEX, "_").replace(/\s+/g, " ").trim();
  const cleaned = normalized.replace(/[.\s]+$/g, "");
  return cleaned || "No Subject";
}

export function buildMsgFileName(subject, sentAt, sender, senderName) {
  const date = sentAt ? new Date(sentAt) : new Date();

  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");

  // Extract the display name from the sender (strip <email@address> part if present)
  let senderPart = "";
  let rawName = "";

  if (senderName) {
    rawName = String(senderName).trim();
  } else if (sender) {
    const nameMatch = String(sender).match(/^([^<]+)<[^>]+>/);
    rawName = nameMatch ? nameMatch[1].trim() : String(sender).trim();
  }

  if (rawName) {
    if (!rawName.includes("@")) {
      senderPart = `${sanitizeFileName(rawName)}_`;
    } else {
      // It's just an email address — use the part before @
      senderPart = `${sanitizeFileName(rawName.split("@")[0])}_`;
    }
  }

  const cleanSubject = sanitizeFileName(subject);
  // Truncate the subject to 100 characters to prevent hitting Windows MAX_PATH (260 chars) limits
  const truncatedSubject = cleanSubject.length > 100 ? cleanSubject.substring(0, 100) + "..." : cleanSubject;

  return `${yyyy}${mm}${dd}_${hh}${mi}${ss}_${senderPart}${truncatedSubject}.eml`;
}
