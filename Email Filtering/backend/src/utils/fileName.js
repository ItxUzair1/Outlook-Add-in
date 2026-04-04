const INVALID_CHARS_REGEX = /[<>:"/\\|?*\x00-\x1f]/g;

export function sanitizeFileName(value) {
  const normalized = String(value || "No Subject").replace(INVALID_CHARS_REGEX, "_").replace(/\s+/g, " ").trim();
  const cleaned = normalized.replace(/[.\s]+$/g, "");
  return cleaned || "No Subject";
}

export function buildMsgFileName(subject, sentAt) {
  const date = sentAt ? new Date(sentAt) : new Date();

  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");

  return `${yyyy}${mm}${dd}_${hh}${mi}${ss}_${sanitizeFileName(subject)}.eml`;
}
