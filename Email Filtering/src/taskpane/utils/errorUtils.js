/**
 * Shared error message extraction utility.
 * Consolidates error handling across frontend components.
 * @param {string|object} value - The error value to convert
 * @param {string} fallback - Default message if extraction fails
 * @returns {string} Extracted error message
 */
export function toErrorMessage(value, fallback = "Unknown Error") {
  if (typeof value === "string" && value.trim()) {
    return value;
  }

  if (value && typeof value === "object") {
    if (typeof value.message === "string" && value.message.trim()) {
      return value.message;
    }

    if (typeof value.error === "string" && value.error.trim()) {
      return value.error;
    }

    try {
      const serialized = JSON.stringify(value);
      if (serialized && serialized !== "{}") {
        return serialized;
      }
    } catch {
      // Ignore serialization issues and use fallback
    }
  }

  return fallback;
}
