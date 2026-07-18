/**
 * platformUtils.js — Office platform detection helpers.
 *
 * Used to differentiate desktop (PC/Mac) from mobile (iOS/Android) so the
 * add-in can adapt its behaviour without duplicating code.
 */

/* global Office */

/**
 * Returns true when the add-in is running inside Outlook on iOS or Android.
 * Safe to call before Office.onReady — returns false if Office is not yet
 * initialised (safe fallback assumes desktop).
 */
export function isMobilePlatform() {
  try {
    const p = Office?.context?.platform;
    return (
      p === Office.PlatformType.iOS ||
      p === Office.PlatformType.Android
    );
  } catch {
    return false;
  }
}

/**
 * Returns true when the add-in is running on a Windows PC or Mac.
 * Falls back to true so that any unknown/unsupported host is treated as desktop.
 */
export function isDesktopPlatform() {
  try {
    const p = Office?.context?.platform;
    return (
      p === Office.PlatformType.PC ||
      p === Office.PlatformType.Mac ||
      p === Office.PlatformType.OfficeOnline
    );
  } catch {
    return true; // safe fallback — assume desktop
  }
}
