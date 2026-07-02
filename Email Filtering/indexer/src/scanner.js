const fs = require('fs');
const path = require('path');

/**
 * Recursively scans a directory for .msg and .eml files.
 * Uses readdir({ withFileTypes: true }) to avoid an extra stat() syscall
 * per entry — saves ~1 disk read per file vs the old implementation.
 * @param {string} dir Absolute path to scan
 * @yields {string} Absolute file path
 */
async function* scanDirectory(dir) {
  let entries;
  try {
    // withFileTypes gives us Dirent objects that already know isDirectory()
    // without a separate stat() call — much faster on large trees.
    entries = await fs.promises.readdir(dir, { withFileTypes: true });
  } catch (err) {
    console.warn(`[Scanner] Cannot read directory ${dir}: ${err.message}`);
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    try {
      if (entry.isDirectory()) {
        yield* scanDirectory(fullPath);
      } else if (entry.isFile()) {
        const ext = entry.name.toLowerCase();
        if (ext.endsWith('.msg') || ext.endsWith('.eml')) {
          yield fullPath;
        }
      }
    } catch (err) {
      console.warn(`[Scanner] Skipping inaccessible path: ${fullPath} — ${err.message}`);
    }
  }
}

module.exports = { scanDirectory };
