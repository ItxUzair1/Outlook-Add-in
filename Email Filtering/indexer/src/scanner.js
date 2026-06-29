const fs = require('fs');
const path = require('path');

/**
 * Recursively scans a directory for .msg and .eml files.
 * @param {string} dir Absolute path to scan
 * @param {string[]} fileList Accumulator array
 * @returns {string[]} Array of absolute file paths
 */
async function scanDirectory(dir, fileList = []) {
  try {
    const stat = await fs.promises.stat(dir).catch(() => null);
    if (!stat) {
      console.warn(`Directory not found: ${dir}`);
      return fileList;
    }

    const files = await fs.promises.readdir(dir).catch(() => []);

    for (const file of files) {
      const fullPath = path.join(dir, file);
      try {
        const fileStat = await fs.promises.stat(fullPath);
        
        if (fileStat.isDirectory()) {
          await scanDirectory(fullPath, fileList);
        } else {
          const ext = path.extname(fullPath).toLowerCase();
          if (ext === '.msg' || ext === '.eml') {
            fileList.push(fullPath);
          }
        }
      } catch (err) {
        console.warn(`Skipping inaccessible path: ${fullPath} - ${err.message}`);
      }
    }
  } catch (err) {
    console.warn(`Error reading directory ${dir} - ${err.message}`);
  }

  return fileList;
}

module.exports = {
  scanDirectory
};
