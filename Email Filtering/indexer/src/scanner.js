const fs = require('fs');
const path = require('path');

/**
 * Recursively scans a directory for .msg and .eml files.
 * @param {string} dir Absolute path to scan
 * @param {string[]} fileList Accumulator array
 * @returns {string[]} Array of absolute file paths
 */
function scanDirectory(dir, fileList = []) {
  if (!fs.existsSync(dir)) {
    console.warn(`Directory not found: ${dir}`);
    return fileList;
  }

  const files = fs.readdirSync(dir);

  for (const file of files) {
    const fullPath = path.join(dir, file);
    try {
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        scanDirectory(fullPath, fileList);
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

  return fileList;
}

module.exports = {
  scanDirectory
};
