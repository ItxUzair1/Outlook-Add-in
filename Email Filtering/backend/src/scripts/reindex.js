import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { getSearchIndex, saveSearchIndex } from "../storage/repositories.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Basic EML body extractor. 
 * Skips headers and returns the first 10,000 characters of text/html or text/plain parts.
 */
function extractBodyFromEml(content) {
    const splitIndex = content.indexOf("\r\n\r\n");
    const bodyPart = splitIndex !== -1 ? content.slice(splitIndex + 4) : content;
    
    // Very basic MIME cleaning: remove boundaries and headers within multiparts
    // This is "dirty" but works for search indexing!
    return bodyPart
        .replace(/Content-Type:.*?\r\n/gi, "")
        .replace(/Content-Transfer-Encoding:.*?\r\n/gi, "")
        .replace(/Content-Disposition:.*?\r\n/gi, "")
        .replace(/--[a-zA-Z0-9'()+ ,./?=_]{0,70}/g, "") // remove boundaries
        .replace(/<[^>]*>/g, " ") // remove HTML tags
        .replace(/\s+/g, " ") // normalize whitespace
        .trim()
        .slice(0, 10000);
}

/**
 * Best-effort MSG text scraper.
 * Since MSG is a binary format (OLE2), we look for large chunks of printable ASCII/UTF-16 characters.
 */
function extractBodyFromMsg(buffer) {
    // Look for ASCII strings of at least 10 human-readable characters
    const matches = buffer.toString("utf-8").match(/[a-zA-Z0-9\s.,!?-]{20,}/g);
    if (!matches) return "";
    
    // Most emails will have some headers like "From:" or "Subject:" first.
    // The longest match is usually the body.
    const longestMatch = matches.sort((a, b) => b.length - a.length)[0];
    return longestMatch.trim().slice(0, 10000);
}

async function runRepair() {
    console.log("🚀 Starting Search Index Repair (Body Extraction)...");
    
    const index = await getSearchIndex();
    let updatedCount = 0;
    let failCount = 0;
    let skipCount = 0;

    for (let i = 0; i < index.length; i++) {
        const record = index[i];
        
        // Skip if it already has a body
        if (record.body && record.body.trim().length > 0) {
            skipCount++;
            continue;
        }

        const fullPath = record.filePath;
        if (!fullPath) {
            skipCount++;
            continue;
        }

        try {
            const ext = path.extname(fullPath).toLowerCase();
            const stats = await fs.stat(fullPath);
            
            // Don't try to index huge files
            if (stats.size > 10 * 1024 * 1024) {
               console.warn(`[SKIP] Too large: ${path.basename(fullPath)}`);
               skipCount++;
               continue;
            }

            if (ext === ".eml") {
                const content = await fs.readFile(fullPath, "utf-8");
                record.body = extractBodyFromEml(content);
                updatedCount++;
                console.log(`[OK] Indexed EML: ${path.basename(fullPath)}`);
            } 
            else if (ext === ".msg") {
                const buffer = await fs.readFile(fullPath);
                record.body = extractBodyFromMsg(buffer);
                updatedCount++;
                console.log(`[OK] Indexed MSG: ${path.basename(fullPath)}`);
            }
            else {
                // It might be a regular attachment file
                record.body = `[File Attachment] ${path.basename(fullPath)}`;
                updatedCount++;
                skipCount++;
            }

        } catch (err) {
            console.error(`[FAIL] Could not read ${fullPath}: ${err.message}`);
            failCount++;
        }
    }

    if (updatedCount > 0) {
        await saveSearchIndex(index);
        console.log(`\n✅ REPAIR COMPLETE!`);
        console.log(`- Updated: ${updatedCount} records`);
        console.log(`- Skipped: ${skipCount} records`);
        console.log(`- Failed:  ${failCount} records`);
    } else {
        console.log("\n✨ No records needed repair.");
    }
}

runRepair().catch(err => {
    console.error("Fatal error during repair:", err);
});
