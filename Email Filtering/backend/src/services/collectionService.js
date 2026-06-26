import fs from "fs/promises";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

/**
 * Parses an .mmcollection XML file and returns an array of location objects.
 * @param {string} filePath Absolute path to the .mmcollection file
 * @returns {Promise<Array>} Array of location objects
 */
export async function loadCollectionFile(filePath) {
  try {
    const xmlData = await fs.readFile(filePath, "utf-8");
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_"
    });
    const result = parser.parse(xmlData);

    const locations = [];
    if (result?.mailmanager?.locations?.store) {
      // It can be an array of stores or a single store object
      const stores = Array.isArray(result.mailmanager.locations.store) 
        ? result.mailmanager.locations.store 
        : [result.mailmanager.locations.store];

      const getStr = (val) => {
        if (val === null || val === undefined) return "";
        if (typeof val === "object") return ""; // fast-xml-parser parses empty XML elements as {}
        return String(val);
      };

      for (const store of stores) {
        // fast-xml-parser returns "" for bare boolean attributes (e.g. <store isUnused>),
        // "true" for explicit ="true", and true (boolean) in some configs.
        // Operator precedence: parens are required so the bare-attr check is self-contained.
        const parseBoolAttr = (val) =>
          val === true || val === "true" || (val === "" && val !== undefined);

        locations.push({
          id: getStr(store["@_id"]),
          type: getStr(store.type),
          description: getStr(store.description),
          folder: getStr(store.folder),
          // Read collection field if the XML contains one (future-proofing)
          collection: store.collection ? getStr(store.collection) : undefined,
          isSuggested: parseBoolAttr(store["@_isSuggested"]),
          isUnused: parseBoolAttr(store["@_isUnused"]),
        });
      }
    }

    return locations;
  } catch (error) {
    throw new Error(`Failed to load collection file: ${error.message}`);
  }
}

/**
 * Saves an array of location objects back to an .mmcollection XML file.
 * @param {string} filePath Absolute path to the .mmcollection file
 * @param {Array} locations Array of location objects
 */
export async function saveCollectionFile(filePath, locations) {
  try {
    const stores = locations.map(loc => ({
      "@_id": loc.id,
      type: loc.type || "msg",
      description: loc.description || "",
      folder: loc.folder || loc.path || "",
      "@_isSuggested": loc.isSuggested ? "true" : undefined,
      "@_isUnused": loc.isUnused ? "true" : undefined
    }));

    const xmlObj = {
      mailmanager: {
        "@_xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        locations: {
          store: stores
        }
      }
    };

    const builder = new XMLBuilder({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      format: true
    });

    const xmlData = builder.build(xmlObj);
    
    // Write XML with standard declaration
    await fs.writeFile(filePath, `<?xml version="1.0" encoding="utf-8"?>\n${xmlData}`, "utf-8");
  } catch (error) {
    throw new Error(`Failed to save collection file: ${error.message}`);
  }
}
