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

      for (const store of stores) {
        locations.push({
          id: store["@_id"],
          type: store.type,
          description: store.description,
          folder: store.folder,
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
      folder: loc.folder || ""
    }));

    const xmlObj = {
      mailmanager: {
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
