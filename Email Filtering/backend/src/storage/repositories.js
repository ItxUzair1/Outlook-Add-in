import path from "path";
import { config } from "../config/index.js";
import { readJson, writeJson } from "./jsonStore.js";

const locationsPath = path.join(config.dataDir, "locations.json");
const searchIndexPath = path.join(config.dataDir, "search-index.json");
const senderFavouritesPath = path.join(config.dataDir, "sender-favourites.json");

export async function getLocations() {
  return readJson(locationsPath, []);
}

export async function saveLocations(data) {
  return writeJson(locationsPath, data);
}

export async function getSearchIndex() {
  return readJson(searchIndexPath, []);
}

export async function saveSearchIndex(data) {
  return writeJson(searchIndexPath, data, { compact: true });
}

export async function getSenderFavouritesStore() {
  return readJson(senderFavouritesPath, {});
}

export async function saveSenderFavouritesStore(data) {
  return writeJson(senderFavouritesPath, data);
}
