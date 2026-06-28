import path from "path";
import { config } from "../config/index.js";
import { readJson, writeJson } from "./jsonStore.js";

const locationsPath = path.join(config.dataDir, "locations.json");

const senderFavouritesPath = path.join(config.dataDir, "sender-favourites.json");

export async function getLocations() {
  return readJson(locationsPath, []);
}

export async function saveLocations(data) {
  return writeJson(locationsPath, data);
}



export async function getSenderFavouritesStore() {
  return readJson(senderFavouritesPath, {});
}

export async function saveSenderFavouritesStore(data) {
  return writeJson(senderFavouritesPath, data);
}
