import sharp from "sharp";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const assets = path.join(__dirname, "..", "assets");

async function buildPaddedSquare(srcPath, size, outFileName) {
  console.log(`Processing -> ${outFileName} (${size}x${size})...`);
  
  const marginSize = Math.round(size * 0.90);

  await sharp(srcPath)
    .resize(marginSize, marginSize, {
      fit: "contain",
      background: { r: 0, g: 0, b: 0, alpha: 0 }
    })
    .extend({
      top: Math.floor((size - marginSize) / 2),
      bottom: Math.ceil((size - marginSize) / 2),
      left: Math.floor((size - marginSize) / 2),
      right: Math.ceil((size - marginSize) / 2),
      background: { r: 0, g: 0, b: 0, alpha: 0 }
    })
    .resize(size, size)
    .png()
    .toFile(path.join(assets, outFileName));
}

async function run() {
  const srcPath = path.join(assets, "Koyomail-05-removebg-preview.png");
  
  if (!fs.existsSync(srcPath)) {
    console.error("Source file not found: Koyomail-05-removebg-preview.png");
    return;
  }

  const sizes = [16, 32, 64, 80, 128, 256, 512];
  
  console.log("Generating standard Koyomail-02 app icons...");
  for (const size of sizes) {
    await buildPaddedSquare(srcPath, size, `Koyomail-02-appicon-${size}.png`);
  }
  
  // Default icon
  fs.copyFileSync(
    path.join(assets, `Koyomail-02-appicon-128.png`),
    path.join(assets, `Koyomail-02-appicon.png`)
  );

  console.log("Done.");
}

run().catch(console.error);
