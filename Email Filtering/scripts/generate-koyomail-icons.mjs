/**
 * Rebuilds Koyomail brand icons from assets/Koyomail-01.png and assets/Koyomail-02.png:
 * trims margins safely and exports fixed sizes (128, 256, 512).
 * Run: npm run icons:koyomail
 */
import sharp from "sharp";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const assets = path.join(__dirname, "..", "assets");

async function buildPaddedSquare(srcName, size, outSuffix) {
  const srcPath = path.join(assets, `${srcName}.png`);
  const outFileName = `${srcName}-${outSuffix}.png`;
  
  console.log(`Processing ${srcName} -> ${outFileName} (${size}x${size})...`);
  
  // Use a very low threshold to avoid cutting into anti-aliased edges
  // Then use resize with fit: 'contain' to add the desired padding automatically.
  // We specify a slightly smaller 'canvas' for contain to create a safe margin.
  const marginSize = Math.round(size * 0.90); // Use 90% of the space, leaving 5% margin on each side

  await sharp(srcPath)
    .trim({ threshold: 10 }) 
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
    .resize(size, size) // Ensure exact dimensions
    .png()
    .toFile(path.join(assets, outFileName));
}

async function run() {
  const sizes = [128, 256, 512];
  const logos = ["Koyomail-01", "Koyomail-02"];

  for (const logo of logos) {
    for (const size of sizes) {
      await buildPaddedSquare(logo, size, `appicon-${size}`);
    }
  }
  
  // Copy for compatibility
  await sharp(path.join(assets, "Koyomail-02-appicon-128.png"))
    .toFile(path.join(assets, "Koyomail-02-appicon.png"));
  
  console.log("Done. Generated 128, 256, 512 variants with safe margins.");
}

run().catch(console.error);
