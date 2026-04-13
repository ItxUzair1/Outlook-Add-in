/**
 * Rebuilds Koyomail-02-appicon.png and Koyomail-02-appicon-256.png from assets/Koyomail-02.png:
 * trims margins, pads to a transparent square, exports fixed sizes.
 * Run: npm run icons:koyomail
 */
import sharp from "sharp";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const assets = path.join(__dirname, "..", "assets");
const src = path.join(assets, "Koyomail-02.png");
const transparent = { r: 0, g: 0, b: 0, alpha: 0 };

async function buildPaddedSquare(size, outFile) {
  const trimmed = sharp(src).trim({ threshold: 30 });
  const meta = await trimmed.metadata();
  const w = meta.width;
  const h = meta.height;
  const side = Math.max(w, h);
  /* Tight padding so the mark reads large at 128px (toolbar / Outlook slots) */
  const margin = Math.max(4, Math.round(side * 0.06));
  const outer = side + 2 * margin;

  const left = margin + Math.floor((side - w) / 2);
  const right = outer - w - left;
  const top = margin + Math.floor((side - h) / 2);
  const bottom = outer - h - top;

  await trimmed
    .extend({
      top,
      bottom,
      left,
      right,
      background: transparent,
    })
    .resize(size, size)
    .png()
    .toFile(path.join(assets, outFile));
}

await buildPaddedSquare(128, "Koyomail-02-appicon.png");
await buildPaddedSquare(256, "Koyomail-02-appicon-256.png");
console.log("Wrote Koyomail-02-appicon.png and Koyomail-02-appicon-256.png (transparent, padded).");
