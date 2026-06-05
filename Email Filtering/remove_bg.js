const sharp = require('sharp');

async function removeWhiteBackground() {
  // Step 1: Read the original JPEG as raw RGBA pixels
  const { data, info } = await sharp('assets/new_logo.jpeg')
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });

  const { width, height, channels } = info;
  console.log(`Image: ${width}x${height}, channels: ${channels}`);

  // Step 2: Aggressively remove white/near-white background
  // Use a generous threshold since JPEG compression creates artifacts
  const THRESHOLD = 220; // Anything with R,G,B all above 220 is treated as background

  for (let i = 0; i < data.length; i += channels) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];

    if (r > THRESHOLD && g > THRESHOLD && b > THRESHOLD) {
      // Make fully transparent
      data[i + 3] = 0;
    }
  }

  // Step 3: Convert raw pixels back to PNG
  const pngBuffer = await sharp(data, {
    raw: { width, height, channels }
  }).png().toBuffer();

  // Step 4: Trim transparent edges, then pad to a square with transparent bg
  const trimmed = await sharp(pngBuffer).trim().toBuffer();
  const trimMeta = await sharp(trimmed).metadata();
  console.log(`After trim: ${trimMeta.width}x${trimMeta.height}`);

  // Make it a nice large square for the manifest (800x800)
  const size = 800;
  await sharp(trimmed)
    .resize(size, size, {
      fit: 'contain',
      background: { r: 0, g: 0, b: 0, alpha: 0 }
    })
    .png()
    .toFile('assets/new_logo_transparent.png');

  console.log('Done! Saved assets/new_logo_transparent.png');
}

removeWhiteBackground().catch(console.error);
