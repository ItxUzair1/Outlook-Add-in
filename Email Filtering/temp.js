const sharp = require('sharp');
async function processImg() {
  const { data, info } = await sharp('assets/new_logo.jpeg')
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });

  for (let i = 0; i < data.length; i += info.channels) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    // If it's close to white, make it transparent
    if (r > 240 && g > 240 && b > 240) {
      data[i + 3] = 0;
    }
  }

  // Convert raw back to PNG buffer
  const pngBuffer = await sharp(data, { raw: { width: info.width, height: info.height, channels: info.channels } }).png().toBuffer();
  
  // Now trim and resize
  const trimmed = await sharp(pngBuffer).trim().toBuffer();
  await sharp(trimmed)
    .resize(720, 720, { fit: 'contain', background: { r: 255, g: 255, b: 255, alpha: 0 } })
    .extend({ top: 40, bottom: 40, left: 40, right: 40, background: { r: 255, g: 255, b: 255, alpha: 0 } })
    .toFile('assets/new_logo_transparent.png');
  console.log('Done');
}
processImg().catch(console.error);
