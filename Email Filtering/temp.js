const sharp = require('sharp');
async function processImg() {
  const trimmed = await sharp('assets/new_logo.jpeg').trim().toBuffer();
  await sharp(trimmed)
    .resize(720, 720, { fit: 'contain', background: { r: 255, g: 255, b: 255, alpha: 1 } })
    .extend({ top: 40, bottom: 40, left: 40, right: 40, background: { r: 255, g: 255, b: 255, alpha: 1 } })
    .toFile('assets/new_logo.png');
  console.log('Done');
}
processImg().catch(console.error);
