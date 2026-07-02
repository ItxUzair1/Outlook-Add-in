const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

async function build() {
  console.log('Packaging Electron app...');
  try {
    const packagerModule = await import('@electron/packager');
    const packager = packagerModule.packager || packagerModule.default || packagerModule;
    const releaseDir = path.join(__dirname, 'release');
    const appPaths = await packager({
      dir: '.',
      name: 'KoyoIndexer',
      platform: 'win32',
      arch: 'x64',
      out: releaseDir,
      icon: path.join(__dirname, 'icon.ico'),
      overwrite: true,
      ignore: [
        /^\/src\/test-scan\.js$/,
        /^\/admin-dashboard$/,
        /^\/dist/,
        /^\/release/
      ],
      asar: true
    });

    const appDir = appPaths[0];
    console.log(`App packaged successfully at ${appDir}`);

    const distDir = path.join(__dirname, 'dist');
    if (!fs.existsSync(distDir)) fs.mkdirSync(distDir, { recursive: true });

    // Create a portable ZIP instead of a Squirrel installer
    const zipOut = path.join(distDir, 'KoyoIndexer-portable.zip');
    console.log('Creating portable ZIP...');
    execSync(
      `powershell -Command "Compress-Archive -Path '${appDir}\\*' -DestinationPath '${zipOut}' -Force"`,
      { stdio: 'inherit' }
    );
    console.log(`\nDone! Portable ZIP created at: ${zipOut}`);
    console.log(`You can also run directly: ${appDir}\\KoyoIndexer.exe`);

  } catch (err) {
    console.error('Build failed!', err);
    process.exit(1);
  }
}

build();

