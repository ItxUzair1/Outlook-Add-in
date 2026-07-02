const { parentPort } = require('worker_threads');
const { parseEmailFile } = require('./parser');

parentPort.on('message', async ({ id, filePath }) => {
  try {
    const result = await parseEmailFile(filePath);
    parentPort.postMessage({ id, result });
  } catch (err) {
    parentPort.postMessage({ id, error: err.message || String(err) });
  }
});
