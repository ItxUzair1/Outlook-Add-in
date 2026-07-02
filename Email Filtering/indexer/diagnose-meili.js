#!/usr/bin/env node
/**
 * Standalone Meilisearch connection check.
 *
 * Usage (from indexer folder):
 *   node diagnose-meili.js
 *   npm run diagnose
 *
 * While KoyoIndexer.exe is running, you can also use:
 *   curl http://localhost:4001/api/diagnostics
 */

const path = require('path');
const pkg = require('./package.json');
const state = require('./src/state');
const {
  runMeiliDiagnostics,
  formatDiagnosticsForConsole,
} = require('./src/meiliDiagnostics');

async function main() {
  try {
    const report = await runMeiliDiagnostics({ state, pkg });
    console.log(formatDiagnosticsForConsole(report));
    process.exit(report.ok ? 0 : 1);
  } catch (err) {
    console.error('Diagnostics failed:', err.message);
    process.exit(1);
  }
}

main();
