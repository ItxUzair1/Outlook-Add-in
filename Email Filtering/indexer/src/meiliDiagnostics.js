const fs = require('fs');
const path = require('path');
const { MeiliSearch } = require('meilisearch');

const ENV_PATH = path.join(__dirname, '..', '.env');
const INDEX_NAME = 'emails';

function loadEnv() {
  require('dotenv').config({ path: ENV_PATH });
}

function getMeiliConfig() {
  loadEnv();

  const meiliUrl = process.env.MEILI_URL || 'http://localhost:7700';
  const hasApiKey = Boolean(process.env.MEILI_MASTER_KEY);
  const usingFallback = !process.env.MEILI_URL;

  let meiliHost = meiliUrl;
  try {
    meiliHost = new URL(meiliUrl).host;
  } catch {
    // keep raw value if URL parsing fails
  }

  return {
    meiliUrl,
    meiliHost,
    hasApiKey,
    usingFallback,
    envFileFound: fs.existsSync(ENV_PATH),
    envFilePath: ENV_PATH,
  };
}

async function runMeiliDiagnostics(options = {}) {
  const { state = null, pkg = null } = options;
  const config = getMeiliConfig();

  const report = {
    ok: false,
    timestamp: new Date().toISOString(),
    version: pkg?.version || null,
    meilisearch: {
      configuredHost: config.meiliHost,
      configuredUrl: config.meiliUrl,
      usingLocalhostFallback: config.usingFallback,
      apiKeyConfigured: config.hasApiKey,
      envFileFound: config.envFileFound,
      connected: false,
      healthStatus: null,
      indexName: INDEX_NAME,
      documentCount: null,
      isIndexing: null,
      error: null,
    },
  };

  if (!config.hasApiKey) {
    report.meilisearch.error = 'MEILI_MASTER_KEY is not set in .env';
    return report;
  }

  if (config.usingFallback) {
    report.meilisearch.error =
      'MEILI_URL is missing — using localhost:7700 fallback (uploads will not reach Railway)';
  }

  try {
    const client = new MeiliSearch({
      host: config.meiliUrl,
      apiKey: process.env.MEILI_MASTER_KEY,
    });

    const health = await client.health();
    report.meilisearch.healthStatus = health.status;
    report.meilisearch.connected = health.status === 'available';

    const stats = await client.index(INDEX_NAME).getStats();
    report.meilisearch.documentCount = stats.numberOfDocuments;
    report.meilisearch.isIndexing = stats.isIndexing;

    if (report.meilisearch.connected && !config.usingFallback) {
      report.ok = true;
    }
  } catch (err) {
    report.meilisearch.connected = false;
    report.meilisearch.error = err.message;
  }

  if (state) {
    const publicState = state.getPublicState();
    const filesIndexed = publicState.stats?.filesIndexed ?? 0;
    const ledgerCount = publicState.uploadedFilesCount ?? 0;

    report.local = {
      filesIndexed,
      uploadedLedgerCount: ledgerCount,
      unparseableCount: publicState.unparseableFilesCount ?? 0,
      indexingStatus: publicState.indexingStatus,
    };

    if (report.meilisearch.documentCount != null) {
      report.local.documentCountMismatch =
        filesIndexed !== report.meilisearch.documentCount;
      report.local.mismatchDelta = filesIndexed - report.meilisearch.documentCount;
    }
  }

  return report;
}

function formatDiagnosticsForConsole(report) {
  const lines = [];
  const m = report.meilisearch;
  const status = (ok, label) => (ok ? `OK   ${label}` : `FAIL ${label}`);

  lines.push('========================================');
  lines.push(' Koyomail Indexer — Meilisearch Diagnostics');
  lines.push('========================================');
  if (report.version) lines.push(`App version:        ${report.version}`);
  lines.push(`Checked at:         ${report.timestamp}`);
  lines.push('');
  lines.push('Configuration:');
  lines.push(`  .env file found:  ${m.envFileFound ? 'yes' : 'NO'}`);
  lines.push(`  Meilisearch host: ${m.configuredHost}`);
  lines.push(`  API key set:      ${m.apiKeyConfigured ? 'yes' : 'NO'}`);
  lines.push(
    `  URL fallback:     ${m.usingLocalhostFallback ? 'YES (localhost — wrong for production!)' : 'no'}`
  );
  lines.push('');
  lines.push('Connection:');
  lines.push(`  ${status(m.connected, 'Connected to Meilisearch')}`);
  if (m.healthStatus) lines.push(`  Health status:    ${m.healthStatus}`);
  if (m.documentCount != null) {
    lines.push(`  Documents in DB:  ${m.documentCount}`);
  }
  if (m.isIndexing != null) {
    lines.push(`  Indexing active:  ${m.isIndexing ? 'yes' : 'no'}`);
  }
  if (m.error) lines.push(`  Error:            ${m.error}`);

  if (report.local) {
    lines.push('');
    lines.push('Local indexer state:');
    lines.push(`  Dashboard indexed: ${report.local.filesIndexed}`);
    lines.push(`  Ledger entries:    ${report.local.uploadedLedgerCount}`);
    lines.push(`  Unparseable files: ${report.local.unparseableCount}`);
    lines.push(`  Indexing status:   ${report.local.indexingStatus}`);
    if (report.local.documentCountMismatch) {
      lines.push(
        `  WARNING: Local count (${report.local.filesIndexed}) does not match ` +
        `Meilisearch (${m.documentCount}) — delta ${report.local.mismatchDelta}`
      );
    }
  }

  lines.push('');
  lines.push(report.ok ? 'RESULT: PASS — Meilisearch connection looks good.' : 'RESULT: FAIL — Fix configuration before indexing.');
  lines.push('========================================');
  return lines.join('\n');
}

module.exports = {
  runMeiliDiagnostics,
  formatDiagnosticsForConsole,
  getMeiliConfig,
};
