import { MongoClient } from "mongodb";


// ─── Connection ───────────────────────────────────────────────────────────────
// NOTE: env vars are read lazily (inside getCollection) to ensure dotenv has
// already been initialised by config/index.js before we read them.

let _client = null;
let _col    = null;

async function getCollection() {
  if (_col) return _col;

  // Read env vars lazily — dotenv is guaranteed to be loaded by now
  const MONGO_URI = process.env.MONGO_URI;
  const DB_NAME   = process.env.MONGO_DB_NAME  || 'koyomail_analytics';
  const COL_NAME  = process.env.MONGO_COL_NAME || 'search_events';

  if (!MONGO_URI) {
    throw new Error("MONGO_URI environment variable is not set");
  }

  console.log(`[analyticsStore] Connecting... URI starts with: ${MONGO_URI.substring(0, 30)}`);

  const client = new MongoClient(MONGO_URI, {
    serverSelectionTimeoutMS: 8000,
    connectTimeoutMS: 8000,
    socketTimeoutMS: 10000,
    tls: true,
  });

  await client.connect();
  // Quick ping to confirm the connection is alive
  await client.db("admin").command({ ping: 1 });
  console.log(`[analyticsStore] ✅ Ping OK. Using DB=${DB_NAME} COL=${COL_NAME}`);

  _client = client;
  _col = _client.db(DB_NAME).collection(COL_NAME);
  return _col;
}

/**
 * Pre-warm the MongoDB connection on server startup so the first search
 * does not pay the connection penalty. Called once from server.js.
 */
export async function warmupAnalyticsConnection() {
  try {
    await getCollection();
    console.log("[analyticsStore] ✅ Connection pre-warmed on startup.");
  } catch (err) {
    console.warn("[analyticsStore] ⚠️ Pre-warm failed (will retry on first search):", err.message);
  }
}

/**
 * Increment the search count for a given year + project.
 * Writes a single event document to MongoDB Atlas.
 * Retries up to 2 times if the write fails (handles transient cold-start
 * MongoDB connection issues that silently drop the first analytics event).
 */
export async function incrementSearchCount(year, project) {
  const MAX_RETRIES = 2;
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const col = await getCollection();
      await col.insertOne({
        ts:      Date.now(),
        year:    String(year),
        project: String(project),
      });
      console.log(`[analyticsStore] ✅ Recorded: year=${year}, project=${project}`);
      return; // Success — exit
    } catch (err) {
      // Reset connection so next attempt retries fresh
      _client = null;
      _col = null;
      if (attempt < MAX_RETRIES) {
        console.warn(`[analyticsStore] ⚠️ Attempt ${attempt} failed, retrying in 3s... (${err.message})`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      } else {
        console.error("[analyticsStore] ❌ All attempts failed:", err.message);
      }
    }
  }
}
