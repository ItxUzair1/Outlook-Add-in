import { MongoClient } from "mongodb";

let _client = null;
let _col    = null;

async function getCollection() {
  if (_col) return _col;

  const MONGO_URI = process.env.MONGO_URI;
  const DB_NAME   = process.env.MONGO_DB_NAME  || 'koyomail_analytics';
  const COL_NAME  = 'indexing_requests';

  if (!MONGO_URI) {
    throw new Error("MONGO_URI environment variable is not set");
  }

  const client = new MongoClient(MONGO_URI, {
    serverSelectionTimeoutMS: 8000,
    connectTimeoutMS: 8000,
    socketTimeoutMS: 10000,
    tls: true,
  });

  await client.connect();
  _client = client;
  _col = _client.db(DB_NAME).collection(COL_NAME);
  return _col;
}

export async function createIndexingRequest(projectName, userEmail) {
  const MAX_RETRIES = 2;
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const col = await getCollection();
      const result = await col.insertOne({
        projectName: String(projectName),
        userEmail: String(userEmail),
        status: 'pending',
        createdAt: new Date(),
      });
      return result.insertedId;
    } catch (err) {
      _client = null;
      _col = null;
      if (attempt < MAX_RETRIES) {
        await new Promise(resolve => setTimeout(resolve, 3000));
      } else {
        throw new Error("Failed to create indexing request: " + err.message);
      }
    }
  }
}
