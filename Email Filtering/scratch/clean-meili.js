const { MeiliSearch } = require('meilisearch');
require('dotenv').config({ path: '../indexer/.env' });

async function cleanDB() {
  console.log('Connecting to Meilisearch at:', process.env.MEILI_URL);

  const client = new MeiliSearch({
    host: process.env.MEILI_URL,
    apiKey: process.env.MEILI_MASTER_KEY,
  });

  try {
    const index = client.index('emails');
    console.log('Deleting all documents from emails index...');
    const task = await index.deleteAllDocuments();
    console.log('Delete task queued:', task);
    
    // Wait for task
    const resolvedTask = await client.waitForTask(task.taskUid);
    console.log('Task resolved:', resolvedTask.status);

    console.log('Successfully cleaned Meilisearch database!');
  } catch (err) {
    console.error('Failed to clean database:', err.message);
  }
}

cleanDB();
