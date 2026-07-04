const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });
const { MeiliSearch } = require('meilisearch');

async function setupDatabase() {
  console.log('Connecting to Meilisearch at:', process.env.MEILI_URL);

  const client = new MeiliSearch({
    host: process.env.MEILI_URL,
    apiKey: process.env.MEILI_MASTER_KEY,
  });

  try {
    // Check health
    const health = await client.health();
    console.log('Meilisearch Health:', health);

    // Create the emails index
    console.log('Ensuring "emails" index exists...');
    const index = client.index('emails');

    // Update settings
    console.log('Configuring searchable, filterable, and sortable attributes...');
    await index.updateSettings({
      searchableAttributes: [
        'subject',
        'sender',
        'recipients',
        'cc',
        'bcc',
        'comment',
        'body',
        'filePath'
      ],
      filterableAttributes: [
        'hasAttachments',
        'sentAt',
        'filePath',
        'sender',
        'recipients',
        'cc',
        'bcc',
        'indexedRootPath',
        'indexedRootType',
        'collectionId',
        'isPublic',
        'allowedUsers'
      ],
      sortableAttributes: [
        'sentAt'
      ]
    });

    console.log('✅ Database setup successfully!');
    
    // Fetch and display current settings to confirm
    const settings = await index.getSettings();
    console.log('\nCurrent Searchable Attributes:', settings.searchableAttributes);
    console.log('Current Filterable Attributes:', settings.filterableAttributes);
    console.log('Current Sortable Attributes:', settings.sortableAttributes);

  } catch (error) {
    console.error('❌ Failed to setup database:', error.message);
    if (error.cause) console.error('Cause:', error.cause);
  }
}

setupDatabase();
