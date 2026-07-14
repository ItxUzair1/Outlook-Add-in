require('dotenv').config();
const express = require('express');
const fetch = require('node-fetch');

const app = express();
app.use(express.json());

// State to track if someone searched at night or on the weekend
let lastSearchTime = 0;

// Central Keep-Alive Loop for Meilisearch
setInterval(() => {
  const now = new Date();
  
  // Use Intl.DateTimeFormat to force exact UK time, handling Daylight Savings natively
  const formatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Europe/London',
    hour: 'numeric',
    weekday: 'short',
    hour12: false
  });
  const parts = formatter.formatToParts(now);
  const hourStr = parts.find(p => p.type === 'hour').value;
  const weekdayStr = parts.find(p => p.type === 'weekday').value;
  
  const hour = parseInt(hourStr, 10);
  const dayMap = { 'Sun': 0, 'Mon': 1, 'Tue': 2, 'Wed': 3, 'Thu': 4, 'Fri': 5, 'Sat': 6 };
  const day = dayMap[weekdayStr];

  let shouldPing = false;

  // Rule 1: Monday to Thursday: 7 AM to 7 PM (07:00 to 18:59)
  if (day >= 1 && day <= 4 && hour >= 7 && hour < 19) {
    shouldPing = true;
  } 
  // Rule 2: Friday: 7 AM to 9 PM (07:00 to 20:59)
  else if (day === 5 && hour >= 7 && hour < 21) {
    shouldPing = true;
  }

  // Rule 3: WEEKEND & OFF-HOURS (30-Minute Wake)
  // If a user performed an occasional search during off-hours, we keep the server awake for 30 minutes!
  const timeSinceLastSearch = Date.now() - lastSearchTime;
  if (!shouldPing && timeSinceLastSearch < 30 * 60 * 1000) {
    shouldPing = true;
  }

  if (shouldPing) {
    const meiliUrl = process.env.MEILI_URL || 'http://localhost:7700';
    const meiliKey = process.env.MEILI_MASTER_KEY || '';
    
    fetch(`${meiliUrl}/indexes/emails/search`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${meiliKey}`
      },
      body: JSON.stringify({ q: "paul", limit: 1000 })
    })
      .then(res => {
        if (res.ok) console.log(`[Keep-Alive] Pung Meilisearch at ${now.toISOString()} (UK Hour: ${hour})`);
        else console.warn(`[Keep-Alive] Meilisearch returned status: ${res.status}`);
      })
      .catch(err => console.error(`[Keep-Alive] Failed to ping Meilisearch:`, err.message));
  }
}, 4 * 60 * 1000); // 4 minutes

// API Endpoint: Local desktops can hit this endpoint when a search occurs during off-hours
// to tell the central server to keep Meilisearch awake for 30 minutes!
app.post('/api/search-event', (req, res) => {
  lastSearchTime = Date.now();
  console.log(`[Keep-Alive] Search event received! Server will stay awake for 30 minutes.`);
  res.json({ success: true, message: "Server will stay awake for 30 mins." });
});

// Basic health check for Railway
app.get('/', (req, res) => {
  res.status(200).send('OK');
});

// Railway requires apps to bind to a port
const port = process.env.PORT || 3000;
app.listen(port, '0.0.0.0', () => {
  console.log(`✓ Central Pinger running on port ${port}`);
  console.log(`✓ UK Timezone business hours strictly enforced.`);
});
