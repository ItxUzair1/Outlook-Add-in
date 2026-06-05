const https = require('https');
const url = 'https://surprising-light-production-6637.up.railway.app/assets/new_logo_transparent.png';

https.get(url, (res) => {
  console.log('Status:', res.statusCode);
  console.log('Content-Type:', res.headers['content-type']);
  console.log('Content-Length:', res.headers['content-length']);
  res.destroy();
}).on('error', (e) => {
  console.error('Error:', e.message);
});
