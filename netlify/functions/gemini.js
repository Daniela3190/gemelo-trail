const https = require('https');

const GEMINI_KEY = process.env.GEMINI_KEY;

const CORS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json',
};

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: CORS, body: '' };
  }

  return new Promise((resolve) => {
    let body;
    try {
      body = JSON.parse(event.body);
    } catch (e) {
      return resolve({ statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'invalid JSON' }) });
    }

    const options = {
      hostname: 'generativelanguage.googleapis.com',
      path:     `/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_KEY}`,
      method:   'POST',
      headers:  { 'Content-Type': 'application/json' },
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({
        statusCode: res.statusCode,
        headers: CORS,
        body: data,
      }));
    });

    req.on('error', (err) => resolve({
      statusCode: 502,
      headers: CORS,
      body: JSON.stringify({ error: err.message }),
    }));

    req.write(JSON.stringify(body));
    req.end();
  });
};
