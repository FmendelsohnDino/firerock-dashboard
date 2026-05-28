const https = require('https');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  const { content } = JSON.parse(event.body);
  const TOKEN = process.env.GITHUB_TOKEN;
  const REPO = 'FmendelsohnDino/firerock-dashboard';
  const FILE = 'sales_report.xlsx';

  // Get current SHA
  const getSha = () => new Promise((resolve) => {
    const req = https.request({
      hostname: 'api.github.com',
      path: `/repos/${REPO}/contents/${FILE}`,
      method: 'GET',
      headers: { 'Authorization': `token ${TOKEN}`, 'User-Agent': 'FireRock', 'Accept': 'application/vnd.github.v3+json' }
    }, (res) => {
      let data = '';
      res.on('data', d => data += d);
      res.on('end', () => {
        try { resolve(JSON.parse(data).sha || null); } catch { resolve(null); }
      });
    });
    req.on('error', () => resolve(null));
    req.end();
  });

  const sha = await getSha();
  const payload = JSON.stringify({ message: 'Auto-update sales report', content, ...(sha ? { sha } : {}) });

  const result = await new Promise((resolve) => {
    const req = https.request({
      hostname: 'api.github.com',
      path: `/repos/${REPO}/contents/${FILE}`,
      method: 'PUT',
      headers: {
        'Authorization': `token ${TOKEN}`, 'User-Agent': 'FireRock',
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload)
      }
    }, (res) => {
      let data = '';
      res.on('data', d => data += d);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', (e) => resolve({ status: 500, body: e.message }));
    req.write(payload);
    req.end();
  });

  return {
    statusCode: result.status,
    headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
    body: result.body
  };
};
