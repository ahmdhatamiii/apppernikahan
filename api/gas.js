export default async function handler(req, res) {
  const GAS_URL = process.env.GAS_EXEC_URL || "https://script.google.com/macros/s/AKfycbw3x7wWVwlaUf_UdPA2lw9dl7aB1yEuIzTf9hi4d1NFWYfaavcNhlu7roXCCLRtGHQp/exec";

  const action = req.query.action || '';
  const params = req.query.params || '[]';

  if (!action) {
    return res.status(400).json({ status: 'error', message: 'Parameter action wajib' });
  }

  const url = `${GAS_URL}?action=${encodeURIComponent(action)}&params=${encodeURIComponent(params)}`;

  try {
    let response = await fetch(url, { redirect: 'follow' });
    let text = await response.text();

    if (text.includes('<!DOCTYPE html') || text.includes('<!doctype html')) {
      const metaMatch = text.match(/content=["']0;\s*url=(.*?)["']/i);
      const jsMatch = text.match(/location\.(?:replace|href)\s*[\(=]\s*["'](https:\/\/[^"']+)["']/);
      const redirectUrl = metaMatch && metaMatch[1] ? metaMatch[1] : (jsMatch && jsMatch[1] ? jsMatch[1] : null);

      if (redirectUrl) {
        const redir = await fetch(redirectUrl, { redirect: 'follow' });
        text = await redir.text();
      }
    }

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Content-Type', 'application/json');

    if (req.method === 'OPTIONS') {
      return res.status(200).end();
    }

    try {
      JSON.parse(text);
      return res.status(200).send(text);
    } catch (e) {
      return res.status(502).json({
        status: 'error',
        message: 'GAS tidak mengembalikan JSON valid. Pastikan Web App ter-deploy dengan execute as Me & access Anyone.',
        preview: text.slice(0, 300)
      });
    }
  } catch (error) {
    return res.status(500).json({ status: 'error', message: error.message });
  }
}
