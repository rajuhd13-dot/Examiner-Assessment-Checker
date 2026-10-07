import express from 'express';
import { createServer as createViteServer } from 'vite';

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  app.use(express.json({ limit: '50mb' }));

  // API proxy endpoint to bypass CORS and handle Google Apps Script redirects
  app.get('/api/proxy', async (req, res) => {
    const targetUrl = req.query.url as string;
    if (!targetUrl) {
      return res.status(400).json({ success: false, error: 'Missing url parameter' });
    }

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 30000); // 30s timeout for large datasets (47k+ records)

      const response = await fetch(targetUrl, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'application/json, text/plain, */*'
        },
        signal: controller.signal,
        redirect: 'follow'
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        return res.status(response.status).json({
          success: false,
          error: `Upstream HTTP ${response.status}: ${response.statusText}`
        });
      }

      const contentType = response.headers.get('content-type') || '';
      if (contentType.includes('application/json')) {
        const json = await response.json();
        return res.json(json);
      } else {
        const text = await response.text();
        try {
          const json = JSON.parse(text);
          return res.json(json);
        } catch {
          return res.send(text);
        }
      }
    } catch (err: any) {
      return res.status(502).json({
        success: false,
        error: err.name === 'AbortError' ? 'Upstream request timed out' : (err.message || 'Failed to fetch upstream')
      });
    }
  });

  // Direct sync helper endpoint
  app.get('/api/sync', async (req, res) => {
    const targetUrl = (req.query.url as string) || 'https://script.google.com/macros/s/AKfycbyesBP9GFM2tcfDdr_eBUGhUA-lLxF-9jSUHN4xpngdKyLb2vaeRJ9SbmxgiY5Zg-0-jg/exec';
    const separator = targetUrl.includes('?') ? '&' : '?';

    // 1. Try action=sync
    const syncUrl = `${targetUrl}${separator}action=sync&_t=${Date.now()}`;
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000);

      const upstreamRes = await fetch(syncUrl, { redirect: 'follow', signal: controller.signal });
      clearTimeout(timeoutId);

      if (upstreamRes.ok) {
        const json = await upstreamRes.json();
        if (json && (json.success || json.ok) && (json.rows || json.data)) {
          return res.json(json);
        }
      }
    } catch (e) {
      // Ignore
    }

    // 2. Try action=filterOptions
    try {
      const fallbackUrl = `${targetUrl}${separator}action=filterOptions&_t=${Date.now()}`;
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000);

      const upstreamRes = await fetch(fallbackUrl, { redirect: 'follow', signal: controller.signal });
      clearTimeout(timeoutId);

      if (upstreamRes.ok) {
        const json = await upstreamRes.json();
        if (json && (json.success || json.ok)) {
          return res.json(json);
        }
      }
    } catch (e) {
      // Ignore
    }

    return res.json({ success: false, error: 'Google Sheet Sync is currently offline or unreachable.' });
  });

  // Mount Vite middleware in development
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa'
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static('dist'));
    app.get('*', (_req, res) => {
      res.sendFile('dist/index.html', { root: '.' });
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server listening on http://0.0.0.0:${PORT}`);
  });
}

startServer();
