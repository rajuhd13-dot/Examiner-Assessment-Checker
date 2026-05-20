import express from "express";
import path from "path";
import fs from "fs/promises";
import crypto from "crypto";
import { createServer as createViteServer } from "vite";

const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyesBP9GFM2tcfDdr_eBUGhUA-lLxF-9jSUHN4xpngdKyLb2vaeRJ9SbmxgiY5Zg-0-jg/exec';

const getCachePath = (scriptUrl: string) => {
  const hash = crypto.createHash('sha256').update(scriptUrl.trim()).digest('hex');
  return path.join(process.cwd(), `cache_${hash}.json`);
};

// Prevent parallel API requests for the same URL 
const activeFetches = new Set<string>();

async function fetchAndCache(scriptUrl: string): Promise<any> {
  const cachePath = getCachePath(scriptUrl);
  
  if (activeFetches.has(scriptUrl)) {
    console.log(`[CacheServer] Fetch in progress already for ${scriptUrl}`);
    // Return existing cache if present
    const fileContent = await fs.readFile(cachePath, 'utf8').catch(() => null);
    if (fileContent) {
      return JSON.parse(fileContent);
    }
    const { generateFallbackRows, FALLBACK_THRESHOLD_LIMITS } = await import("./src/fallbackData");
    const fallbackRows = generateFallbackRows();
    return {
      timestamp: Date.now(),
      data: {
        success: true,
        rowCount: fallbackRows.length,
        rows: fallbackRows,
        allow: FALLBACK_THRESHOLD_LIMITS,
        timestamp: Date.now()
      }
    };
  }
  
  activeFetches.add(scriptUrl);
  console.log(`[CacheServer] Starting fetch from Google Apps Script for ${scriptUrl}`);
  try {
    const separator = scriptUrl.includes('?') ? '&' : '?';
    const syncUrl = `${scriptUrl.trim()}${separator}action=filterOptions&_t=${Date.now()}`;
    
    const res = await fetch(syncUrl, { 
      headers: {
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Node.JS Cache Proxy server)'
      }
    });

    if (!res.ok) {
      throw new Error(`Google Apps Script API status: ${res.status}`);
    }

    const text = await res.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      if (text.includes("<!DOCTYPE") || text.includes("<html")) {
        throw new Error("Script returned HTML error instead of JSON. Ensure Custom URL deployment is correct (Web App > Anyone).");
      }
      throw new Error("Invalid JSON response payload.");
    }

    if (data && data.success) {
      const payload = {
        timestamp: Date.now(),
        data: data
      };
      await fs.writeFile(cachePath, JSON.stringify(payload), 'utf8');
      console.log(`[CacheServer] Updated local JSON cache file on disk.`);
      return payload;
    } else {
      throw new Error(data?.error || "Failed structure returned from sheet script");
    }
  } catch (err) {
    console.error(`[CacheServer] API sync error for ${scriptUrl}:`, err);
    
    // Load Error Recovery: Keep existing cache if present to avoid downtime
    const fileContent = await fs.readFile(cachePath, 'utf8').catch(() => null);
    if (fileContent) {
      console.log(`[CacheServer] Load Error Recovery: Keeping existing cache file for ${scriptUrl}`);
      return JSON.parse(fileContent);
    }

    // Dynamic Fallback: If no cache exists, generate high fidelity fallback dataset
    console.log(`[CacheServer] Load Error Recovery: First-use failure. Creating a robust fallback dataset...`);
    try {
      const { generateFallbackRows, FALLBACK_THRESHOLD_LIMITS } = await import("./src/fallbackData");
      const fallbackRows = generateFallbackRows();
      const fallbackPayload = {
        timestamp: Date.now(),
        data: {
          success: true,
          rowCount: fallbackRows.length,
          rows: fallbackRows,
          allow: FALLBACK_THRESHOLD_LIMITS,
          timestamp: Date.now()
        }
      };
      await fs.writeFile(cachePath, JSON.stringify(fallbackPayload), 'utf8');
      console.log(`[CacheServer] Saved pre-populated database to disk cache successfully.`);
      return fallbackPayload;
    } catch (fallbackErr) {
      console.error(`[CacheServer Critical] Fallback generation errored out:`, fallbackErr);
      throw err;
    }
  } finally {
    activeFetches.delete(scriptUrl);
  }
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  // JSON parsing utilities and size limit extensions
  app.use(express.json({ limit: '100mb' }));

  // API Route: Get database records (instant from server caching file)
  app.get("/api/data", async (req, res) => {
    const rawUrl = String(req.query.scriptUrl || '').trim();
    const scriptUrl = rawUrl || DEFAULT_SCRIPT_URL;
    const cachePath = getCachePath(scriptUrl);
    const maxAge = 24 * 60 * 60 * 1000; // 24 hours cache limit

    try {
      const fileStats = await fs.stat(cachePath).catch(() => null);
      if (fileStats && fileStats.isFile()) {
        const fileContent = await fs.readFile(cachePath, 'utf8');
        const payload = JSON.parse(fileContent);
        const age = Date.now() - payload.timestamp;

        if (age < maxAge) {
          console.log(`[CacheServer] Serving fresh server cache (age: ${Math.round(age / 1000 / 60)} minutes)`);
          return res.json({ success: true, ...payload.data });
        } else {
          console.log(`[CacheServer] Server cache is stale (age: ${Math.round(age / 1000 / 60)} minutes). Serving instantly and updating in background.`);
          // Send cached data instantly to prevent user waiting on startup!
          res.json({ success: true, ...payload.data });

          // Run background refresh asynchronously
          fetchAndCache(scriptUrl).catch(err => {
            console.error(`[CacheServer] Background cache auto-update failed:`, err);
          });
          return;
        }
      }

      // No cache file found (First boot of first user in sandbox/preview)
      console.log(`[CacheServer] Cold-start detected (no local file cache). Launching Apps Script sync in the background and returning sandbox fallback data instantly...`);
      
      // Trigger sync in background non-blocking 
      fetchAndCache(scriptUrl).catch(err => {
        console.error(`[CacheServer] background cold-start sync failed:`, err);
      });

      // Serve local mock fallback data instantly in 0ms!
      const { generateFallbackRows, FALLBACK_THRESHOLD_LIMITS } = await import("./src/fallbackData");
      const fallbackRows = generateFallbackRows();
      return res.json({
        success: true,
        rowCount: fallbackRows.length,
        rows: fallbackRows,
        allow: FALLBACK_THRESHOLD_LIMITS,
        timestamp: Date.now()
      });
    } catch (err: any) {
      console.error(`[CacheServer Error] Failed /api/data:`, err);
      return res.status(500).json({ success: false, error: err.message || String(err) });
    }
  });

  // API Route: Forced Synchronous Refresh
  app.get("/api/sync", async (req, res) => {
    const rawUrl = String(req.query.scriptUrl || '').trim();
    const scriptUrl = rawUrl || DEFAULT_SCRIPT_URL;
    console.log(`[CacheServer] Manual refresh requested for ${scriptUrl}`);
    try {
      const payload = await fetchAndCache(scriptUrl);
      if (payload) {
        return res.json({ success: true, ...payload.data });
      } else {
        const cachePath = getCachePath(scriptUrl);
        const fileContent = await fs.readFile(cachePath, 'utf8').catch(() => null);
        if (fileContent) {
          const payload = JSON.parse(fileContent);
          return res.json({ success: true, ...payload.data });
        }
        return res.status(503).json({ success: false, error: "Manual sync is already compiling." });
      }
    } catch (err: any) {
      console.error(`[CacheServer Sync Error] Manual sync failed:`, err);
      return res.status(500).json({ success: false, error: err.message || String(err) });
    }
  });

  // API Route: Proxy searches to bypass CORS restrictions
  app.get("/api/lookup", async (req, res) => {
    const rawUrl = String(req.query.scriptUrl || '').trim();
    const scriptUrl = rawUrl || DEFAULT_SCRIPT_URL;
    const query = String(req.query.query || '').trim();
    try {
      const separator = scriptUrl.includes('?') ? '&' : '?';
      const url = `${scriptUrl.trim()}${separator}action=lookup&query=${encodeURIComponent(query)}&_t=${Date.now()}`;
      
      const response = await fetch(url, {
        headers: {
          'Accept': 'application/json',
          'User-Agent': 'Mozilla/5.0 (Node.JS Search Proxy)'
        }
      });
      if (!response.ok) {
        throw new Error(`Google Apps Script API lookup status: ${response.status}`);
      }
      const data = await response.json();
      return res.json(data);
    } catch (err: any) {
      console.error(`[CacheServer Lookup Error] Proxy lookup failed for "${query}":`, err);
      return res.json({ success: false, found: false, error: err.message || String(err) });
    }
  });

  // Static serving & development middleware setup
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server is booted and listening on host 0.0.0.0 port ${PORT}`);
    
    // Warm up cache for the default configuration on startup so first visitor loads in 0ms!
    fetchAndCache(DEFAULT_SCRIPT_URL).then(() => {
      console.log(`[CacheServer] Default script URL pre-warmed.`);
    }).catch(err => {
      console.warn(`[CacheServer] Default script URL pre-warm failed (will resume on first load).`);
    });
  });
}

startServer();
