// Service Worker for the Sheila Receipts PWA.
// Keeps the app shell available offline so the home-screen icon launches even
// with no network. API calls (/api/...) are network-only; the page falls back to
// the IndexedDB queue. Background Sync drains the queue when connectivity returns,
// POSTing each item to its mode's endpoint (/api/save-out or /api/save-in).

const CACHE_NAME = 'sheila-shell-v4';
const SHELL_URLS = ['/', '/index.html', '/treatments', '/treatments.html', '/manifest.webmanifest', '/treatments.webmanifest'];

const DB_NAME = 'sheila_queue_db';
const DB_VERSION = 1;
const STORE = 'pending';

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    // addAll fails the whole install if any URL 404s; add individually & tolerantly.
    await Promise.all(SHELL_URLS.map(u => cache.add(u).catch(() => {})));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter(n => n !== CACHE_NAME).map(n => caches.delete(n)));
    await self.clients.claim();
  })());
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.url) || '/';
  event.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of all) { if ('focus' in c) return c.focus(); }
    if (self.clients.openWindow) return self.clients.openWindow(url);
  })());
});

/* ---------- IndexedDB (mirror of index.html) ---------- */
function swOpenDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function swList() {
  const db = await swOpenDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction([STORE], 'readonly');
    const req = tx.objectStore(STORE).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}
async function swRemove(id) {
  const db = await swOpenDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction([STORE], 'readwrite');
    const req = tx.objectStore(STORE).delete(id);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

function endpointFor(mode) { return mode === 'in' ? '/api/save-in' : '/api/save-out'; }
function payloadFromItem(it) {
  const base = {
    date: it.date || '', totalCost: Number(it.totalCost) || 0, currency: it.currency || '',
    receiptCode: it.receiptCode || '', note: it.note || '', image: it.image,
    lat: it.lat, lng: it.lng, capturedAt: it.capturedAt || '',
  };
  if (it.mode === 'in') return { ...base, clientName: it.clientName || '', treatment: it.treatment || '' };
  return { ...base, merchant: it.merchant || '', article: it.article || 'General Business Expenses', businessPersonal: it.businessPersonal || 'Business' };
}

async function drainQueue() {
  const items = await swList();
  if (!items.length) return;
  let filed = 0, lastError = null;
  for (const item of items) {
    try {
      const resp = await fetch(endpointFor(item.mode), {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payloadFromItem(item)),
      });
      if (!resp.ok) throw new Error('save HTTP ' + resp.status);
      await swRemove(item.id);
      filed++;
    } catch (e) { lastError = e; break; }
  }
  if (filed > 0) {
    try {
      await self.registration.showNotification('Receipts filed ✓', {
        body: filed === 1 ? 'A queued receipt was filed automatically.' : filed + ' queued receipts were filed automatically.',
        tag: 'sheila-filed', data: { url: '/' },
      });
    } catch (_) {}
    const clients = await self.clients.matchAll({ includeUncontrolled: true });
    for (const c of clients) c.postMessage({ type: 'queue-drained', filed });
  }
  if (lastError) throw lastError;  // keep Background Sync registered to retry
}

self.addEventListener('sync', (event) => {
  if (event.tag === 'drain-queue') event.waitUntil(drainQueue());
});
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'drain-now') event.waitUntil(drainQueue().catch(() => {}));
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);
  if (url.pathname.startsWith('/api/')) return;  // never cache API calls
  if (req.method !== 'GET') return;
  event.respondWith((async () => {
    const cached = await caches.match(req, { ignoreSearch: true });
    if (cached) return cached;
    try {
      const fresh = await fetch(req);
      if (fresh.ok && url.origin === self.location.origin) {
        const cache = await caches.open(CACHE_NAME);
        cache.put(req, fresh.clone()).catch(() => {});
      }
      return fresh;
    } catch (e) {
      return new Response('<h1>Offline</h1><p>Reconnect once, then it works offline next time.</p>', { status: 503, headers: { 'Content-Type': 'text/html' } });
    }
  })());
});
