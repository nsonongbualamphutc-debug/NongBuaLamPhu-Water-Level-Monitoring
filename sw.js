/**
 * sw.js — Service Worker
 * ระบบสถานการณ์น้ำ จ.หนองบัวลำภู
 * ─────────────────────────────────────────────
 * กลยุทธ์:
 *   - Static assets  → Cache First (โหลดเร็ว offline ได้)
 *   - API / GeoJSON  → Network First, fallback cache
 *   - ภาพ/font       → Stale While Revalidate
 */

const CACHE_NAME    = 'nbp-water-v1.4';
const CACHE_STATIC  = 'nbp-static-v1.4';
const CACHE_API     = 'nbp-api-v1.4';

/* ── ไฟล์ที่ cache ตั้งแต่ install ── */
const PRECACHE_URLS = [
  './',
  './index.html',
  './daily_briefing.html',
  './paneang.html',
  './mong.html',
  './mo.html',
  './phuay.html',
  './rainfall.html',
  './reservoir.html',
  './input.html',
  './404.html',
  './config.js',
  './waterways-loader.js',
  './map-layers.js',
  './smart_summary.js',
  './manifest.json',
  './nbp_water_lines.geojson',
  './nbp_water_points.geojson',
];

/* ── domains ที่ใช้ Network First ── */
const API_ORIGINS = [
  'script.google.com',
  'open-meteo.com',
  'air4thai.pcd.go.th',
  'api.rainviewer.com',
];

/* ════════════════════════════════════════
 *  INSTALL — precache static assets
 * ════════════════════════════════════════ */
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_STATIC).then(cache => {
      return Promise.allSettled(
        PRECACHE_URLS.map(url =>
          cache.add(url).catch(err =>
            console.warn('[SW] precache miss:', url, err.message)
          )
        )
      );
    }).then(() => self.skipWaiting())
  );
});

/* ════════════════════════════════════════
 *  ACTIVATE — ลบ cache เวอร์ชันเก่า
 * ════════════════════════════════════════ */
self.addEventListener('activate', event => {
  const CURRENT = [CACHE_NAME, CACHE_STATIC, CACHE_API];
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => !CURRENT.includes(k))
          .map(k => {
            console.log('[SW] deleting old cache:', k);
            return caches.delete(k);
          })
      )
    ).then(() => self.clients.claim())
  );
});

/* ════════════════════════════════════════
 *  FETCH — routing strategy
 * ════════════════════════════════════════ */
self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);

  /* ① ข้าม non-GET, chrome-extension, ws:// */
  if (request.method !== 'GET') return;
  if (!['http:', 'https:'].includes(url.protocol)) return;

  /* ② API / GeoJSON จาก script.google.com, open-meteo, air4thai → Network First */
  if (API_ORIGINS.some(o => url.hostname.includes(o))) {
    event.respondWith(networkFirst(request, CACHE_API, 60));
    return;
  }

  /* ③ GeoJSON ใน repo → Network First (ข้อมูลอาจเปลี่ยน) */
  if (url.pathname.endsWith('.geojson')) {
    event.respondWith(networkFirst(request, CACHE_STATIC, 300));
    return;
  }

  /* ④ Fonts / CDN (leaflet, chart.js, fonts.google) → Stale While Revalidate */
  if (
    url.hostname.includes('fonts.') ||
    url.hostname.includes('unpkg.com') ||
    url.hostname.includes('cdn.jsdelivr.net') ||
    url.hostname.includes('cdnjs.cloudflare.com')
  ) {
    event.respondWith(staleWhileRevalidate(request, CACHE_STATIC));
    return;
  }

  /* ⑤ HTML pages → Network First, fallback cache ─ ถ้า offline ใช้ cached version */
  if (request.headers.get('accept')?.includes('text/html')) {
    event.respondWith(networkFirstHTML(request));
    return;
  }

  /* ⑥ อื่นๆ (JS, CSS, images) → Cache First */
  event.respondWith(cacheFirst(request, CACHE_STATIC));
});

/* ════════════════════════════════════════
 *  STRATEGIES
 * ════════════════════════════════════════ */

/** Cache First — ใช้ cache ก่อน, ถ้าไม่มีถึงไป network */
async function cacheFirst(request, cacheName) {
  const cached = await caches.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(cacheName);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    return new Response('Offline', { status: 503 });
  }
}

/** Network First — ไป network ก่อน, ถ้า fail ใช้ cache */
async function networkFirst(request, cacheName, maxAgeSeconds = 60) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(cacheName);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request);
    if (cached) return cached;
    return new Response(
      JSON.stringify({ ok: false, error: 'offline', _sw: true }),
      { headers: { 'Content-Type': 'application/json' } }
    );
  }
}

/** Network First สำหรับ HTML — fallback ไป index.html ถ้า 404 offline */
async function networkFirstHTML(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_STATIC);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request);
    if (cached) return cached;
    /* fallback ไปหน้า offline */
    const offlinePage = await caches.match('./404.html') ||
                        await caches.match('./index.html');
    return offlinePage || new Response('<h1>Offline</h1>', {
      headers: { 'Content-Type': 'text/html' }
    });
  }
}

/** Stale While Revalidate — คืน cache ทันที แล้ว update ใน background */
async function staleWhileRevalidate(request, cacheName) {
  const cache  = await caches.open(cacheName);
  const cached = await cache.match(request);

  const fetchPromise = fetch(request).then(response => {
    if (response.ok) cache.put(request, response.clone());
    return response;
  }).catch(() => null);

  return cached || fetchPromise;
}

/* ════════════════════════════════════════
 *  PUSH NOTIFICATIONS (Line Notify fallback)
 * ════════════════════════════════════════ */
self.addEventListener('push', event => {
  if (!event.data) return;
  let data = {};
  try { data = event.data.json(); } catch { data = { title: event.data.text() }; }

  const title   = data.title || '⚠️ แจ้งเตือนสถานการณ์น้ำ';
  const options = {
    body:    data.body    || 'มีการเปลี่ยนแปลงสถานการณ์น้ำ',
    icon:    data.icon    || './icon-192.png',
    badge:   data.badge   || './icon-192.png',
    tag:     data.tag     || 'nbp-water-alert',
    vibrate: [200, 100, 200],
    data:    { url: data.url || './' },
    actions: [
      { action: 'view',    title: '📊 ดูแดชบอร์ด' },
      { action: 'dismiss', title: '✕ ปิด' },
    ],
  };

  event.waitUntil(
    self.registration.showNotification(title, options)
  );
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  if (event.action === 'dismiss') return;

  const url = event.notification.data?.url || './';
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true })
      .then(clientList => {
        for (const client of clientList) {
          if (client.url.includes(url) && 'focus' in client) {
            return client.focus();
          }
        }
        return clients.openWindow(url);
      })
  );
});

/* ════════════════════════════════════════
 *  BACKGROUND SYNC (บันทึก offline แล้วส่งทีหลัง)
 * ════════════════════════════════════════ */
self.addEventListener('sync', event => {
  if (event.tag === 'nbp-sync-water') {
    event.waitUntil(syncPendingData());
  }
});

async function syncPendingData() {
  try {
    const cache = await caches.open('nbp-pending');
    const keys  = await cache.keys();
    for (const req of keys) {
      try {
        await fetch(req);
        await cache.delete(req);
        console.log('[SW] synced pending:', req.url);
      } catch (e) {
        console.warn('[SW] sync failed, will retry:', e.message);
      }
    }
  } catch (e) {
    console.error('[SW] syncPendingData error:', e);
  }
}

console.log('[SW] Service Worker loaded —', CACHE_NAME);
