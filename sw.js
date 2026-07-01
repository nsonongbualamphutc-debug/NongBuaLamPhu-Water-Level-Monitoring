/* ════════════════════════════════════════════════════════════
 *  Service Worker — สถานการณ์น้ำ จ.หนองบัวลำภู
 *  ทำให้ PWA เปิด offline ได้ + โหลดเร็วขึ้น
 *
 *  กลยุทธ์:
 *  1. หน้า HTML            → Network-first (เอาของใหม่ก่อน, offline ใช้ cache)
 *  2. ไฟล์ static ในเว็บ    → Stale-while-revalidate (โชว์ cache ทันที + อัปเดตเบื้องหลัง)
 *  3. API ภายนอก (GAS,     → Network-first + เก็บสำเนาล่าสุดไว้
 *     Open-Meteo, RainViewer)  offline จะเห็นข้อมูลครั้งล่าสุดแทนหน้าว่าง
 *
 *  ⚠️ เวลาแก้ไฟล์ HTML/CSS แล้ว deploy ใหม่ ให้ขยับเลข VERSION ทุกครั้ง
 *     เพื่อบังคับล้าง cache เก่าของผู้ใช้
 * ════════════════════════════════════════════════════════════ */
const VERSION       = "wnb-v5";
const STATIC_CACHE  = VERSION + "-static";
const RUNTIME_CACHE = VERSION + "-runtime";

/* ไฟล์หลักที่ precache ตอนติดตั้ง (พลาดบางไฟล์ได้ ไม่ทำให้ install ล้ม) */
const PRECACHE_URLS = [
  "./",
  "./index.html",
  "./daily_briefing.html",
  "./paneang.html",
  "./mong.html",
  "./mo.html",
  "./phuay.html",
  "./rainfall.html",
  "./reservoir.html",
  "./input.html",
  "./theme.css",
  "./config.js",
  "./manifest.json",
  "./verdict.css",
  "./verdict.js",
  "./logo-nbp.png",
  "./icon-192.png",
  "./icon-512.png",
  "./favicon.ico",
  "./favicon-32.png",
  "./nbp_water_points.geojson",
  "./nbp_water_lines.geojson"
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(STATIC_CACHE).then(function (cache) {
      /* add ทีละไฟล์ + กลืน error — ไฟล์ไหนไม่มีก็ข้าม ไม่ล้มทั้ง install */
      return Promise.all(
        PRECACHE_URLS.map(function (url) {
          return cache.add(url).catch(function () {});
        })
      );
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k.indexOf(VERSION) !== 0; })
            .map(function (k) { return caches.delete(k); })
      );
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function (event) {
  const req = event.request;
  if (req.method !== "GET") return; /* ไม่ยุ่งกับ POST (บันทึกข้อมูล) */

  const url = new URL(req.url);

  /* ── 1. การนำทางไปหน้า HTML → network-first ── */
  if (req.mode === "navigate") {
    event.respondWith(
      fetch(req)
        .then(function (res) {
          const copy = res.clone();
          caches.open(STATIC_CACHE).then(function (c) { c.put(req, copy); });
          return res;
        })
        .catch(function () {
          return caches.match(req).then(function (hit) {
            return hit || caches.match("./index.html");
          });
        })
    );
    return;
  }

  /* ── 2. ไฟล์ในโดเมนเดียวกัน (css/js/geojson) → stale-while-revalidate ── */
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then(function (hit) {
        const fresh = fetch(req).then(function (res) {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(STATIC_CACHE).then(function (c) { c.put(req, copy); });
          }
          return res;
        }).catch(function () { return hit; });
        return hit || fresh;
      })
    );
    return;
  }

  /* ── 3. API ภายนอก → network-first + เก็บสำเนาล่าสุด ──
   *  ครอบคลุม: script.google.com (GAS), open-meteo, rainviewer,
   *  raw.githubusercontent (GeoJSON เขตอำเภอ), tile servers ฯลฯ */
  event.respondWith(
    fetch(req)
      .then(function (res) {
        if (res && (res.ok || res.type === "opaque")) {
          const copy = res.clone();
          caches.open(RUNTIME_CACHE).then(function (c) { c.put(req, copy); });
        }
        return res;
      })
      .catch(function () {
        return caches.match(req);
      })
  );
});
