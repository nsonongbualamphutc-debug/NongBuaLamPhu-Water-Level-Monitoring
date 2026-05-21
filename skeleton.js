/**
 * skeleton.js
 * Skeleton loading สำหรับ index.html
 * ─────────────────────────────────────────────────────
 * วิธีใช้:
 *   1. อัปโหลดไฟล์นี้ขึ้น repo
 *   2. ใน index.html เพิ่ม <script src="skeleton.js"></script>
 *      ก่อน </body>
 *   3. เรียก Skeleton.show() ตอนเริ่ม fetch
 *      เรียก Skeleton.hide() ตอน data พร้อม
 * ─────────────────────────────────────────────────────
 */

(function(global) {
  'use strict';

  /* ── inject CSS ── */
  function injectStyles() {
    if (document.getElementById('nbp-sk-style')) return;
    const s = document.createElement('style');
    s.id = 'nbp-sk-style';
    s.textContent = `
/* Skeleton base */
.sk-pulse {
  background: linear-gradient(
    90deg,
    rgba(148,163,184,.12) 25%,
    rgba(148,163,184,.22) 50%,
    rgba(148,163,184,.12) 75%
  );
  background-size: 200% 100%;
  animation: skShimmer 1.5s infinite;
  border-radius: 8px;
}
[data-theme="dark"] .sk-pulse {
  background: linear-gradient(
    90deg,
    rgba(255,255,255,.05) 25%,
    rgba(255,255,255,.1)  50%,
    rgba(255,255,255,.05) 75%
  );
  background-size: 200% 100%;
  animation: skShimmer 1.5s infinite;
}
@keyframes skShimmer {
  0%   { background-position:  200% 0; }
  100% { background-position: -200% 0; }
}

/* Summary cards skeleton */
.sk-summary-row {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  gap: 14px;
  margin-bottom: 20px;
}
.sk-summary-card {
  height: 96px;
  border-radius: 14px;
}

/* Station card skeleton */
.sk-station-row {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 12px;
  margin-bottom: 20px;
}
.sk-station-card {
  height: 140px;
  border-radius: 14px;
}

/* Map skeleton */
.sk-map {
  width: 100%;
  height: 520px;
  border-radius: 14px;
  margin-bottom: 20px;
}

/* Chart skeleton */
.sk-chart {
  width: 100%;
  height: 220px;
  border-radius: 14px;
  margin-bottom: 14px;
}

/* Section title skeleton */
.sk-title {
  height: 22px;
  width: 220px;
  border-radius: 6px;
  margin-bottom: 14px;
}

/* fade in/out */
.sk-container {
  transition: opacity .3s ease;
}
.sk-container.sk-hidden {
  display: none;
}

/* hide actual content while loading */
.sk-loading {
  opacity: 0;
  pointer-events: none;
  transition: opacity .35s ease;
}
.sk-ready {
  opacity: 1 !important;
  pointer-events: auto;
}
`;
    document.head.appendChild(s);
  }

  /* ── สร้าง skeleton elements ── */
  function createSummarySkeletons(count) {
    const row = document.createElement('div');
    row.className = 'sk-summary-row sk-container';
    row.id = 'sk-summary';
    for (let i = 0; i < count; i++) {
      const d = document.createElement('div');
      d.className = 'sk-pulse sk-summary-card';
      row.appendChild(d);
    }
    return row;
  }

  function createStationSkeletons(count) {
    const row = document.createElement('div');
    row.className = 'sk-station-row sk-container';
    row.id = 'sk-stations';
    for (let i = 0; i < count; i++) {
      const d = document.createElement('div');
      d.className = 'sk-pulse sk-station-card';
      row.appendChild(d);
    }
    return row;
  }

  function createMapSkeleton() {
    const d = document.createElement('div');
    d.className = 'sk-pulse sk-map sk-container';
    d.id = 'sk-map';
    return d;
  }

  function createChartSkeleton(id) {
    const d = document.createElement('div');
    d.className = 'sk-pulse sk-chart sk-container';
    d.id = id || 'sk-chart';
    return d;
  }

  /* ── CONFIG: ระบุ target elements ใน index.html ──
     แก้ selector ให้ตรงกับ index.html จริง
  ──────────────────────────────────────────────── */
  const TARGETS = [
    /* { selector: 'element ที่จะซ่อน', skeletonFn: ฟังก์ชันสร้าง skeleton } */
    {
      selector:   '.summary-cards, .summary-row, #summaryCards, #kpiRow',
      skeletonFn: () => createSummarySkeletons(6),
      skId:       'sk-summary',
    },
    {
      selector:   '.station-cards, .stations-grid, #stationCards, #stationGrid',
      skeletonFn: () => createStationSkeletons(8),
      skId:       'sk-stations',
    },
    {
      selector:   '#mainMap, .map-panel, #resMap',
      skeletonFn: createMapSkeleton,
      skId:       'sk-map',
    },
  ];

  let _active = false;

  /* ════════════════════════════════════
   *  PUBLIC API
   * ════════════════════════════════════ */

  /**
   * Skeleton.show()
   * เรียกก่อนเริ่ม fetch API
   * @param {Object} opts
   *   opts.targets  {Array}   — selector เพิ่มเติม
   *   opts.map      {boolean} — แสดง map skeleton ด้วยไหม (default: false)
   */
  function show(opts) {
    if (_active) return;
    _active = true;
    opts = opts || {};
    injectStyles();

    TARGETS.forEach(function(t) {
      if (t.skId === 'sk-map' && !opts.map) return;

      const el = document.querySelector(t.selector);
      if (!el) return;

      /* ซ่อน content จริง */
      el.classList.add('sk-loading');

      /* ใส่ skeleton ก่อนหน้า element */
      const sk = t.skeletonFn();
      el.parentNode.insertBefore(sk, el);
    });

    /* ถ้า map container มีอยู่แล้ว ให้แสดง placeholder */
    if (opts.map) {
      const mapEl = document.querySelector('#mainMap, #resMap');
      if (mapEl && mapEl.offsetHeight < 10) {
        mapEl.style.minHeight = '520px';
      }
    }
  }

  /**
   * Skeleton.hide()
   * เรียกหลัง render data เสร็จ
   */
  function hide() {
    if (!_active) return;

    /* ลบ skeleton elements */
    document.querySelectorAll('.sk-container').forEach(function(el) {
      el.style.opacity = '0';
      el.style.transition = 'opacity .3s';
      setTimeout(function() { el.remove(); }, 300);
    });

    /* แสดง content จริง */
    document.querySelectorAll('.sk-loading').forEach(function(el) {
      el.classList.remove('sk-loading');
      el.classList.add('sk-ready');
    });

    _active = false;
  }

  /**
   * Skeleton.card(container, count, height)
   * สร้าง skeleton cards แบบ inline ใน container ที่ระบุ
   * ใช้เมื่อต้องการ skeleton เฉพาะ section
   *
   * @example
   * Skeleton.card('#stationSection', 6, 140);
   */
  function card(containerSelector, count, height) {
    injectStyles();
    const container = document.querySelector(containerSelector);
    if (!container) return;

    container.innerHTML = '';
    const grid = document.createElement('div');
    grid.style.cssText = 'display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px;';

    for (let i = 0; i < (count || 6); i++) {
      const d = document.createElement('div');
      d.className = 'sk-pulse';
      d.style.height = (height || 120) + 'px';
      d.style.borderRadius = '12px';
      grid.appendChild(d);
    }
    container.appendChild(grid);

    return {
      /** เรียก .replace(html) เพื่อแทนที่ skeleton ด้วย content จริง */
      replace: function(html) {
        grid.style.opacity = '0';
        grid.style.transition = 'opacity .25s';
        setTimeout(function() {
          container.innerHTML = html;
          container.style.opacity = '0';
          container.style.transition = 'opacity .3s';
          requestAnimationFrame(function() {
            container.style.opacity = '1';
          });
        }, 250);
      }
    };
  }

  /**
   * Skeleton.wrap(el)
   * wrap element เดี่ยวด้วย skeleton แบบ inline
   * คืน { done() } เพื่อ unwrap
   */
  function wrap(elOrSelector) {
    injectStyles();
    const el = typeof elOrSelector === 'string'
      ? document.querySelector(elOrSelector)
      : elOrSelector;
    if (!el) return { done: function(){} };

    const h = el.offsetHeight || 60;
    const w = el.offsetWidth  || 200;

    const sk = document.createElement('div');
    sk.className = 'sk-pulse';
    sk.style.cssText = `width:${w}px;height:${h}px;border-radius:8px;display:inline-block;`;

    el.style.display = 'none';
    el.parentNode.insertBefore(sk, el);

    return {
      done: function() {
        sk.remove();
        el.style.display = '';
        el.style.opacity = '0';
        el.style.transition = 'opacity .3s';
        requestAnimationFrame(function() { el.style.opacity = '1'; });
      }
    };
  }

  /* ── Export ── */
  global.Skeleton = { show, hide, card, wrap };

})(window);


/* ════════════════════════════════════════════════════════
 * วิธีใช้ใน index.html
 * ════════════════════════════════════════════════════════
 *
 * ก่อน fetch:
 *   Skeleton.show();
 *
 * หลัง render เสร็จ:
 *   Skeleton.hide();
 *
 * หรือแบบ per-section:
 *   const sk = Skeleton.card('#stationGrid', 8, 140);
 *   // ... fetch ...
 *   sk.replace('<div>...station cards html...</div>');
 *
 * ตัวอย่าง integration ใน loadData():
 * ─────────────────────────────────────
 * async function loadData() {
 *   Skeleton.show();
 *   try {
 *     const res = await fetch(API_URL + '?action=summary');
 *     const data = await res.json();
 *     renderSummaryCards(data);
 *     renderStationCards(data.stations);
 *     renderCharts(data);
 *   } catch(e) {
 *     console.error(e);
 *   } finally {
 *     Skeleton.hide();
 *   }
 * }
 * ═══════════════════════════════════════════════════════ */
