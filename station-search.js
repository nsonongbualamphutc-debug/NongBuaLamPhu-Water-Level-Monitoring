/**
 * station-search.js
 * เครื่องมือค้นหา + กรองสถานี — ใช้ได้ทุกหน้าที่มีรายการสถานี
 * ระบบสถานการณ์น้ำ จ.หนองบัวลำภู
 * ─────────────────────────────────────────────────────────
 * ฟีเจอร์:
 *   - ช่องค้นหาชื่อสถานี (real-time, ภาษาไทย)
 *   - ปุ่มกรองตามสถานะ (ทั้งหมด / ปกติ / เฝ้าระวัง / วิกฤติ)
 *   - แสดงจำนวนผลลัพธ์
 *   - dark theme aware
 *   - คืน callback เมื่อ filter เปลี่ยน
 *
 * วิธีใช้:
 *   1. อัปโหลด station-search.js ขึ้น repo
 *   2. เพิ่ม <script src="station-search.js"></script> ใน <head>
 *   3. เรียก StationSearch.init({...})
 * ─────────────────────────────────────────────────────────
 */

(function(global) {
  'use strict';

  /* ── styles ── */
  function injectStyles() {
    if (document.getElementById('nbp-ss-style')) return;
    const s = document.createElement('style');
    s.id = 'nbp-ss-style';
    s.textContent = `
.ss-wrap {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  padding: 12px 16px;
  background: var(--card, #fff);
  border-bottom: 1px solid var(--line, #e2e8f0);
  font-family: 'Sarabun', sans-serif;
}

/* search input */
.ss-search {
  position: relative;
  flex: 1;
  min-width: 200px;
}
.ss-search input {
  width: 100%;
  padding: 9px 14px 9px 36px;
  border: 1px solid var(--line, #e2e8f0);
  border-radius: 10px;
  font-family: 'Sarabun', sans-serif;
  font-size: 13px;
  color: var(--ink, #0f172a);
  background: var(--bg, #f4f7fb);
  outline: none;
  transition: border-color .15s, box-shadow .15s;
}
.ss-search input:focus {
  border-color: #0ea5e9;
  box-shadow: 0 0 0 3px rgba(14,165,233,.12);
}
.ss-search-icon {
  position: absolute;
  left: 12px;
  top: 50%;
  transform: translateY(-50%);
  font-size: 14px;
  opacity: .5;
  pointer-events: none;
}
.ss-clear {
  position: absolute;
  right: 8px;
  top: 50%;
  transform: translateY(-50%);
  width: 20px; height: 20px;
  border: none;
  background: rgba(100,116,139,.15);
  border-radius: 50%;
  color: var(--muted, #64748b);
  font-size: 13px;
  cursor: pointer;
  display: none;
  align-items: center;
  justify-content: center;
  line-height: 1;
}
.ss-clear.show { display: flex; }
.ss-clear:hover { background: rgba(100,116,139,.3); }
[data-theme="dark"] .ss-search input {
  background: rgba(255,255,255,.07);
  border-color: rgba(255,255,255,.12);
  color: #e2e8f0;
}

/* status filter buttons */
.ss-filters {
  display: flex;
  gap: 5px;
  flex-wrap: wrap;
}
.ss-fbtn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 7px 13px;
  border-radius: 9px;
  border: 1px solid var(--line, #e2e8f0);
  background: transparent;
  color: var(--muted, #64748b);
  font-family: 'Sarabun', sans-serif;
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  transition: all .15s;
  white-space: nowrap;
}
.ss-fbtn:hover {
  border-color: #94a3b8;
  color: var(--ink, #0f172a);
}
.ss-fbtn .ss-dot {
  width: 8px; height: 8px;
  border-radius: 50%;
  flex-shrink: 0;
}
.ss-fbtn.active { color: #fff; border-color: transparent; }
.ss-fbtn.active.all    { background: linear-gradient(135deg,#475569,#64748b); }
.ss-fbtn.active.normal { background: linear-gradient(135deg,#15803d,#16a34a); }
.ss-fbtn.active.warn   { background: linear-gradient(135deg,#b45309,#f59e0b); }
.ss-fbtn.active.crit   { background: linear-gradient(135deg,#b91c1c,#dc2626); }
.ss-fbtn .ss-count {
  background: rgba(0,0,0,.12);
  border-radius: 6px;
  padding: 1px 6px;
  font-size: 10px;
  font-weight: 700;
}
.ss-fbtn.active .ss-count { background: rgba(255,255,255,.25); }
[data-theme="dark"] .ss-fbtn {
  border-color: rgba(255,255,255,.12);
  color: rgba(226,232,240,.7);
}

/* result count */
.ss-result {
  margin-left: auto;
  font-size: 12px;
  color: var(--muted, #64748b);
  font-weight: 600;
  white-space: nowrap;
}
.ss-result b { color: #0ea5e9; }

/* no-result message */
.ss-empty {
  display: none;
  padding: 40px 20px;
  text-align: center;
  color: var(--muted, #64748b);
  font-size: 14px;
}
.ss-empty.show { display: block; }
.ss-empty-icon { font-size: 36px; opacity: .4; margin-bottom: 8px; }

@media (max-width: 600px) {
  .ss-wrap { gap: 8px; }
  .ss-search { min-width: 100%; }
  .ss-result { width: 100%; text-align: center; margin: 0; }
}
`;
    document.head.appendChild(s);
  }

  /* ── status meta ── */
  const STATUS = [
    { id: 'all',    label: 'ทั้งหมด',   color: '#64748b' },
    { id: 'normal', label: 'ปกติ',      color: '#16a34a' },
    { id: 'warn',   label: 'เฝ้าระวัง', color: '#f59e0b' },
    { id: 'crit',   label: 'วิกฤติ',    color: '#dc2626' },
  ];

  /* map สถานะภาษาไทย → id */
  function statusToId(thaiStatus) {
    const s = String(thaiStatus || '').trim();
    if (s.indexOf('วิกฤต') !== -1) return 'crit';
    if (s.indexOf('เฝ้าระวัง') !== -1 || s.indexOf('ระวัง') !== -1) return 'warn';
    if (s.indexOf('ปกติ') !== -1) return 'normal';
    return 'normal';
  }

  /* ── state ── */
  let _config        = null;
  let _searchText    = '';
  let _activeStatus  = 'all';
  let _container     = null;

  /* ── build UI ── */
  function buildUI(host) {
    host.innerHTML = `
<div class="ss-wrap">
  <div class="ss-search">
    <span class="ss-search-icon">🔍</span>
    <input type="text" id="ssInput" placeholder="ค้นหาชื่อสถานี..."
           autocomplete="off">
    <button class="ss-clear" id="ssClear" title="ล้าง">×</button>
  </div>
  <div class="ss-filters" id="ssFilters">
    ${STATUS.map(st =>
      `<button class="ss-fbtn${st.id === 'all' ? ' active all' : ''}"
               data-status="${st.id}"
               onclick="StationSearch._setStatus('${st.id}')">
        ${st.id !== 'all'
          ? `<span class="ss-dot" style="background:${st.color}"></span>`
          : ''}
        ${st.label}
        <span class="ss-count" id="ssCount-${st.id}">0</span>
      </button>`
    ).join('')}
  </div>
  <div class="ss-result" id="ssResult">แสดง <b>0</b> สถานี</div>
</div>`;

    /* search input handler */
    const input = document.getElementById('ssInput');
    const clear = document.getElementById('ssClear');

    input.addEventListener('input', function() {
      _searchText = this.value.trim().toLowerCase();
      clear.classList.toggle('show', _searchText.length > 0);
      _applyFilter();
    });

    clear.addEventListener('click', function() {
      input.value = '';
      _searchText = '';
      clear.classList.remove('show');
      _applyFilter();
      input.focus();
    });
  }

  /* ── set status filter ── */
  function _setStatus(statusId) {
    _activeStatus = statusId;
    document.querySelectorAll('.ss-fbtn').forEach(function(b) {
      const isActive = b.dataset.status === statusId;
      b.classList.toggle('active', isActive);
      /* reset color class */
      b.classList.remove('all','normal','warn','crit');
      if (isActive) b.classList.add(statusId);
    });
    _applyFilter();
  }

  /* ── apply filter — ซ่อน/แสดง element ── */
  function _applyFilter() {
    if (!_config) return;

    const items = document.querySelectorAll(_config.itemSelector);
    let visible = 0;
    const counts = { all: 0, normal: 0, warn: 0, crit: 0 };

    items.forEach(function(el) {
      /* อ่านชื่อ + สถานะจาก element */
      const name   = getName(el).toLowerCase();
      const status = statusToId(getStatus(el));

      counts.all++;
      counts[status] = (counts[status] || 0) + 1;

      /* ตรวจเงื่อนไข */
      const matchSearch = !_searchText || name.indexOf(_searchText) !== -1;
      const matchStatus = _activeStatus === 'all' || status === _activeStatus;
      const show = matchSearch && matchStatus;

      el.style.display = show ? '' : 'none';
      if (show) visible++;
    });

    /* update counts on buttons */
    Object.keys(counts).forEach(function(k) {
      const el = document.getElementById('ssCount-' + k);
      if (el) el.textContent = counts[k];
    });

    /* update result text */
    const resultEl = document.getElementById('ssResult');
    if (resultEl) {
      resultEl.innerHTML = 'แสดง <b>' + visible + '</b> / ' +
                            counts.all + ' สถานี';
    }

    /* empty message */
    toggleEmpty(visible === 0);

    /* callback */
    if (typeof _config.onFilter === 'function') {
      _config.onFilter({
        search:  _searchText,
        status:  _activeStatus,
        visible: visible,
        total:   counts.all,
      });
    }
  }

  /* ── helper: อ่านชื่อจาก element ── */
  function getName(el) {
    if (_config.nameAttr && el.hasAttribute(_config.nameAttr)) {
      return el.getAttribute(_config.nameAttr);
    }
    if (_config.nameSelector) {
      const n = el.querySelector(_config.nameSelector);
      if (n) return n.textContent || '';
    }
    return el.textContent || '';
  }

  /* ── helper: อ่านสถานะจาก element ── */
  function getStatus(el) {
    if (_config.statusAttr && el.hasAttribute(_config.statusAttr)) {
      return el.getAttribute(_config.statusAttr);
    }
    if (_config.statusSelector) {
      const s = el.querySelector(_config.statusSelector);
      if (s) return s.textContent || '';
    }
    /* fallback — เดาจาก class */
    if (el.classList.contains('crit'))   return 'วิกฤติ';
    if (el.classList.contains('warn'))   return 'เฝ้าระวัง';
    if (el.classList.contains('normal')) return 'ปกติ';
    return '';
  }

  /* ── empty message ── */
  function toggleEmpty(show) {
    let el = document.getElementById('ssEmpty');
    if (!el && show) {
      el = document.createElement('div');
      el.id = 'ssEmpty';
      el.className = 'ss-empty';
      el.innerHTML =
        '<div class="ss-empty-icon">🔍</div>' +
        '<div>ไม่พบสถานีที่ตรงกับการค้นหา</div>';
      /* แทรกหลัง container ของ items */
      const firstItem = document.querySelector(_config.itemSelector);
      if (firstItem && firstItem.parentNode) {
        firstItem.parentNode.appendChild(el);
      }
    }
    if (el) el.classList.toggle('show', show);
  }

  /* ════════════════════════════════════
   *  PUBLIC API
   * ════════════════════════════════════ */

  /**
   * StationSearch.init(options)
   *
   * @param {Object} options
   *   options.container    {string} — selector ที่จะวาง search bar ไว้ด้านบน
   *   options.itemSelector {string} — selector ของ station card/row แต่ละอัน
   *   options.nameSelector {string} — selector ของชื่อสถานีภายใน item (optional)
   *   options.nameAttr     {string} — หรือใช้ attribute เช่น 'data-name' (optional)
   *   options.statusSelector {string} — selector ของ badge สถานะ (optional)
   *   options.statusAttr   {string} — หรือ attribute เช่น 'data-status' (optional)
   *   options.onFilter     {Function} — callback({ search, status, visible, total })
   *
   * ตัวอย่าง:
   * ──────────
   * StationSearch.init({
   *   container:      '.scene-head',
   *   itemSelector:   '.gauge-card',
   *   nameSelector:   '.gc-name',
   *   statusSelector: '.gc-badge',
   * });
   */
  function init(options) {
    injectStyles();
    _config = options || {};

    if (!_config.itemSelector) {
      console.error('[StationSearch] ต้องระบุ itemSelector');
      return;
    }

    /* find host */
    let target = typeof _config.container === 'string'
      ? document.querySelector(_config.container)
      : _config.container;

    if (!target) {
      console.error('[StationSearch] container not found:', _config.container);
      return;
    }

    /* สร้าง search bar แล้ว insert หลัง target */
    _container = document.createElement('div');
    _container.id = 'ssContainer';
    if (target.nextSibling) {
      target.parentNode.insertBefore(_container, target.nextSibling);
    } else {
      target.parentNode.appendChild(_container);
    }

    buildUI(_container);

    /* initial filter */
    setTimeout(_applyFilter, 100);
  }

  /**
   * StationSearch.refresh()
   * เรียกใหม่หลัง render station ใหม่ (เช่นหลัง loadData)
   */
  function refresh() {
    _applyFilter();
  }

  /**
   * StationSearch.reset()
   * ล้างการค้นหาและ filter
   */
  function reset() {
    _searchText   = '';
    _activeStatus = 'all';
    const input = document.getElementById('ssInput');
    if (input) input.value = '';
    const clear = document.getElementById('ssClear');
    if (clear) clear.classList.remove('show');
    _setStatus('all');
  }

  /* ── Export ── */
  global.StationSearch = {
    init:       init,
    refresh:    refresh,
    reset:      reset,
    _setStatus: _setStatus,
  };

})(window);


/* ════════════════════════════════════════════════════════
 * ตัวอย่างการใช้งานในแต่ละหน้า
 * ════════════════════════════════════════════════════════
 *
 * ── paneang.html (และ mong/mo/phuay) ──
 * วางหลัง renderGauges() ครั้งแรก:
 *
 *   StationSearch.init({
 *     container:      '.scene-head',     // วาง search bar ใต้ scene-head
 *     itemSelector:   '.gauge-card',     // การ์ดสถานีแต่ละอัน
 *     nameSelector:   '.gc-name',        // ชื่อสถานีในการ์ด
 *     statusSelector: '.gc-badge',       // badge สถานะในการ์ด
 *   });
 *
 * // หลัง loadData() ทุกครั้ง เรียก:
 *   StationSearch.refresh();
 *
 * ── index.html (รายการสถานีทั้งหมด) ──
 *
 *   StationSearch.init({
 *     container:    '#stationListHeader',
 *     itemSelector: '.station-card',
 *     nameSelector: '.station-name',
 *     statusAttr:   'data-status',       // <div class="station-card" data-status="ปกติ">
 *   });
 *
 * ── reservoir.html ──
 *
 *   StationSearch.init({
 *     container:    '.panel-head',
 *     itemSelector: '.res-item',
 *     nameSelector: '.res-name',
 *   });
 *
 * ════════════════════════════════════════════════════════ */
