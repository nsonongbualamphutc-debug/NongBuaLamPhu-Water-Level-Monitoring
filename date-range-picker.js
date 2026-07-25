/**
 * date-range-picker.js
 * Date Range Picker สำหรับหน้าลำน้ำ (paneang / mong / mo / phuay)
 * ─────────────────────────────────────────────────────────────────
 * ฟีเจอร์:
 *   - preset buttons: 24 ชม. / 3 วัน / 7 วัน / 14 วัน / 30 วัน / กำหนดเอง
 *   - custom date range picker (จาก–ถึง)
 *   - เรียก callback เมื่อ range เปลี่ยน
 *   - integrate กับ Chart.js chart ที่มีอยู่
 *   - dark theme aware
 *   - responsive mobile
 *
 * วิธีใช้:
 *   1. อัปโหลดไฟล์นี้ขึ้น repo
 *   2. ใน paneang.html เพิ่ม:
 *        <script src="date-range-picker.js"></script>
 *   3. เรียก DateRangePicker.init({...}) หลัง DOM พร้อม
 * ─────────────────────────────────────────────────────────────────
 */

(function(global) {
  'use strict';

  /* ── Styles ── */
  function injectStyles() {
    if (document.getElementById('nbp-drp-style')) return;
    const s = document.createElement('style');
    s.id = 'nbp-drp-style';
    s.textContent = `
/* ─ Date Range Picker container ─ */
.drp-wrap {
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: wrap;
  padding: 10px 16px;
  background: var(--card, #fff);
  border-bottom: 1px solid var(--line, #e2e8f0);
  font-family: 'Sarabun', sans-serif;
}

/* ─ Preset buttons ─ */
.drp-presets {
  display: flex;
  gap: 5px;
  flex-wrap: wrap;
  align-items: center;
  flex: 1;
  min-width: 0;
}
.drp-btn {
  padding: 6px 13px;
  border-radius: 9px;
  border: 1px solid var(--line, #e2e8f0);
  background: transparent;
  color: var(--muted, #64748b);
  font-family: 'Sarabun', sans-serif;
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  transition: all .18s;
  white-space: nowrap;
}
.drp-btn:hover {
  background: var(--info-bg, #eff6ff);
  border-color: var(--brand-2, #0284c7);
  color: var(--brand, #0c4a6e);
}
.drp-btn.active {
  background: linear-gradient(135deg, #0284c7, #0ea5e9);
  border-color: transparent;
  color: #fff;
  box-shadow: 0 3px 10px rgba(14,165,233,.3);
}
[data-theme="dark"] .drp-btn {
  background: rgba(255,255,255,.05);
  border-color: rgba(255,255,255,.12);
  color: rgba(226,232,240,.7);
}
[data-theme="dark"] .drp-btn:hover {
  background: rgba(14,165,233,.15);
  border-color: rgba(14,165,233,.4);
  color: #38bdf8;
}
[data-theme="dark"] .drp-btn.active {
  background: linear-gradient(135deg, #0284c7, #0ea5e9);
  color: #fff;
}

/* ─ Divider ─ */
.drp-div {
  width: 1px;
  height: 24px;
  background: var(--line, #e2e8f0);
  flex-shrink: 0;
}

/* ─ Custom range inputs ─ */
.drp-custom {
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: wrap;
}
.drp-custom label {
  font-size: 11.5px;
  color: var(--muted, #64748b);
  font-weight: 600;
  white-space: nowrap;
}
.drp-input {
  padding: 5px 10px;
  border: 1px solid var(--line, #e2e8f0);
  border-radius: 8px;
  font-family: 'Sarabun', sans-serif;
  font-size: 12px;
  color: var(--ink, #0f172a);
  background: var(--bg, #f4f7fb);
  outline: none;
  transition: border-color .15s;
  cursor: pointer;
}
.drp-input:focus {
  border-color: #0ea5e9;
  box-shadow: 0 0 0 3px rgba(14,165,233,.12);
}
[data-theme="dark"] .drp-input {
  background: rgba(255,255,255,.07);
  border-color: rgba(255,255,255,.12);
  color: #e2e8f0;
  color-scheme: dark;
}
.drp-apply {
  padding: 6px 14px;
  background: linear-gradient(135deg, #0284c7, #0ea5e9);
  border: none;
  border-radius: 9px;
  color: #fff;
  font-family: 'Sarabun', sans-serif;
  font-size: 12px;
  font-weight: 700;
  cursor: pointer;
  transition: all .18s;
  white-space: nowrap;
}
.drp-apply:hover {
  box-shadow: 0 4px 14px rgba(14,165,233,.4);
  transform: translateY(-1px);
}

/* ─ Range display badge ─ */
.drp-badge {
  margin-left: auto;
  padding: 5px 12px;
  background: var(--info-bg, #f0f9ff);
  border: 1px solid rgba(14,165,233,.2);
  border-radius: 8px;
  font-size: 11px;
  font-weight: 600;
  color: #0284c7;
  white-space: nowrap;
  display: flex;
  align-items: center;
  gap: 5px;
  flex-shrink: 0;
}
[data-theme="dark"] .drp-badge {
  background: rgba(14,165,233,.12);
  border-color: rgba(14,165,233,.3);
  color: #38bdf8;
}

/* ─ Loading state ─ */
.drp-loading {
  font-size: 11px;
  color: var(--muted, #64748b);
  padding: 4px 8px;
  display: none;
}
.drp-loading.show { display: block; }
.drp-loading::before {
  content: '';
  display: inline-block;
  width: 10px; height: 10px;
  border: 2px solid rgba(14,165,233,.2);
  border-top-color: #0ea5e9;
  border-radius: 50%;
  animation: drpSpin .7s linear infinite;
  margin-right: 5px;
  vertical-align: middle;
}
@keyframes drpSpin { to { transform: rotate(360deg); } }

/* ─ Responsive ─ */
@media (max-width: 600px) {
  .drp-wrap { gap: 6px; padding: 8px 12px; }
  .drp-div  { display: none; }
  .drp-custom { width: 100%; }
  .drp-badge  { width: 100%; justify-content: center; }
}
`;
    document.head.appendChild(s);
  }

  /* ── PRESETS ── */
  const PRESETS = [
    { label: '24 ชม.',  days: 1,  id: '1d'  },
    { label: '3 วัน',   days: 3,  id: '3d'  },
    { label: '7 วัน',   days: 7,  id: '7d'  },
    { label: '14 วัน',  days: 14, id: '14d' },
    { label: '30 วัน',  days: 30, id: '30d' },
  ];

  /* ── State ── */
  let _config   = null;
  let _current  = { preset: '7d', from: null, to: null };
  let _container = null;

  /* ── Date helpers ── */
  function toDateStr(d) {
    return d.toISOString().slice(0, 10);
  }
  function fromDays(n) {
    const d = new Date();
    d.setDate(d.getDate() - n);
    return d;
  }
  function formatThai(dateStr) {
    const d = new Date(dateStr);
    return d.toLocaleDateString('th-TH', { day:'numeric', month:'short', year:'2-digit' });
  }
  function daysBetween(a, b) {
    return Math.round((new Date(b) - new Date(a)) / 86400000);
  }

  /* ── Build HTML ── */
  function buildUI(container) {
    const today = toDateStr(new Date());
    const d7ago = toDateStr(fromDays(7));

    container.innerHTML = `
<div class="drp-wrap" id="drpWrap">
  <div class="drp-presets" id="drpPresets">
    ${PRESETS.map(p =>
      `<button class="drp-btn${p.id === _current.preset ? ' active' : ''}"
               data-preset="${p.id}" data-days="${p.days}"
               onclick="DateRangePicker._onPreset('${p.id}', ${p.days})"
      >${p.label}</button>`
    ).join('')}
    <button class="drp-btn${_current.preset === 'custom' ? ' active' : ''}"
            data-preset="custom"
            onclick="DateRangePicker._toggleCustom()">📅 กำหนดเอง</button>
  </div>

  <div class="drp-div"></div>

  <div class="drp-custom" id="drpCustomPanel" style="display:none">
    <label>จาก</label>
    <input type="date" class="drp-input" id="drpFrom"
           value="${d7ago}" max="${today}">
    <label>ถึง</label>
    <input type="date" class="drp-input" id="drpTo"
           value="${today}" max="${today}">
    <button class="drp-apply" onclick="DateRangePicker._applyCustom()">ดูข้อมูล</button>
  </div>

  <span class="drp-loading" id="drpLoading">กำลังโหลด...</span>

  <div class="drp-badge" id="drpBadge">
    📅 <span id="drpBadgeTxt">7 วันล่าสุด</span>
  </div>
</div>`;

    /* limit max date on both inputs */
    document.getElementById('drpFrom').addEventListener('change', function() {
      const toEl = document.getElementById('drpTo');
      if (toEl.value < this.value) toEl.value = this.value;
    });
    document.getElementById('drpTo').addEventListener('change', function() {
      const fromEl = document.getElementById('drpFrom');
      if (fromEl.value > this.value) fromEl.value = this.value;
    });
  }

  /* ── event handlers ── */
  function _onPreset(presetId, days) {
    _current.preset = presetId;
    _current.from   = toDateStr(fromDays(days));
    _current.to     = toDateStr(new Date());

    /* update active button */
    document.querySelectorAll('.drp-btn').forEach(function(b) {
      b.classList.toggle('active', b.dataset.preset === presetId);
    });

    /* hide custom panel */
    const panel = document.getElementById('drpCustomPanel');
    if (panel) panel.style.display = 'none';

    /* update badge */
    _updateBadge(days + ' วันล่าสุด');

    /* call back */
    _triggerChange();
  }

  function _toggleCustom() {
    const panel = document.getElementById('drpCustomPanel');
    if (!panel) return;
    const show = panel.style.display === 'none';
    panel.style.display = show ? 'flex' : 'none';
    if (show) {
      /* set active */
      document.querySelectorAll('.drp-btn').forEach(function(b) {
        b.classList.toggle('active', b.dataset.preset === 'custom');
      });
    }
  }

  function _applyCustom() {
    const fromEl = document.getElementById('drpFrom');
    const toEl   = document.getElementById('drpTo');
    if (!fromEl || !toEl) return;

    const from = fromEl.value;
    const to   = toEl.value;

    if (!from || !to || from > to) {
      _showError('กรุณาเลือกวันที่ให้ถูกต้อง');
      return;
    }

    const days = daysBetween(from, to) + 1;
    if (days > 90) {
      _showError('ช่วงสูงสุด 90 วัน');
      return;
    }

    _current.preset = 'custom';
    _current.from   = from;
    _current.to     = to;

    _updateBadge(formatThai(from) + ' – ' + formatThai(to) + ' (' + days + ' วัน)');
    _triggerChange();
  }

  function _updateBadge(text) {
    const el = document.getElementById('drpBadgeTxt');
    if (el) el.textContent = text;
  }

  function _showError(msg) {
    const el = document.getElementById('drpBadgeTxt');
    if (el) {
      el.textContent = '⚠️ ' + msg;
      setTimeout(function() {
        el.textContent = _currentLabel();
      }, 3000);
    }
  }

  function _currentLabel() {
    if (_current.preset !== 'custom') {
      const p = PRESETS.find(function(x) { return x.id === _current.preset; });
      return p ? p.days + ' วันล่าสุด' : '';
    }
    return formatThai(_current.from) + ' – ' + formatThai(_current.to);
  }

  /* ── loading indicator ── */
  function _setLoading(on) {
    const el = document.getElementById('drpLoading');
    if (el) el.classList.toggle('show', on);
    document.querySelectorAll('.drp-btn, .drp-apply').forEach(function(b) {
      b.disabled = on;
    });
  }

  /* ── trigger onChange callback ── */
  function _triggerChange() {
    if (!_config || typeof _config.onChange !== 'function') return;
    _setLoading(true);
    try {
      const result = _config.onChange({
        from:   _current.from,
        to:     _current.to,
        preset: _current.preset,
        days:   daysBetween(_current.from, _current.to) + 1,
      });
      /* ถ้า callback คืน Promise */
      if (result && typeof result.then === 'function') {
        result.finally(function() { _setLoading(false); });
      } else {
        _setLoading(false);
      }
    } catch(e) {
      _setLoading(false);
      console.error('[DateRangePicker] onChange error:', e);
    }
  }

  /* ════════════════════════════════════
   *  PUBLIC API
   * ════════════════════════════════════ */

  /**
   * DateRangePicker.init(options)
   *
   * @param {Object} options
   *   options.container   {string|Element}  — selector หรือ element ที่จะ inject UI เข้าไป
   *   options.defaultDays {number}          — preset เริ่มต้น (default: 7)
   *   options.onChange    {Function}        — callback({ from, to, preset, days })
   *                                           รับ date string 'YYYY-MM-DD'
   *                                           สามารถ return Promise
   *
   * ตัวอย่าง:
   * ──────────
   * DateRangePicker.init({
   *   container:   '#chartSection',   // จะ prepend ก่อน element นี้
   *   defaultDays: 7,
   *   onChange: async function({ from, to, days }) {
   *     const data = await fetchWaterLevel(from, to);
   *     updateChart(data);
   *   }
   * });
   */
  function init(options) {
    injectStyles();
    _config = options || {};

    const defaultDays = _config.defaultDays || 7;
    _current.from   = toDateStr(fromDays(defaultDays));
    _current.to     = toDateStr(new Date());
    _current.preset = defaultDays + 'd';

    /* find container */
    let target;
    if (typeof _config.container === 'string') {
      target = document.querySelector(_config.container);
    } else {
      target = _config.container;
    }

    if (!target) {
      console.error('[DateRangePicker] container not found:', _config.container);
      return;
    }

    /* สร้าง wrapper div แล้ว insert ก่อน target */
    _container = document.createElement('div');
    _container.id = 'drpContainer';
    target.parentNode.insertBefore(_container, target);

    buildUI(_container);

    /* trigger initial load */
    if (_config.autoLoad !== false) {
      setTimeout(function() { _triggerChange(); }, 100);
    }
  }

  /**
   * DateRangePicker.getRange()
   * คืน { from, to, preset, days } ของ range ปัจจุบัน
   */
  function getRange() {
    return {
      from:   _current.from,
      to:     _current.to,
      preset: _current.preset,
      days:   daysBetween(_current.from, _current.to) + 1,
    };
  }

  /**
   * DateRangePicker.setLoading(bool)
   * เรียกจากภายนอกเพื่อ control loading state
   */
  function setLoading(on) {
    _setLoading(on);
  }

  /* ── Export ── */
  global.DateRangePicker = {
    init:          init,
    getRange:      getRange,
    setLoading:    setLoading,
    /* internal (เรียกจาก onclick inline) */
    _onPreset:     _onPreset,
    _toggleCustom: _toggleCustom,
    _applyCustom:  _applyCustom,
  };

})(window);


/* ════════════════════════════════════════════════════════
 * ตัวอย่าง integration ใน paneang.html / mong.html / mo.html / phuay.html
 * ════════════════════════════════════════════════════════
 *
 * 1. เพิ่ม <script> ใน <head>:
 *      <script src="date-range-picker.js"></script>
 *
 * 2. ใน JS section เพิ่ม init() หลัง DOM load:
 *
 * DateRangePicker.init({
 *   container:   '#chartPanel',       // section ที่จะวาง picker ไว้ด้านบน
 *   defaultDays: 7,
 *   onChange: async function({ from, to, days }) {
 *     // เรียก API พร้อม date range
 *     const url = API_URL +
 *       '?action=water' +
 *       '&river=paneang' +
 *       '&from=' + from +
 *       '&to='   + to +
 *       '&callback=cb_water';
 *
 *     const data = await fetchJSONP(url);
 *     updateChart(data, from, to);
 *     updateStationCards(data.stations);
 *   }
 * });
 *
 * หรือถ้า fetch แบบ JSONP ทั่วไป:
 *
 * DateRangePicker.init({
 *   container: '#historySection',
 *   defaultDays: 7,
 *   onChange: function({ from, to }) {
 *     loadHistoryData(from, to);
 *     return new Promise(function(resolve) {
 *       // resolve() เมื่อโหลดเสร็จ เพื่อซ่อน loading
 *       window.__drpResolve = resolve;
 *     });
 *   }
 * });
 *
 * // หลัง render เสร็จ:
 * if (window.__drpResolve) {
 *   window.__drpResolve();
 *   window.__drpResolve = null;
 * }
 *
 * ════════════════════════════════════════════════════════ */
