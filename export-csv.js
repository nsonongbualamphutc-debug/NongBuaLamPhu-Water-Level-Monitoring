/**
 * export-csv.js
 * เครื่องมือ Export CSV / Excel — ใช้ได้ทุกหน้าในระบบ
 * ระบบสถานการณ์น้ำ จ.หนองบัวลำภู
 * ─────────────────────────────────────────────────────────
 * รองรับภาษาไทย (UTF-8 BOM) เปิดใน Excel ไม่เป็นตัวยุ่ง
 *
 * วิธีใช้:
 *   1. อัปโหลด export-csv.js ขึ้น repo
 *   2. เพิ่ม <script src="export-csv.js"></script> ใน <head>
 *   3. เรียกใช้งาน:
 *
 *   // จาก array ของ object
 *   ExportCSV.fromData(stations, {
 *     filename: 'สถานีลำน้ำพะเนียง',
 *     columns: [
 *       { key:'name',    label:'ชื่อสถานี' },
 *       { key:'current', label:'ระดับน้ำ (ม.รทก.)' },
 *       { key:'status',  label:'สถานะ' },
 *     ]
 *   });
 *
 *   // จาก <table> โดยตรง
 *   ExportCSV.fromTable('#myTable', 'รายงานสถานี');
 *
 *   // สร้างปุ่ม export อัตโนมัติ
 *   ExportCSV.button('#toolbarDiv', () => ExportCSV.fromData(...));
 * ─────────────────────────────────────────────────────────
 */

(function(global) {
  'use strict';

  /* ── UTF-8 BOM — ทำให้ Excel อ่านภาษาไทยถูก ── */
  const BOM = '\uFEFF';

  /* ════════════════════════════════════════
   *  UTILITY
   * ════════════════════════════════════════ */

  /** escape ค่าให้ปลอดภัยใน CSV */
  function escapeCSV(val) {
    if (val == null) return '';
    let s = String(val);
    /* ถ้ามี comma, quote, newline → ครอบด้วย quote */
    if (/[",\n\r]/.test(s)) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  /** สร้างชื่อไฟล์พร้อม timestamp */
  function buildFilename(base, ext) {
    const now = new Date();
    const stamp =
      now.getFullYear() +
      String(now.getMonth() + 1).padStart(2, '0') +
      String(now.getDate()).padStart(2, '0') + '_' +
      String(now.getHours()).padStart(2, '0') +
      String(now.getMinutes()).padStart(2, '0');
    return (base || 'export') + '_' + stamp + '.' + (ext || 'csv');
  }

  /** trigger download */
  function download(content, filename, mime) {
    const blob = new Blob([content], { type: mime || 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(function() {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  }

  /** toast แจ้งผล */
  function toast(msg, ok) {
    let el = document.getElementById('nbp-export-toast');
    if (!el) {
      el = document.createElement('div');
      el.id = 'nbp-export-toast';
      el.style.cssText =
        'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);' +
        'background:#0f172a;color:#fff;padding:12px 22px;border-radius:12px;' +
        'font-family:Sarabun,sans-serif;font-size:13px;font-weight:600;' +
        'box-shadow:0 8px 28px rgba(0,0,0,.3);z-index:99999;' +
        'display:flex;align-items:center;gap:8px;transition:opacity .3s,transform .3s;';
      document.body.appendChild(el);
    }
    el.innerHTML = (ok ? '✅ ' : '⚠️ ') + msg;
    el.style.background = ok ? '#059669' : '#dc2626';
    el.style.opacity = '1';
    el.style.transform = 'translateX(-50%) translateY(0)';
    clearTimeout(el._t);
    el._t = setTimeout(function() {
      el.style.opacity = '0';
      el.style.transform = 'translateX(-50%) translateY(10px)';
    }, 2800);
  }

  /* ════════════════════════════════════════
   *  CORE — สร้าง CSV string
   * ════════════════════════════════════════ */

  /**
   * สร้าง CSV จาก array ของ object
   * @param {Array}  data    — [{ key:value, ... }, ...]
   * @param {Array}  columns — [{ key, label, format? }, ...]
   * @returns {string}
   */
  function buildCSV(data, columns) {
    if (!Array.isArray(data) || data.length === 0) {
      return null;
    }

    /* ถ้าไม่ระบุ columns → ใช้ key ทั้งหมดจาก object แรก */
    if (!columns || !columns.length) {
      columns = Object.keys(data[0]).map(function(k) {
        return { key: k, label: k };
      });
    }

    const lines = [];

    /* header */
    lines.push(columns.map(function(c) {
      return escapeCSV(c.label || c.key);
    }).join(','));

    /* rows */
    data.forEach(function(row) {
      const cells = columns.map(function(c) {
        let val = row[c.key];
        if (typeof c.format === 'function') {
          val = c.format(val, row);
        }
        return escapeCSV(val);
      });
      lines.push(cells.join(','));
    });

    return BOM + lines.join('\r\n');
  }

  /* ════════════════════════════════════════
   *  PUBLIC API
   * ════════════════════════════════════════ */

  /**
   * ExportCSV.fromData(data, options)
   * @param {Array}  data
   * @param {Object} options
   *   options.filename {string}  — ชื่อไฟล์ (ไม่ต้องใส่ .csv)
   *   options.columns  {Array}   — [{ key, label, format }]
   *   options.title    {string}  — แถวหัวเรื่องด้านบน (optional)
   */
  function fromData(data, options) {
    options = options || {};

    let csv = buildCSV(data, options.columns);
    if (!csv) {
      toast('ไม่มีข้อมูลให้ export', false);
      return false;
    }

    /* เพิ่ม title row + วันที่ ด้านบน */
    if (options.title) {
      const meta = BOM +
        escapeCSV(options.title) + '\r\n' +
        escapeCSV('ออกรายงาน: ' + new Date().toLocaleString('th-TH')) + '\r\n' +
        '\r\n';
      csv = meta + csv.slice(BOM.length); /* ตัด BOM ซ้ำ */
    }

    const filename = buildFilename(options.filename, 'csv');
    download(csv, filename, 'text/csv;charset=utf-8;');
    toast('ดาวน์โหลด ' + filename + ' แล้ว', true);
    return true;
  }

  /**
   * ExportCSV.fromTable(tableSelector, filename, options)
   * Export จาก <table> โดยตรง — อ่าน <thead> เป็น header, <tbody> เป็น data
   * ข้ามคอลัมน์ที่มี class "no-export"
   */
  function fromTable(tableSelector, filename, options) {
    options = options || {};
    const table = typeof tableSelector === 'string'
      ? document.querySelector(tableSelector)
      : tableSelector;

    if (!table) {
      toast('ไม่พบตาราง: ' + tableSelector, false);
      return false;
    }

    const lines = [];

    /* header */
    const headCells = table.querySelectorAll('thead th, thead td');
    const skipCols  = []; /* index ของคอลัมน์ที่ข้าม */
    if (headCells.length) {
      const headerRow = [];
      headCells.forEach(function(th, i) {
        if (th.classList.contains('no-export')) {
          skipCols.push(i);
          return;
        }
        headerRow.push(escapeCSV(th.textContent.trim()));
      });
      lines.push(headerRow.join(','));
    }

    /* body */
    const bodyRows = table.querySelectorAll('tbody tr');
    bodyRows.forEach(function(tr) {
      /* ข้ามแถว loading / empty */
      if (tr.classList.contains('no-export')) return;
      if (tr.querySelector('[colspan]')) return;

      const cells = tr.querySelectorAll('td, th');
      const rowData = [];
      cells.forEach(function(td, i) {
        if (skipCols.indexOf(i) !== -1) return;
        /* ใช้ data-export ถ้ามี ไม่งั้นใช้ textContent */
        const val = td.hasAttribute('data-export')
          ? td.getAttribute('data-export')
          : td.textContent.trim().replace(/\s+/g, ' ');
        rowData.push(escapeCSV(val));
      });
      if (rowData.length) lines.push(rowData.join(','));
    });

    if (lines.length <= 1) {
      toast('ตารางไม่มีข้อมูล', false);
      return false;
    }

    let csv = BOM + lines.join('\r\n');

    /* title */
    if (options.title) {
      csv = BOM +
        escapeCSV(options.title) + '\r\n' +
        escapeCSV('ออกรายงาน: ' + new Date().toLocaleString('th-TH')) + '\r\n\r\n' +
        lines.join('\r\n');
    }

    const fname = buildFilename(filename || 'ตาราง', 'csv');
    download(csv, fname, 'text/csv;charset=utf-8;');
    toast('ดาวน์โหลด ' + fname + ' แล้ว', true);
    return true;
  }

  /**
   * ExportCSV.button(container, onClick, options)
   * สร้างปุ่ม export พร้อม style ใส่ใน container
   *
   * @param {string|Element} container
   * @param {Function}       onClick   — สิ่งที่ทำเมื่อกดปุ่ม
   * @param {Object}         options
   *   options.label {string} — ข้อความปุ่ม (default: '📥 Export CSV')
   *   options.theme {string} — 'light' | 'dark' (default: 'light')
   */
  function button(container, onClick, options) {
    options = options || {};
    const el = typeof container === 'string'
      ? document.querySelector(container)
      : container;
    if (!el) {
      console.warn('[ExportCSV] container not found:', container);
      return null;
    }

    const btn = document.createElement('button');
    btn.className = 'nbp-export-btn';
    btn.innerHTML = options.label || '📥 Export CSV';

    const dark = options.theme === 'dark';
    btn.style.cssText =
      'display:inline-flex;align-items:center;gap:6px;' +
      'padding:8px 16px;border-radius:10px;cursor:pointer;' +
      'font-family:Sarabun,sans-serif;font-size:13px;font-weight:600;' +
      'border:1px solid ' + (dark ? 'rgba(255,255,255,.15)' : '#cbd5e1') + ';' +
      'background:' + (dark ? 'rgba(255,255,255,.06)' : '#fff') + ';' +
      'color:' + (dark ? '#e2e8f0' : '#0f172a') + ';' +
      'transition:all .18s;';

    btn.onmouseenter = function() {
      btn.style.background = dark
        ? 'linear-gradient(135deg,#0891b2,#0e7490)'
        : 'linear-gradient(135deg,#0284c7,#0ea5e9)';
      btn.style.color = '#fff';
      btn.style.borderColor = 'transparent';
      btn.style.transform = 'translateY(-1px)';
    };
    btn.onmouseleave = function() {
      btn.style.background = dark ? 'rgba(255,255,255,.06)' : '#fff';
      btn.style.color = dark ? '#e2e8f0' : '#0f172a';
      btn.style.borderColor = dark ? 'rgba(255,255,255,.15)' : '#cbd5e1';
      btn.style.transform = 'translateY(0)';
    };
    btn.onclick = function() {
      try {
        onClick();
      } catch(e) {
        console.error('[ExportCSV] error:', e);
        toast('เกิดข้อผิดพลาดในการ export', false);
      }
    };

    el.appendChild(btn);
    return btn;
  }

  /**
   * ExportCSV.print(options)
   * เปิดหน้าต่างพิมพ์ — ใช้ @media print CSS ของหน้านั้น
   * @param {Object} options
   *   options.title {string} — ตั้งชื่อเอกสารชั่วคราวตอนพิมพ์
   */
  function print(options) {
    options = options || {};
    const originalTitle = document.title;
    if (options.title) {
      document.title = options.title + ' — ' +
        new Date().toLocaleDateString('th-TH');
    }
    window.print();
    setTimeout(function() {
      document.title = originalTitle;
    }, 500);
  }

  /* ── Export ── */
  global.ExportCSV = {
    fromData:  fromData,
    fromTable: fromTable,
    button:    button,
    print:     print,
  };

})(window);


/* ════════════════════════════════════════════════════════
 * ตัวอย่างการใช้งานในแต่ละหน้า
 * ════════════════════════════════════════════════════════
 *
 * ── paneang.html / mong.html / mo.html / phuay.html ──
 * วางในส่วน renderTable() หรือหลัง DOM พร้อม:
 *
 *   ExportCSV.button('.scene-head', function() {
 *     const rows = STATIONS.map(function(s) {
 *       return {
 *         id:      s.id,
 *         name:    s.name,
 *         current: stationData[s.id].current,
 *         bank:    s.bank,
 *         status:  getStatus(s),
 *       };
 *     });
 *     ExportCSV.fromData(rows, {
 *       filename: 'ลำน้ำพะเนียง',
 *       title:    'รายงานระดับน้ำ ลำน้ำพะเนียง จ.หนองบัวลำภู',
 *       columns: [
 *         { key:'id',      label:'รหัสสถานี' },
 *         { key:'name',    label:'ชื่อสถานี' },
 *         { key:'current', label:'ระดับน้ำ (ม.รทก.)' },
 *         { key:'bank',    label:'ระดับตลิ่ง (ม.รทก.)' },
 *         { key:'status',  label:'สถานะ' },
 *       ]
 *     });
 *   });
 *
 * ── หรือ export จากตารางที่มีอยู่แล้ว (ง่ายสุด) ──
 *
 *   ExportCSV.button('#tableHeader', function() {
 *     ExportCSV.fromTable('#stationTable', 'สถานีลำน้ำพะเนียง', {
 *       title: 'รายงานสถานีตรวจวัด ลำน้ำพะเนียง'
 *     });
 *   });
 *
 * ── reservoir.html ──
 *
 *   ExportCSV.button('.tp-head', function() {
 *     ExportCSV.fromTable('table', 'อ่างเก็บน้ำ-หนองบัวลำภู', {
 *       title: 'รายงานอ่างเก็บน้ำ 14 แห่ง จ.หนองบัวลำภู'
 *     });
 *   });
 *
 * ── ปุ่มพิมพ์ PDF ──
 *
 *   ExportCSV.button('.tp-head', function() {
 *     ExportCSV.print({ title: 'รายงานสถานการณ์น้ำ' });
 *   }, { label: '🖨️ พิมพ์รายงาน' });
 *
 * ════════════════════════════════════════════════════════ */
