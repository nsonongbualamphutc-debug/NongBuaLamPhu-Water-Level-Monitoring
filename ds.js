/* ===================================================================
 *  ds.js — Design System กลาง (ส่วน JavaScript)
 *  ศูนย์บัญชาการข้อมูลน้ำ จ.หนองบัวลำภู · สำนักงานสถิติจังหวัด
 *
 *  ระบบ Tooltip กลาง — ใช้ได้ทุกหน้า
 *    • ใส่  data-tip="ข้อความอธิบาย"  ที่ element ไหนก็ได้
 *    • หรือ data-tipid="KEY"  แล้วกำหนด  window.TIPS["KEY"] = "<html>"
 *      สำหรับทูลทิปแบบมีหลอดเกณฑ์/ตาราง
 * =================================================================== */
(function () {
  "use strict";
  if (window.__dsTip) return;          // กันติดตั้งซ้ำ
  window.__dsTip = true;
  window.TIPS = window.TIPS || {};

  function boot() {
    var tip = document.getElementById('tipbox');
    if (!tip) {
      tip = document.createElement('div');
      tip.className = 'tipbox';
      tip.id = 'tipbox';
      document.body.appendChild(tip);
    }
    function place(e) {
      var r = tip.getBoundingClientRect();
      var x = e.clientX + 15, y = e.clientY + 15;
      if (x + r.width  > window.innerWidth  - 8) x = Math.max(8, e.clientX - r.width  - 15);
      if (y + r.height > window.innerHeight - 8) y = Math.max(8, e.clientY - r.height - 15);
      tip.style.left = x + 'px';
      tip.style.top  = y + 'px';
    }
    document.addEventListener('mouseover', function (e) {
      var t = e.target.closest && e.target.closest('[data-tip],[data-tipid]');
      if (!t) return;
      var html = t.dataset.tipid
        ? window.TIPS[t.dataset.tipid]
        : '<div class="simple">' + t.dataset.tip + '</div>';
      if (!html) return;
      tip.innerHTML = html;
      tip.style.display = 'block';
      place(e);
    });
    document.addEventListener('mousemove', function (e) {
      if (tip.style.display === 'block') place(e);
    });
    document.addEventListener('mouseout', function (e) {
      if (e.target.closest && e.target.closest('[data-tip],[data-tipid]')) tip.style.display = 'none';
    });
    window.addEventListener('scroll', function () { tip.style.display = 'none'; }, true);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
