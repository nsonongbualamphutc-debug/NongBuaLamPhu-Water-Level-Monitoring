/* ════════════════════════════════════════════════════════════
 *  PATCH v2 — verdict.js
 *  คอมโพเนนต์ "คำตัดสินสถานการณ์" ใช้ร่วมทุกหน้า
 *
 *  หลักการ: ไม่แตะโค้ดเดิมของหน้า — ใช้วิธี "wrap" ฟังก์ชัน render
 *  ที่หน้านั้นมีอยู่แล้ว ให้เรียกของเดิมก่อน แล้วค่อยวาด verdict ต่อ
 *  → ถ้าลบไฟล์นี้ออก ทุกอย่างกลับเป็นเหมือนเดิม 100%
 *
 *  โหมด (อ่านจาก <div id="verdictZone" data-vd="...">):
 *    data-vd="index"     → hero เต็ม + มาตรตลิ่ง + จับตา + เช็คอำเภอ
 *    data-vd="river"     → แถบย่อรายลำน้ำ (data-river="ลำน้ำมอ")
 *    data-vd="reservoir" → แถบย่ออ่างเก็บน้ำ
 * ════════════════════════════════════════════════════════════ */
(function () {
  "use strict";
  var zone = document.getElementById("verdictZone");
  if (!zone) return;
  var MODE = zone.getAttribute("data-vd") || "index";
  var RIVER_NAME = zone.getAttribute("data-river") || "";
  var SITE = "https://nsonongbualamphutc-debug.github.io/NongBuaLamPhu-Water-Level-Monitoring/";
  var AMPHOES = ["เมืองหนองบัวลำภู", "นาวัง", "นากลาง", "ศรีบุญเรือง", "สุวรรณคูหา", "โนนสัง"];

  /* ── helpers ── */
  function num(v) { var n = parseFloat(v); return isNaN(n) ? null : n; }
  function stOf(s) {
    var cur = num(s.current), bank = num(s.crit) || num(s.bank), warn = num(s.warn);
    if (cur === null) return "normal";
    if (bank && cur >= bank) return "crit";
    if (warn && cur >= warn) return "warn";
    return "normal";
  }
  function thDate() {
    var n = new Date();
    var day = ["อาทิตย์","จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์"][n.getDay()];
    return "วัน" + day + "ที่ " + n.toLocaleDateString("th-TH", {day:"numeric",month:"long",year:"numeric"}) +
           " · " + n.toLocaleTimeString("th-TH", {hour:"2-digit",minute:"2-digit",hour12:false}) + " น.";
  }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
    return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]; }); }
  function toast(msg) {
    var t = document.querySelector(".vd-toast");
    if (!t) { t = document.createElement("div"); t.className = "vd-toast"; document.body.appendChild(t); }
    t.textContent = msg; t.classList.add("show");
    setTimeout(function () { t.classList.remove("show"); }, 2600);
  }
  function copyText(txt) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(txt).then(function(){ toast("คัดลอกรายงานแล้ว ✅ วางในไลน์กลุ่มได้เลย"); },
        function(){ fallbackCopy(txt); });
    } else fallbackCopy(txt);
  }
  function fallbackCopy(txt) {
    var ta = document.createElement("textarea");
    ta.value = txt; ta.style.position = "fixed"; ta.style.opacity = "0";
    document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); toast("คัดลอกรายงานแล้ว ✅"); }
    catch (e) { toast("คัดลอกไม่สำเร็จ — กรุณาคัดลอกเอง"); }
    document.body.removeChild(ta);
  }

  /* วิเคราะห์ภาพรวมจาก stations — พร้อมกรองคุณภาพข้อมูล */
  var STALE_MS = 48 * 3600 * 1000; /* ข้อมูลเก่ากว่า 48 ชม. = ค้าง ไม่นำมาตัดสิน */
  function parseTS(v) {
    if (!v) return null;
    var s = String(v).trim();
    var d = new Date(s);
    if (!isNaN(d.getTime())) return d.getTime();
    d = new Date(s.replace(" ", "T"));
    if (!isNaN(d.getTime())) return d.getTime();
    /* รูปแบบจากชีต เช่น "2026-05-28T00:00:00.000Z 14:19" → จับวันที่+เวลาแยก */
    var m = s.match(/(\d{4}-\d{2}-\d{2}).*?(\d{1,2}:\d{2})/);
    if (m) {
      var t = m[2].length === 4 ? "0" + m[2] : m[2];
      d = new Date(m[1] + "T" + t + ":00");
      if (!isNaN(d.getTime())) return d.getTime();
    }
    return null;
  }
  /* ตรวจค่าที่เป็นไปไม่ได้ทางกายภาพ (เช่น กรอกผิดหน่วย 500 ม.รทก. ทั้งที่ตลิ่ง 211) */
  function isSuspect(s) {
    var cur = num(s.current);
    if (cur === null) return false;
    if (MODE === "reservoir") return cur < 0 || cur > 130;          /* % เก็บกัก */
    var bank = num(s.crit) || num(s.bank);
    if (!bank) return false;
    var gap = bank - cur;
    return cur <= 0 || gap < -10 || gap > 100;  /* เกินตลิ่ง >10 ม. หรือต่ำกว่า >100 ม. = กรอกผิดแน่ */
  }
  function analyze(stations, rain) {
    var suspects = [], stale = [];
    var live = (stations || []).filter(function (s) {
      if (num(s.current) === null) return false;
      if (isSuspect(s)) { suspects.push(s); return false; }
      var ts = parseTS(s._lastUpdate || s.last_update);
      if (ts && (Date.now() - ts) > STALE_MS) { stale.push(s); return false; }
      return true;
    });
    var counts = { normal: 0, warn: 0, crit: 0 };
    live.forEach(function (s) { counts[stOf(s)]++; });
    /* จุดใกล้ตลิ่งที่สุด (ระยะ bank - current น้อยสุด) */
    var nearest = null;
    live.forEach(function (s) {
      var bank = num(s.crit) || num(s.bank); if (!bank) return;
      var gap = bank - num(s.current);
      if (!nearest || gap < nearest.gap) nearest = { s: s, gap: gap, bank: bank, cur: num(s.current) };
    });
    /* ฝนสูงสุด */
    var maxRain = null, avgRain = 0;
    (rain || []).forEach(function (r) {
      var v = num(r.rain24 != null ? r.rain24 : r.rain_24hr) || 0;
      avgRain += v;
      if (!maxRain || v > maxRain.v) maxRain = { amphoe: r.amphoe, v: v };
    });
    if (rain && rain.length) avgRain = avgRain / rain.length;
    /* ระดับสถานการณ์รวม */
    var level = counts.crit > 0 ? "crit" : counts.warn > 0 ? "warn" : "normal";
    /* lead = สถานี/พื้นที่ที่เป็นประเด็น */
    var hot = live.filter(function (s) { return stOf(s) !== "normal"; })
      .sort(function (a, b) { return ((num(a.crit)||num(a.bank)) - num(a.current)) - ((num(b.crit)||num(b.bank)) - num(b.current)); });
    return { live: live, counts: counts, nearest: nearest, maxRain: maxRain, avgRain: avgRain, level: level, hot: hot,
             total: (stations || []).length, suspects: suspects, stale: stale };
  }

  function levelWord(l) { return { normal: "ปกติ", warn: "เฝ้าระวัง", crit: "วิกฤติ" }[l]; }
  function levelEmoji(l) { return { normal: "🟢", warn: "⚠️", crit: "🔴" }[l]; }

  /* ── สร้างข้อความรายงานราชการ (ปุ่มลดงาน) ── */
  function buildReport(a, scopeLabel) {
    var L = [];
    L.push("📋 รายงานสถานการณ์น้ำ" + (scopeLabel ? " " + scopeLabel : "") + " จ.หนองบัวลำภู");
    L.push("ประจำ" + thDate());
    L.push("");
    L.push("ระดับสถานการณ์: " + levelEmoji(a.level) + " " + levelWord(a.level) +
      (a.hot.length ? " (" + a.hot.slice(0, 2).map(function (s) { return s.name; }).join(", ") + ")" : ""));
    L.push("• สถานีปกติ " + a.counts.normal + " / เฝ้าระวัง " + a.counts.warn + " / วิกฤติ " + a.counts.crit +
      (a.total > a.live.length ? " (รอข้อมูล " + (a.total - a.live.length) + ")" : ""));
    if (a.nearest) {
      L.push("• จุดใกล้ตลิ่งที่สุด: " + a.nearest.s.name + " (" + (a.nearest.s.id || "") + ") " +
        (a.nearest.gap >= 0 ? "ต่ำกว่าตลิ่ง " + a.nearest.gap.toFixed(2) + " ม." : "‼️ เกินตลิ่ง " + Math.abs(a.nearest.gap).toFixed(2) + " ม."));
    }
    if (a.maxRain && a.maxRain.v > 0) L.push("• ฝนสะสมสูงสุด: อ." + a.maxRain.amphoe + " " + a.maxRain.v.toFixed(1) + " มม./24 ชม.");
    if (a.avgRain > 0) L.push("• ฝนเฉลี่ยทั้งจังหวัด " + a.avgRain.toFixed(1) + " มม./24 ชม.");
    if (a.suspects && a.suspects.length) {
      L.push("• ⚠️ พบข้อมูลผิดปกติรอตรวจสอบ " + a.suspects.length + " สถานี: " +
        a.suspects.map(function (s) { return s.id || s.name; }).join(", ") + " (ไม่นำมาคำนวณ)");
    }
    if (a.stale && a.stale.length) {
      L.push("• 🕐 ข้อมูลค้างเกิน 48 ชม. " + a.stale.length + " สถานี: " +
        a.stale.map(function (s) { return s.id || s.name; }).join(", "));
    }
    L.push("");
    L.push("ดูข้อมูลเรียลไทม์: " + SITE);
    L.push("ที่มา: สำนักงานสถิติจังหวัดหนองบัวลำภู");
    return L.join("\n");
  }

  /* ── render: มาตรตลิ่ง ── */
  function gaugeHTML(a) {
    if (!a.nearest) return "";
    var n = a.nearest, gap = n.gap;
    /* ตำแหน่งผิวน้ำ — clamp ให้ป้ายระยะ (สูง ~30px) อยู่ในกรอบ gauge 110px เสมอ */
    var pctTop = +(Math.max(12, Math.min(60, 86 - ((n.cur / n.bank) * 70)))).toFixed(1);
    var over = gap < 0;
    if (MODE === "reservoir") {
      /* โหมดอ่าง: ตีความเป็น % เก็บกัก */
      return '<div class="vd-gauge-card">' +
        '<div class="vd-gauge-head"><h2>📏 อ่างที่ % เก็บกักสูงสุดตอนนี้</h2>' +
        '<span class="st">' + esc(n.s.name) + (n.s.amphoe ? " · อ." + esc(n.s.amphoe) : "") + '</span></div>' +
        '<div class="vd-gauge">' +
          '<div class="water" style="top:' + pctTop + '%"></div>' +
          '<div class="bank"><span>🚩 ความจุ 100%</span></div>' +
          '<div class="dist' + (over ? " over" : "") + '" style="top:' + pctTop + '%"><b>' +
            (over ? "‼️ เกินความจุ" : "เก็บกัก " + n.cur.toFixed(0) + "% · เหลือรับน้ำ " + gap.toFixed(0) + "%") + '</b></div>' +
        '</div></div>';
    }
    return '<div class="vd-gauge-card">' +
      '<div class="vd-gauge-head"><h2>📏 จุดที่ใกล้ตลิ่งที่สุด' + (MODE === "river" ? "ของลำน้ำนี้" : "ของจังหวัดตอนนี้") + '</h2>' +
      '<span class="st">' + esc(n.s.name) + (n.s.id ? " · " + esc(n.s.id) : "") + (n.s.amphoe ? " · อ." + esc(n.s.amphoe) : "") + '</span></div>' +
      '<div class="vd-gauge">' +
        '<div class="water" style="top:' + pctTop + '%"></div>' +
        '<div class="bank"><span>🚩 ตลิ่ง ' + n.bank.toFixed(2) + ' ม.</span></div>' +
        '<div class="dist' + (over ? " over" : "") + '" style="top:' + pctTop + '%"><b>' +
          (over ? "‼️ เกินตลิ่ง " + Math.abs(gap).toFixed(2) + " ม." : "เหลือ " + gap.toFixed(2) + " ม.") + '</b></div>' +
      '</div>' +
      '<div class="vd-gauge-foot"><span>ระดับน้ำปัจจุบัน <b>' + n.cur.toFixed(2) + ' ม.รทก.</b></span>' +
      '<span>' + esc(n.s._lastUpdate || n.s.last_update || "") + '</span></div></div>';
  }

  /* ── render: จับตาวันนี้ (index เท่านั้น) ── */
  function watchHTML(a) {
    var cards = [];
    a.hot.slice(0, 2).forEach(function (s) {
      var st = stOf(s), gap = (num(s.crit) || num(s.bank)) - num(s.current);
      cards.push({ cls: st, tag: levelWord(st), h: s.name, p: "ลำ" + (s.river || "น้ำ") + " · อ." + (s.amphoe || "-") + " ควรติดตามต่อเนื่อง",
        num: gap.toFixed(2), unit: "ม. จากตลิ่ง" });
    });
    if (a.maxRain && a.maxRain.v >= 35) {
      cards.push({ cls: a.maxRain.v >= 90 ? "crit" : "warn", tag: "ฝนหนัก", h: "อ." + a.maxRain.amphoe,
        p: "ฝนสะสม 24 ชม. สูงสุดของจังหวัด", num: a.maxRain.v.toFixed(0), unit: "มม. / 24 ชม." });
    }
    if (!cards.length) {
      cards.push({ cls: "ok", tag: "ปกติ", h: "ไม่มีจุดต้องเฝ้าระวังพิเศษ",
        p: "ทุกสถานีต่ำกว่าระดับเฝ้าระวัง — ติดตามรอบถัดไปตามปกติ", num: a.counts.normal + "/" + a.live.length, unit: "สถานีปกติ" });
    }
    var html = cards.slice(0, 3).map(function (c) {
      return '<div class="vd-wcard ' + c.cls + '"><span class="tag">' + esc(c.tag) + '</span><h3>' + esc(c.h) +
        '</h3><p>' + esc(c.p) + '</p><div class="num">' + esc(c.num) + ' <small>' + esc(c.unit) + '</small></div></div>';
    }).join("");
    return '<div class="vd-watch">' + html + '</div>';
  }

  /* ── render: เช็คพื้นที่ของท่าน (index) ── */
  function ampHTML(a, rain) {
    var byAmp = {};
    AMPHOES.forEach(function (am) { byAmp[am] = "normal"; });
    a.live.forEach(function (s) {
      var am = s.amphoe === "เมือง" ? "เมืองหนองบัวลำภู" : s.amphoe;
      if (!(am in byAmp)) return;
      var st = stOf(s);
      if (st === "crit" || (st === "warn" && byAmp[am] === "normal")) byAmp[am] = st;
    });
    (rain || []).forEach(function (r) {
      var am = r.amphoe === "เมือง" ? "เมืองหนองบัวลำภู" : r.amphoe;
      var v = num(r.rain24 != null ? r.rain24 : r.rain_24hr) || 0;
      if (am in byAmp && v >= 90 && byAmp[am] === "normal") byAmp[am] = "warn";
    });
    var msg = { normal: "ใช้ชีวิตได้ตามปกติ", warn: "ติดตามประกาศจาก อปท. ในพื้นที่", crit: "ปฏิบัติตามคำแนะนำของเจ้าหน้าที่" };
    var html = AMPHOES.map(function (am) {
      var st = byAmp[am];
      return '<div class="vd-amp"><div class="nm">' + esc(am) + '</div>' +
        '<span class="stt vd-s-' + st + '">' + levelWord(st) + '</span>' +
        '<div class="msg">' + msg[st] + '</div></div>';
    }).join("");
    return '<div class="vd-amp-zone"><h2>🏡 เช็คพื้นที่ของท่าน</h2>' +
      '<p>สรุปสถานะรายอำเภอเป็นภาษาง่าย ๆ สำหรับประชาชน</p>' +
      '<div class="vd-amp-grid">' + html + '</div></div>';
  }

  /* ── ประโยคสรุป ── */
  function leadSentence(a) {
    if (a.level === "normal") {
      if (MODE === "reservoir") return "อ่างเก็บน้ำทุกแห่งมี % เก็บกักอยู่ในเกณฑ์ <b>ปกติ</b> ยังมีพื้นที่รองรับน้ำฝนได้";
      if (MODE === "river") return "ทุกสถานีตรวจวัดใน" + esc(RIVER_NAME) + "ต่ำกว่าระดับเฝ้าระวัง <b>อยู่ในเกณฑ์ปกติ</b>";
      return "ทุกสถานีตรวจวัดต่ำกว่าระดับเฝ้าระวัง อ่างเก็บน้ำและลำน้ำทั้ง 4 สาย <b>อยู่ในเกณฑ์ปกติ</b>" +
        (a.avgRain > 0 ? " ฝนเฉลี่ย 24 ชม. " + a.avgRain.toFixed(1) + " มม." : "");
    }
    var names = a.hot.slice(0, 2).map(function (s) { return "<b>" + esc(s.name) + (s.id && MODE !== "reservoir" ? " (" + esc(s.id) + ")" : "") + "</b>"; }).join(" และ ");
    if (MODE === "reservoir") {
      return names + " มี % เก็บกักสูงในเกณฑ์" + levelWord(a.level) + " ควรติดตามการระบายน้ำ — อ่างที่เหลือ <b>อยู่ในเกณฑ์ปกติ</b>";
    }
    var rainTxt = (a.maxRain && a.maxRain.v >= 35) ? "ฝนตกต่อเนื่องในเขต อ." + esc(a.maxRain.amphoe) + " ส่งผลให้" : "";
    return rainTxt + "สถานี " + names + " อยู่ในเกณฑ์" + levelWord(a.level) +
      (a.nearest && a.nearest.gap >= 0 ? " เหลือต่ำกว่าตลิ่ง " + a.nearest.gap.toFixed(2) + " ม." : "") +
      " — สถานีที่เหลือ <b>อยู่ในเกณฑ์ปกติ</b>";
  }

  function hotScope(a) {
    if (a.level === "normal") return "";
    var s = a.hot[0]; if (!s) return "";
    return s.river ? " ลำ" + s.river.replace(/^ลำ/, "") : (s.amphoe ? " อ." + s.amphoe : "");
  }

  /* ── MAIN RENDER ── */
  var lastData = { stations: null, rain: null };
  function renderVerdict(stations, rain) {
    try {
      lastData.stations = stations; lastData.rain = rain;
      var a = analyze(stations, rain);
      var compact = MODE !== "index";
      var scopeName = MODE === "river" ? RIVER_NAME : MODE === "reservoir" ? "อ่างเก็บน้ำ" : "";
      var excluded = a.suspects.length + a.stale.length;
      var freshTxt = (a.live.length >= a.total && a.total > 0)
        ? "กรอกครบ " + a.live.length + "/" + a.total + " สถานี"
        : "ข้อมูลปัจจุบัน " + a.live.length + "/" + a.total + " สถานี";

      /* แถบเตือนคุณภาพข้อมูล — ฟ้องให้ตรวจสอบ ไม่ปล่อยให้ค่าผิดมาตัดสินสถานการณ์ */
      var dqHTML = "";
      if (excluded > 0) {
        var parts = [];
        if (a.suspects.length) parts.push("ค่าผิดปกติ <b>" + a.suspects.length + " สถานี</b> (" +
          a.suspects.slice(0, 3).map(function (s) { return esc(s.id || s.name); }).join(", ") +
          (a.suspects.length > 3 ? " ฯลฯ" : "") + ") — ตัวเลขเป็นไปไม่ได้ทางกายภาพ กรุณาตรวจสอบการกรอก");
        if (a.stale.length) parts.push("ข้อมูลค้างเกิน 48 ชม. <b>" + a.stale.length + " สถานี</b> (" +
          a.stale.slice(0, 3).map(function (s) { return esc(s.id || s.name); }).join(", ") +
          (a.stale.length > 3 ? " ฯลฯ" : "") + ")");
        dqHTML = '<div class="vd-dq">⚠️ ไม่ถูกนำมาคำนวณสถานการณ์: ' + parts.join(" · ") + '</div>';
      }

      var html = '<div class="vd-hero' + (compact ? " compact" : "") + '">' +
        '<div class="vd-fresh' + (a.live.length < a.total ? " stale" : "") + '"><span class="dot"></span> ' + freshTxt + '</div>' +
        '<div class="vd-tag">— สรุปสถานการณ์' + (scopeName ? esc(scopeName) : "") + ' ประจำ' + thDate() + '</div>' +
        '<h1 class="vd-verdict">วันนี้: <span class="lv-' + a.level + '">' + levelWord(a.level) + '</span>' +
        esc(hotScope(a)) + '</h1>' +
        '<p class="vd-sub">' + leadSentence(a) + '</p>' +
        '<div class="vd-meta"><div>ปกติ <b>' + a.counts.normal + '</b></div><div>เฝ้าระวัง <b>' + a.counts.warn +
        '</b></div><div>วิกฤติ <b>' + a.counts.crit + '</b></div>' +
        (a.avgRain > 0 ? '<div>ฝนเฉลี่ย 24 ชม. <b>' + a.avgRain.toFixed(1) + ' มม.</b></div>' : "") + '</div>' +
        dqHTML +
        '<div class="vd-actions"><button class="vd-btn vd-btn-line" id="vdCopyBtn">📋 คัดลอกรายงานเช้า</button>' +
        (MODE === "index" ? '<a class="vd-btn vd-btn-ghost" href="daily_briefing.html">📊 รายงานประจำวันฉบับเต็ม</a>' : "") +
        '</div></div>';

      html += gaugeHTML(a);
      if (MODE === "index") {
        html += watchHTML(a);
        html += ampHTML(a, rain);
      }
      zone.innerHTML = html;
      var btn = document.getElementById("vdCopyBtn");
      if (btn) btn.onclick = function () { copyText(buildReport(a, scopeName)); };
    } catch (e) { console.warn("verdict render error", e); }
  }

  /* ════════════════════════════════════════════
   *  HOOK เข้ากับแต่ละหน้า (ไม่แก้โค้ดเดิม)
   * ════════════════════════════════════════════ */
  if (MODE === "index") {
    /* index.html: wrap window.render(stations, rain) */
    if (typeof window.render === "function") {
      var _origRender = window.render;
      window.render = function (stations, rain) {
        _origRender.apply(this, arguments);
        renderVerdict(stations, rain);
      };
    }
  } else if (MODE === "river") {
    /* หน้าลำน้ำ: STATIONS/stationData เป็น const/let (lexical global) — อ้างชื่อตรงผ่าน indirect eval ไม่ได้
       ใช้ try + ชื่อ bare แทน window.* */
    function riverStations() {
      try {
        /* eslint-disable no-undef */
        var STS = (typeof STATIONS !== "undefined") ? STATIONS : [];
        var SD  = (typeof stationData !== "undefined") ? stationData : {};
        return STS.map(function (s) {
          var live = SD[s.id] || {};
          return Object.assign({}, s, {
            current: (live.current != null) ? live.current : s.current,
            _isLive: !!live._isLive
          });
        });
      } catch (e) { return []; }
    }
    if (typeof window.renderSummary === "function") {
      var _origRS = window.renderSummary;
      window.renderSummary = function () {
        _origRS.apply(this, arguments);
        renderVerdict(riverStations(), null);
      };
    }
    /* เผื่อ renderSummary ถูกเรียกไปแล้วก่อน script นี้โหลด */
    setTimeout(function () { if (!zone.innerHTML) renderVerdict(riverStations(), null); }, 1200);
  } else if (MODE === "reservoir") {
    /* อ่างเก็บน้ำ: RESERVOIRS (เมตา, capacity ล้าน ลบ.ม.) + resData[id].current (ปริมาณปัจจุบัน)
       แปลงเป็น % เก็บกักเทียบเคียงรูปแบบ stations */
    function resStations() {
      try {
        /* eslint-disable no-undef */
        var RES = (typeof RESERVOIRS !== "undefined") ? RESERVOIRS : [];
        var RD  = (typeof resData !== "undefined") ? resData : {};
        return RES.map(function (r) {
          var cap = num(r.capacity) || 0;
          var cur = (RD[r.id] && RD[r.id].current != null) ? num(RD[r.id].current) : null;
          var pct = (cap && cur != null) ? (cur / cap) * 100 : null;
          return { id: r.id, name: r.name, amphoe: r.amphoe,
            current: pct, warn: 80, crit: 100, bank: 100, river: "", _pct: pct };
        });
      } catch (e) { return []; }
    }
    if (typeof window.renderSummary === "function") {
      var _origRR = window.renderSummary;
      window.renderSummary = function () {
        _origRR.apply(this, arguments);
        renderVerdict(resStations(), null);
      };
    }
    setTimeout(function () { if (!zone.innerHTML) renderVerdict(resStations(), null); }, 1200);
  }
})();
