/**
 * smart_summary.js
 * Rule-based AI Summary — ไม่ต้องใช้ API key ใดเลย
 * แทนที่ runGemini() ทั้งหมดใน daily_briefing.html
 *
 * วิธีใช้:
 *   1. ลบฟังก์ชัน runGemini() เดิมออก
 *   2. Paste โค้ดนี้แทน (ก่อน </script> ปิด)
 *   3. ใน init section เพิ่ม:
 *        setTimeout(runGemini, 8000);
 *        setInterval(runGemini, 15 * 60 * 1000);
 */

/* ══════════════════════════════════════════════════
 *  runGemini() — drop-in replacement ไม่ใช้ API key
 * ══════════════════════════════════════════════════ */
function runGemini() {
  const aiTxt  = document.getElementById('aiTxt');
  const aiNote = document.getElementById('aiKeyNote');
  const aiBtn  = document.getElementById('aiBtnRun');

  if (!aiTxt) return;

  // ซ่อน key warning
  if (aiNote) aiNote.style.display = 'none';

  // ── รวบรวมข้อมูลจากหน้าจอ ──
  const data = collectPageData();

  // ── วิเคราะห์ ──
  const summary = analyzeSituation(data);

  // ── แสดงผล typewriter ──
  if (aiBtn) aiBtn.disabled = true;
  typewriterAI(aiTxt, summary, 28, function() {
    if (aiBtn) {
      aiBtn.disabled = false;
      aiBtn.textContent = '🔄 วิเคราะห์ใหม่';
    }
  });
}

/* ──────────────────────────────────────────────────
 *  collectPageData — ดึงตัวเลขจาก DOM
 * ────────────────────────────────────────────────── */
function collectPageData() {
  function num(id) {
    const el = document.getElementById(id);
    const v  = parseFloat((el?.textContent || '').replace(/[^0-9.\-]/g, ''));
    return isNaN(v) ? null : v;
  }
  function txt(id) {
    return (document.getElementById(id)?.textContent || '').trim();
  }

  return {
    temp:     num('heroTemp'),
    feels:    num('heroFeels'),
    hi:       num('heroHi'),
    lo:       num('heroLo'),
    humidity: num('hHum'),
    wind:     num('hWind'),
    windDir:  txt('hWindDir'),
    rain:     num('hRain'),
    uv:       num('hUV'),
    desc:     txt('heroDesc'),

    // AQI
    aqi:      parseFloat((document.querySelector('.aqi-num')?.textContent || '').replace(/[^0-9.]/g, '')) || null,
    aqiQual:  (document.querySelector('.aqi-q')?.textContent || '').trim(),

    // สถานีน้ำ
    stTotal:  num('pws-stations-num'),
    stWarn:   num('pws-warn-num'),
    stCrit:   num('pws-crit-num'),

    // อ่างเก็บน้ำ
    resAvg:   parseFloat((txt('res-avg') || '0').replace('%', '')) || null,
    resWarn:  num('res-warn-num'),
    resLow:   num('res-low-num'),

    // เวลา
    hour: new Date().getHours(),
    month: new Date().getMonth() + 1,  // 1–12
  };
}

/* ──────────────────────────────────────────────────
 *  analyzeSituation — core rule engine
 * ────────────────────────────────────────────────── */
function analyzeSituation(d) {
  const parts = [];

  // ── 1. บริบทเวลา ──
  const timeCtx = d.hour < 6   ? 'ช่วงดึก'
                : d.hour < 12  ? 'ช่วงเช้า'
                : d.hour < 17  ? 'ช่วงบ่าย'
                : d.hour < 20  ? 'ช่วงเย็น'
                : 'ช่วงกลางคืน';

  const monthCtx = (d.month >= 5 && d.month <= 10) ? 'ฤดูฝน' : 'ฤดูแล้ง';

  // ── 2. สรุปอากาศ ──
  const weatherPart = buildWeatherSummary(d, timeCtx, monthCtx);
  if (weatherPart) parts.push(weatherPart);

  // ── 3. คุณภาพอากาศ ──
  const aqiPart = buildAQISummary(d);
  if (aqiPart) parts.push(aqiPart);

  // ── 4. สถานการณ์น้ำ ──
  const waterPart = buildWaterSummary(d, monthCtx);
  if (waterPart) parts.push(waterPart);

  // ── 5. อ่างเก็บน้ำ ──
  const resPart = buildReservoirSummary(d, monthCtx);
  if (resPart) parts.push(resPart);

  // ── 6. คำแนะนำ ──
  const advicePart = buildAdvice(d);
  if (advicePart) parts.push(advicePart);

  // fallback
  if (parts.length === 0) {
    return 'กำลังรอข้อมูลจากระบบ — กรุณารอสักครู่แล้วกด วิเคราะห์ใหม่';
  }

  return parts.join(' ');
}

/* ── weather summary ── */
function buildWeatherSummary(d, timeCtx, monthCtx) {
  if (d.temp === null) return null;

  let s = `${timeCtx}นี้ จ.หนองบัวลำภู`;

  // อุณหภูมิ
  if (d.temp >= 38)      s += ` อากาศร้อนจัด อุณหภูมิ ${d.temp}°C`;
  else if (d.temp >= 35) s += ` อากาศร้อนมาก ${d.temp}°C`;
  else if (d.temp >= 30) s += ` อากาศร้อน ${d.temp}°C`;
  else if (d.temp >= 25) s += ` อากาศอบอุ่น ${d.temp}°C`;
  else if (d.temp >= 20) s += ` อากาศเย็นสบาย ${d.temp}°C`;
  else                   s += ` อากาศเย็น ${d.temp}°C`;

  // ความชื้น
  if (d.humidity !== null) {
    if (d.humidity >= 85)      s += ` ความชื้นสูงมาก (${d.humidity}%)`;
    else if (d.humidity >= 70) s += ` ความชื้น ${d.humidity}%`;
    else if (d.humidity <= 40) s += ` อากาศแห้ง ความชื้น ${d.humidity}%`;
  }

  // ฝน
  if (d.rain !== null) {
    if (d.rain >= 90)      s += ` มีฝนตกหนักมาก สะสม ${d.rain} มม.`;
    else if (d.rain >= 35) s += ` มีฝนตกหนัก ${d.rain} มม.`;
    else if (d.rain >= 15) s += ` มีฝนปานกลาง ${d.rain} มม.`;
    else if (d.rain >= 1)  s += ` มีฝนเล็กน้อย ${d.rain} มม.`;
    else                   s += ` ไม่มีฝน`;
  }

  // ลม
  if (d.wind !== null && d.wind >= 30) {
    s += ` ลมแรง ${d.wind} กม./ชม.`;
  }

  s += '.';
  return s;
}

/* ── AQI summary ── */
function buildAQISummary(d) {
  if (d.aqi === null) return null;

  if (d.aqi > 150)       return `คุณภาพอากาศอยู่ในระดับ "มีผลกระทบต่อสุขภาพ" PM2.5 = ${d.aqi} µg/m³ ทุกคนควรหลีกเลี่ยงกิจกรรมกลางแจ้ง.`;
  if (d.aqi > 100)       return `PM2.5 อยู่ที่ ${d.aqi} µg/m³ ระดับ "ไม่ดีสำหรับกลุ่มเสี่ยง" ผู้สูงอายุและเด็กควรระวัง.`;
  if (d.aqi > 50)        return `PM2.5 ${d.aqi} µg/m³ อยู่ในระดับปานกลาง กลุ่มเสี่ยงควรสังเกตอาการ.`;
  return `คุณภาพอากาศดี PM2.5 = ${d.aqi} µg/m³ เหมาะกับกิจกรรมกลางแจ้ง.`;
}

/* ── water summary ── */
function buildWaterSummary(d, monthCtx) {
  const warn = d.stWarn || 0;
  const crit = d.stCrit || 0;
  const tot  = d.stTotal || 20;

  if (crit >= 3)       return `⚠️ สถานการณ์น้ำวิกฤติ! มี ${crit} สถานีที่ระดับน้ำล้นตลิ่งแล้ว ขอให้ประชาชนในพื้นที่ลุ่มต่ำเฝ้าระวังและเตรียมอพยพ.`;
  if (crit >= 1)       return `⚠️ มี ${crit} สถานีระดับน้ำถึงขั้นวิกฤติ และอีก ${warn} สถานีอยู่ในระดับเฝ้าระวัง ให้ติดตามสถานการณ์ใกล้ชิด.`;
  if (warn >= 5)       return `ระดับน้ำน่าเป็นห่วง มี ${warn} สถานีจาก ${tot} สถานีที่อยู่ในระดับเฝ้าระวัง ควรติดตามข้อมูลอย่างต่อเนื่อง.`;
  if (warn >= 2)       return `มี ${warn} สถานีที่ระดับน้ำอยู่ในช่วงเฝ้าระวัง สถานีที่เหลือปกติ.`;
  if (warn === 0 && crit === 0) {
    return monthCtx === 'ฤดูฝน'
      ? `ระดับน้ำทุกสถานีอยู่ในเกณฑ์ปกติ แม้จะอยู่ในช่วงฤดูฝน.`
      : `ระดับน้ำทุกสถานีอยู่ในเกณฑ์ปกติ.`;
  }
  return null;
}

/* ── reservoir summary ── */
function buildReservoirSummary(d, monthCtx) {
  if (d.resAvg === null) return null;

  const avg  = d.resAvg;
  const low  = d.resLow  || 0;
  const high = d.resWarn || 0;

  if (avg >= 90)      return `อ่างเก็บน้ำ 14 แห่งมีปริมาณน้ำสูงมาก เฉลี่ย ${avg}% ควรระวังน้ำล้นอ่าง.`;
  if (avg >= 70)      return `อ่างเก็บน้ำโดยรวมอยู่ในระดับดี เฉลี่ย ${avg}% เพียงพอต่อการใช้งาน.`;
  if (avg >= 40)      return `อ่างเก็บน้ำอยู่ในระดับปานกลาง ${avg}%${low > 0 ? ` มี ${low} แห่งที่น้ำน้อย` : ''}.`;
  if (avg >= 20)      return `อ่างเก็บน้ำมีปริมาณน้ำน้อย เฉลี่ย ${avg}% มี ${low} แห่งที่อยู่ในระดับต่ำกว่า 30%.`;
  return `อ่างเก็บน้ำส่วนใหญ่มีน้ำน้อยมาก เฉลี่ยเพียง ${avg}%${monthCtx === 'ฤดูแล้ง' ? ' อยู่ในช่วงฤดูแล้ง ควรใช้น้ำอย่างประหยัด' : ''}.`;
}

/* ── advice ── */
function buildAdvice(d) {
  const tips = [];
  const rain   = d.rain   || 0;
  const temp   = d.temp   || 30;
  const aqi    = d.aqi    || 0;
  const crit   = d.stCrit || 0;
  const wind   = d.wind   || 0;
  const uv     = d.uv     || 0;

  if (crit > 0)        tips.push('ประชาชนริมน้ำเตรียมย้ายสิ่งของขึ้นที่สูง');
  if (rain >= 35)      tips.push('ระวังน้ำท่วมฉับพลันและน้ำป่าไหลหลาก');
  if (temp >= 36 && rain < 5) tips.push('ดื่มน้ำให้เพียงพอ หลีกเลี่ยงแสงแดดช่วง 11.00–15.00 น.');
  if (uv >= 8)         tips.push('ดัชนี UV สูง ควรทาครีมกันแดดและสวมหมวก');
  if (aqi > 100)       tips.push('สวมหน้ากากอนามัยเมื่อออกนอกบ้าน');
  if (wind >= 40)      tips.push('ระวังต้นไม้ล้มและป้ายโฆษณาหล่น');

  if (tips.length === 0) return null;
  return `คำแนะนำ: ${tips.join(', ')}.`;
}

/* ──────────────────────────────────────────────────
 *  typewriterAI — แสดงผลแบบพิมพ์ทีละตัว
 * ────────────────────────────────────────────────── */
function typewriterAI(el, text, speedMs, onDone) {
  el.innerHTML = '';
  let i = 0;
  const cursor = document.createElement('span');
  cursor.className = 'ai-cursor';
  el.appendChild(cursor);

  const iv = setInterval(function() {
    if (i >= text.length) {
      clearInterval(iv);
      cursor.remove();
      if (onDone) onDone();
      return;
    }
    cursor.insertAdjacentText('beforebegin', text[i]);
    i++;
  }, speedMs);
}
