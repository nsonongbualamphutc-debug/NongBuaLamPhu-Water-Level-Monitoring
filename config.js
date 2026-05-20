// ═══════════════════════════════════════════════════════════
//  config.js — ค่าตั้งระบบ NBP Water Level Monitoring
//  จัดการ API key และ URL — ปกป้องไม่ให้หลุดในซอร์สโค้ดบน GitHub
//
//  ⚠️ ไฟล์นี้ต้อง:
//    1. อยู่ใน root ของ GitHub repo (ที่เดียวกับ index.html)
//    2. ใส่ Gemini API key ก่อนใช้งาน (ดูคู่มือด้านล่าง)
//    3. หาก repo เป็น public → key จะถูกเห็นได้ทั่ว
//       (ถ้าต้องการความเป็นส่วนตัวสูง ใช้ Cloudflare Pages หรือเก็บใน Apps Script แทน)
// ═══════════════════════════════════════════════════════════

window.APP_CONFIG = {
  // ───────── 1) Backend Apps Script URL ─────────
  // URL ของ Google Apps Script ที่ deploy เป็น Web App
  // ตัวอย่าง: "https://script.google.com/macros/s/AKfycby..../exec"
  API_URL: "https://script.google.com/macros/s/AKfycbyNbwf9250krK8ObZwR83RYgg139pf72T7QMh1paVb2BI3OD5n1U4nIPEDCdIZGkU7F/exec",

  // ───────── 2) Gemini AI API Key ─────────
  // รับฟรีที่ https://aistudio.google.com/app/apikey
  // ใช้สำหรับ AI สรุปรายงานน้ำในหน้า daily_briefing.html
  // ⚠️ แทน "YOUR_KEY_HERE" ด้วย key จริง (ขึ้นต้นด้วย "AIza...")
  GEMINI_API_KEY: "YOUR_KEY_HERE"
};
