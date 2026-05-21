/**
 * ============================================================
 * Code.gs — Line Messaging API แจ้งเตือนน้ำวิกฤติอัตโนมัติ
 * ============================================================
 *
 * หมายเหตุสำคัญ:
 *   Line Notify ถูกปิดบริการ 31 มี.ค. 2568 (2025)
 *   ไฟล์นี้จึงใช้ "Line Messaging API" แทน — broadcast ฟรี
 *
 * ── ขั้นตอนติดตั้ง ──
 *  1. สร้าง Line Official Account ฟรีที่ https://manager.line.biz
 *  2. ไปที่ https://developers.line.biz → สร้าง Messaging API channel
 *  3. คัดลอก "Channel access token" (long-lived)
 *  4. Apps Script Editor → Project Settings → Script Properties:
 *       Key:   LINE_CHANNEL_TOKEN
 *       Value: <channel access token>
 *  5. paste ฟังก์ชันด้านล่างต่อท้าย Code.gs
 *  6. สร้าง Time-driven trigger:
 *       Edit → Triggers → Add Trigger
 *       Function: checkAndNotifyWater
 *       Event:    Time-driven → Every 30 minutes
 *  7. ทดสอบ: Run → testLineNotify
 *
 * ============================================================
 */

/* ════════════════════════════════════════════════════
 *  1. checkAndNotifyWater — เรียกโดย Time-trigger ทุก 30 นาที
 * ════════════════════════════════════════════════════ */
function checkAndNotifyWater() {
  const token = PropertiesService.getScriptProperties()
                                 .getProperty("LINE_CHANNEL_TOKEN");
  if (!token) {
    Logger.log("⚠️ ยังไม่ได้ตั้ง LINE_CHANNEL_TOKEN");
    return;
  }

  /* ดึงสรุปสถานการณ์น้ำ */
  let summary;
  try {
    summary = getSummary();
  } catch(e) {
    Logger.log("❌ getSummary error: " + e);
    return;
  }
  if (!summary || !summary.stations) return;

  /* แยกสถานีวิกฤติ / เฝ้าระวัง */
  const crit = summary.stations.filter(function(s) {
    return s.status === "วิกฤติ";
  });
  const warn = summary.stations.filter(function(s) {
    return s.status === "เฝ้าระวัง";
  });

  /* ไม่มีอะไรผิดปกติ → ไม่ส่ง */
  if (crit.length === 0 && warn.length === 0) {
    Logger.log("✅ สถานการณ์ปกติ — ไม่ต้องแจ้งเตือน");
    return;
  }

  /* anti-spam — ไม่แจ้งซ้ำสถานการณ์เดิมภายใน 3 ชม. */
  const cache    = CacheService.getScriptCache();
  const stateKey = "nbp_alert_state";
  const stateNow = crit.map(function(s){return s.station_id;}).sort().join(",") +
                   "|" + warn.length;
  const lastState = cache.get(stateKey);
  if (lastState === stateNow) {
    Logger.log("⏸ สถานการณ์เดิม — ข้ามการแจ้งซ้ำ");
    return;
  }

  /* สร้างข้อความ */
  const msg = buildAlertMessage_(crit, warn, summary);

  /* ส่ง broadcast */
  const sent = sendLineBroadcast_(token, msg);
  if (sent) {
    cache.put(stateKey, stateNow, 3 * 60 * 60); /* จำ 3 ชม. */
    Logger.log("✅ แจ้งเตือนแล้ว: วิกฤติ " + crit.length +
               " เฝ้าระวัง " + warn.length);
  }
}

/* ════════════════════════════════════════════════════
 *  2. buildAlertMessage_ — สร้างข้อความแจ้งเตือน
 * ════════════════════════════════════════════════════ */
function buildAlertMessage_(crit, warn, summary) {
  const now = Utilities.formatDate(new Date(), "Asia/Bangkok",
                                    "dd/MM/yyyy HH:mm");
  let lines = [];

  if (crit.length > 0) {
    lines.push("🔴 แจ้งเตือนน้ำวิกฤติ");
  } else {
    lines.push("🟡 แจ้งเตือนเฝ้าระวังน้ำ");
  }
  lines.push("จ.หนองบัวลำภู · " + now + " น.");
  lines.push("");

  /* รายการวิกฤติ */
  if (crit.length > 0) {
    lines.push("⚠️ สถานีวิกฤติ (น้ำล้นตลิ่ง) " + crit.length + " แห่ง:");
    crit.slice(0, 8).forEach(function(s) {
      const cur  = s.current_level || s.current || "?";
      const bank = s.bank_level    || s.bank    || "?";
      lines.push("• " + (s.name || s.station_id) +
                 " อ." + (s.amphoe || "-"));
      lines.push("  ระดับน้ำ " + cur + " / ตลิ่ง " + bank + " ม.");
    });
    if (crit.length > 8) {
      lines.push("  ...และอีก " + (crit.length - 8) + " สถานี");
    }
    lines.push("");
  }

  /* สรุปเฝ้าระวัง */
  if (warn.length > 0) {
    lines.push("🟡 สถานีเฝ้าระวัง (ใกล้ตลิ่ง): " + warn.length + " แห่ง");
    warn.slice(0, 5).forEach(function(s) {
      lines.push("• " + (s.name || s.station_id) +
                 " อ." + (s.amphoe || "-"));
    });
    if (warn.length > 5) {
      lines.push("  ...และอีก " + (warn.length - 5) + " สถานี");
    }
    lines.push("");
  }

  /* คำแนะนำ */
  if (crit.length >= 3) {
    lines.push("📢 ขอให้ประชาชนพื้นที่ลุ่มต่ำเฝ้าระวัง");
    lines.push("และเตรียมพร้อมขนย้ายสิ่งของขึ้นที่สูง");
  } else if (crit.length > 0) {
    lines.push("📢 ขอให้ติดตามสถานการณ์อย่างใกล้ชิด");
  }
  lines.push("");
  lines.push("ดูข้อมูลเพิ่มเติม:");
  lines.push("https://nsonongbualamphutc-debug.github.io/NongBuaLamPhu-Water-Level-Monitoring/");

  return lines.join("\n");
}

/* ════════════════════════════════════════════════════
 *  3. sendLineBroadcast_ — ส่งข้อความผ่าน Messaging API
 * ════════════════════════════════════════════════════ */
function sendLineBroadcast_(token, message) {
  const url = "https://api.line.me/v2/bot/message/broadcast";

  /* Line จำกัดข้อความ 5000 ตัวอักษร/ข้อความ */
  const text = message.length > 4900
    ? message.slice(0, 4900) + "\n..."
    : message;

  const payload = {
    messages: [
      { type: "text", text: text }
    ]
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method:            "post",
      contentType:       "application/json",
      headers:           { Authorization: "Bearer " + token },
      payload:           JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code = res.getResponseCode();
    if (code === 200) {
      return true;
    } else {
      Logger.log("❌ Line API error " + code + ": " + res.getContentText());
      return false;
    }
  } catch(err) {
    Logger.log("❌ Line broadcast error: " + err);
    return false;
  }
}

/* ════════════════════════════════════════════════════
 *  4. testLineNotify — ทดสอบจาก Editor (Run ฟังก์ชันนี้)
 * ════════════════════════════════════════════════════ */
function testLineNotify() {
  const token = PropertiesService.getScriptProperties()
                                 .getProperty("LINE_CHANNEL_TOKEN");
  if (!token) {
    Logger.log("❌ ยังไม่ได้ตั้ง LINE_CHANNEL_TOKEN ใน Script Properties");
    return;
  }

  const testMsg =
    "🧪 ทดสอบระบบแจ้งเตือน\n" +
    "ระบบสถานการณ์น้ำ จ.หนองบัวลำภู\n\n" +
    "หากได้รับข้อความนี้ แสดงว่าระบบแจ้งเตือนพร้อมใช้งาน ✅\n" +
    Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm");

  const ok = sendLineBroadcast_(token, testMsg);
  Logger.log(ok ? "✅ ส่งข้อความทดสอบสำเร็จ" : "❌ ส่งไม่สำเร็จ — ตรวจ token");
}

/* ════════════════════════════════════════════════════
 *  5. (ทางเลือก) ส่งแจ้งเตือนทันทีเมื่อบันทึกข้อมูลน้ำวิกฤติ
 *     เรียกใน saveWaterLevel() หลังบันทึกสำเร็จ
 * ════════════════════════════════════════════════════
 *
 * เพิ่มใน saveWaterLevel() ก่อน return:
 *
 *   // ตรวจว่าสถานีนี้วิกฤติไหม → แจ้งทันที
 *   try {
 *     notifyIfCritical_(stationId, levelValue, bankLevel, stationName, amphoe);
 *   } catch(e) { Logger.log("notify error: " + e); }
 *
 */
function notifyIfCritical_(stationId, level, bank, name, amphoe) {
  if (parseFloat(level) < parseFloat(bank)) return; /* ยังไม่ล้นตลิ่ง */

  const token = PropertiesService.getScriptProperties()
                                 .getProperty("LINE_CHANNEL_TOKEN");
  if (!token) return;

  /* anti-spam รายสถานี — 2 ชม. */
  const cache = CacheService.getScriptCache();
  const key   = "nbp_crit_" + stationId;
  if (cache.get(key)) return;

  const now = Utilities.formatDate(new Date(), "Asia/Bangkok",
                                    "dd/MM/yyyy HH:mm");
  const msg =
    "🔴 แจ้งเตือนด่วน — น้ำล้นตลิ่ง\n" +
    (name || stationId) + " อ." + (amphoe || "-") + "\n\n" +
    "ระดับน้ำ " + level + " ม. (ตลิ่ง " + bank + " ม.)\n" +
    "เกินตลิ่ง " + (parseFloat(level) - parseFloat(bank)).toFixed(2) + " ม.\n" +
    now + " น.\n\n" +
    "https://nsonongbualamphutc-debug.github.io/NongBuaLamPhu-Water-Level-Monitoring/";

  if (sendLineBroadcast_(token, msg)) {
    cache.put(key, "1", 2 * 60 * 60);
    Logger.log("✅ แจ้งเตือนน้ำวิกฤติ: " + stationId);
  }
}
