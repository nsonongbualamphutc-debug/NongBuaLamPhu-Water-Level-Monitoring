/**
 * ============================================================
 *  ระบบแดชบอร์ดสถานการณ์น้ำจังหวัดหนองบัวลำภู
 *  Google Apps Script Backend — v3 (+ PIN auth)
 * ============================================================
 *  PIN: ตั้งค่า APP_PIN ใน Script Properties
 * ============================================================ */

// ===== CONFIG =====
const SHEET_STATIONS    = "Stations";
const SHEET_WATER       = "WaterLevel";
const SHEET_RAIN        = "Rainfall";
const SHEET_RESERVOIR   = "Reservoir";
const SHEET_SETTINGS    = "Settings";
const SHEET_PINS        = "StationPins";      // PIN รายสถานี
const SHEET_AMPHOE_PINS = "AmphoePins";       // PIN รายอำเภอ (สำหรับฝน)
const SHEET_RESERVOIR_PINS = "ReservoirPins"; // PIN รายอ่างเก็บน้ำ

// ===== PIN =====
const PIN_PROPERTY_KEY   = "APP_PIN";           // legacy global PIN (เผื่อ reservoir.html, daily report)
const ADMIN_PIN_KEY      = "ADMIN_PIN";         // master PIN สำหรับ Admin
const PIN_REQUIRED       = true;                // false = ปิด PIN ระหว่าง dev
const WRITE_ACTIONS      = ["savewater","saverain","savereservoir","savedailyreport"];
const AMPHOES = ["เมืองหนองบัวลำภู","นากลาง","นาวัง","ศรีบุญเรือง","สุวรรณคูหา","โนนสัง"];

const RESERVOIR_HEADERS = [
  "reservoir_id","reservoir_name","amphoe","capacity",
  "current_volume","date","reporter","updated_at"
];

// ===== ENTRY POINTS =====

function doGet(e) {
  const params   = (e && e.parameter) ? e.parameter : {};
  const action   = (params.action   || "summary").toLowerCase();
  const callback = params.callback;
  let data;

  const WRITE_ACTIONS = ["savewater","saverain","savereservoir","savedailyreport"];

  try {
    // === WRITE ACTIONS via GET — PIN แยกตามสถานี/อำเภอ ===
    if (WRITE_ACTIONS.indexOf(action) !== -1) {
      const pin = String(params.pin || "").trim();
      let pinOk = false;
      let pinError = "PIN ไม่ถูกต้อง";

      if (PIN_REQUIRED) {
        // อนุญาต Admin master PIN override ทุก action
        if (action === "savewater") {
          const stid = String(params.station_id || "").toUpperCase();
          if (!stid) pinError = "ไม่ระบุรหัสสถานี (station_id)";
          else {
            const acc = verifyWaterAccess(stid, pin);
            if (acc.ok) { pinOk = true; if(acc.role==="station") touchStationPinLastUsed(stid); }
            else pinError = "PIN ไม่ถูกต้องสำหรับสถานี " + stid;
          }
        } else if (action === "saverain") {
          const amphoe = String(params.amphoe || "");
          if (!amphoe) pinError = "ไม่ระบุอำเภอ";
          else {
            const acc = verifyRainAccess(amphoe, pin);
            if (acc.ok) { pinOk = true; if(acc.role==="amphoe") touchAmphoePinLastUsed(amphoe); }
            else pinError = "PIN ไม่ถูกต้องสำหรับอำเภอ " + amphoe;
          }
        } else if (action === "savereservoir") {
          const rid = String(params.reservoir_id || "").toUpperCase();
          if (!rid) pinError = "ไม่ระบุรหัสอ่างเก็บน้ำ (reservoir_id)";
          else {
            const acc = verifyReservoirAccess(rid, pin);
            if (acc.ok) { pinOk = true; if(acc.role==="reservoir") touchReservoirPinLastUsed(rid); }
            else pinError = "PIN ไม่ถูกต้องสำหรับอ่าง " + rid;
          }
        } else if (verifyAdminPin(pin)) {
          // savedailyreport — Super admin
          pinOk = true;
        } else {
          // savedailyreport — fallback legacy APP_PIN
          const expected = getAppPin();
          if (expected && pin === expected) pinOk = true;
          else pinError = "ต้องใช้ PIN ผู้ดูแลระบบสำหรับการบันทึก " + action;
        }
      } else {
        pinOk = true;
      }

      if (!pinOk) return respond({ ok:false, error:pinError, code:"INVALID_PIN" }, callback);

      const payload = {};
      Object.keys(params).forEach(function(k){
        if (k === 'callback') return;
        payload[k] = params[k];
      });
      switch (action) {
        case "savewater":       data = saveWaterLevel(payload); break;
        case "saverain":        data = saveRainfall(payload); break;
        case "savereservoir":   data = saveReservoir(payload); break;
        case "savedailyreport": data = saveDailyReport(payload); break;
      }
      return respond(data, callback);
    }

    // === PUBLIC LOGIN — verify PIN, return entity info if ok ===
    if (action === "loginstation") {
      const stid = String(params.station_id || "").toUpperCase();
      const pin  = String(params.pin || "").trim();
      if (!stid) return respond({ ok:false, error:"ไม่ระบุรหัสสถานี" }, callback);
      const acc = verifyWaterAccess(stid, pin);
      if (!acc.ok) return respond({ ok:false, error:"PIN ไม่ถูกต้อง" }, callback);
      return respond({ ok:true, role:acc.role, station: getStationContext(stid) }, callback);
    }
    if (action === "loginamphoe") {
      const am  = String(params.amphoe || "");
      const pin = String(params.pin || "").trim();
      if (!am) return respond({ ok:false, error:"ไม่ระบุอำเภอ" }, callback);
      const acc = verifyRainAccess(am, pin);
      if (!acc.ok) return respond({ ok:false, error:"PIN ไม่ถูกต้อง" }, callback);
      return respond({ ok:true, role:acc.role, amphoe: getAmphoeContext(am) }, callback);
    }
    if (action === "loginreservoir") {
      const rid = String(params.reservoir_id || "").toUpperCase();
      const pin = String(params.pin || "").trim();
      if (!rid) return respond({ ok:false, error:"ไม่ระบุรหัสอ่างเก็บน้ำ" }, callback);
      const acc = verifyReservoirAccess(rid, pin);
      if (!acc.ok) return respond({ ok:false, error:"PIN ไม่ถูกต้อง" }, callback);
      return respond({ ok:true, role:acc.role, reservoir: getReservoirContext(rid) }, callback);
    }

    // ===== Sub-Admin login: ตรวจว่า PIN เป็น sub-admin หรือ super-admin =====
    // ใช้ใน input.html เพื่อให้ user "เข้าระบบ admin" แล้วเลือกสถานีอะไรก็ได้
    if (action === "loginwateradmin")     return respond(checkSubAdminRole("water", String(params.pin||"")), callback);
    if (action === "loginrainadmin")      return respond(checkSubAdminRole("rain", String(params.pin||"")), callback);
    if (action === "loginreservoiradmin") return respond(checkSubAdminRole("reservoir", String(params.pin||"")), callback);
    if (action === "adminlogin") {
      const pin = String(params.pin || "").trim();
      if (!verifyAdminPin(pin)) return respond({ ok:false, error:"PIN ผู้ดูแลระบบไม่ถูกต้อง" }, callback);
      return respond({ ok:true, role:"admin" }, callback);
    }

    // === ADMIN ACTIONS (need admin pin in `adminpin` param) ===
    const adminActions = ["liststationpins","setstationpin","initstationpins","listamphoepins","setamphoepin","initamphoepins","listreservoirpins","setreservoirpin","initreservoirpins","stationcontext","amphoecontext","reservoircontext"];
    if (adminActions.indexOf(action) !== -1) {
      const adminpin = String(params.adminpin || params.pin || "").trim();
      const isAdmin = verifyAdminPin(adminpin);
      if (action === "stationcontext") {
        const stid = String(params.station_id || "").toUpperCase();
        if (!stid) return respond({ ok:false, error:"ไม่ระบุรหัสสถานี" }, callback);
        if (!isAdmin && !verifyWaterAdminPin(adminpin) && !verifyStationPin(stid, adminpin)) return respond({ ok:false, error:"PIN ไม่ถูกต้อง" }, callback);
        return respond({ ok:true, station: getStationContext(stid) }, callback);
      }
      if (action === "amphoecontext") {
        const am = String(params.amphoe || "");
        if (!am) return respond({ ok:false, error:"ไม่ระบุอำเภอ" }, callback);
        if (!isAdmin && !verifyRainAdminPin(adminpin) && !verifyAmphoePin(am, adminpin)) return respond({ ok:false, error:"PIN ไม่ถูกต้อง" }, callback);
        return respond({ ok:true, amphoe: getAmphoeContext(am) }, callback);
      }
      if (action === "reservoircontext") {
        const rid = String(params.reservoir_id || "").toUpperCase();
        if (!rid) return respond({ ok:false, error:"ไม่ระบุรหัสอ่างเก็บน้ำ" }, callback);
        if (!isAdmin && !verifyReservoirAdminPin(adminpin) && !verifyReservoirPin(rid, adminpin)) return respond({ ok:false, error:"PIN ไม่ถูกต้อง" }, callback);
        return respond({ ok:true, reservoir: getReservoirContext(rid) }, callback);
      }
      if (!isAdmin) return respond({ ok:false, error:"ต้องใช้ PIN ผู้ดูแลระบบ" }, callback);
      switch (action) {
        case "liststationpins":   data = listStationPins();  break;
        case "setstationpin":     data = setStationPin_(params.station_id, params.new_pin, params.recorder_name); break;
        case "initstationpins":   data = initStationPins();  break;
        case "listamphoepins":    data = listAmphoePins();   break;
        case "setamphoepin":      data = setAmphoePin_(params.amphoe, params.new_pin, params.recorder_name); break;
        case "initamphoepins":    data = initAmphoePins();   break;
        case "listreservoirpins": data = listReservoirPins(); break;
        case "setreservoirpin":   data = setReservoirPin_(params.reservoir_id, params.new_pin, params.recorder_name); break;
        case "initreservoirpins": data = initReservoirPins(); break;
      }
      return respond(data, callback);
    }

    // === READ ACTIONS (เดิม) ===
    switch (action) {
      case "summary":     data = getSummary(); break;
      case "paneang":     data = getRiverDashboard("paneang"); break;
      case "mong":        data = getRiverDashboard("mong"); break;
      case "mo":          data = getRiverDashboard("mo"); break;
      case "phuay":       data = getRiverDashboard("phuay"); break;
      case "stations":    data = getStations(params.river); break;
      case "stationlist": data = getStationListPublic(); break;  // ใช้ใน login screen
      case "water":       data = getWaterLevels(params.station_id, parseInt(params.days||"7")); break;
      case "rain":        data = getRainfall(parseInt(params.days||"7")); break;
      case "reservoir":   data = getReservoirs(); break;
      case "history":     data = getHistory(parseInt(params.limit||"20")); break;
      case "dailyreport": data = getDailyReport(params.date); break;
      case "ping":        data = { ok:true, time:new Date().toISOString() }; break;
      default:            data = { error:"unknown action: "+action };
    }
  } catch(err) { data = { ok:false, error:err.toString() }; }
  return respond(data, callback);
}

function doPost(e) {
  let payload = {};
  try {
    if (e && e.postData && e.postData.contents) payload = JSON.parse(e.postData.contents);
    else if (e && e.parameter) payload = e.parameter;
  } catch(err) { return respond({ ok:false, error:"Invalid JSON: "+err.toString() }); }

  const WRITE_ACTIONS = ["savewater","saverain","savereservoir","savedailyreport"];
  const action = (payload.action || "").toLowerCase();

  // PIN CHECK สำหรับ action เขียนข้อมูล
  if (PIN_REQUIRED && WRITE_ACTIONS.indexOf(action) !== -1) {
    const expectedPin = getAppPin();
    if (!expectedPin) {
      return respond({ ok:false, error:"ยังไม่ได้ตั้งค่า APP_PIN ใน Script Properties", code:"PIN_NOT_CONFIGURED" });
    }
    const pin = String(payload.pin || "").trim();
    if (pin !== expectedPin) {
      return respond({ ok:false, error:"PIN ไม่ถูกต้อง", code:"INVALID_PIN" });
    }
  }

  let result;
  try {
    switch (action) {
      case "savewater":       result = saveWaterLevel(payload); break;
      case "saverain":        result = saveRainfall(payload); break;
      case "savereservoir":   result = saveReservoir(payload); break;
      case "savedailyreport": result = saveDailyReport(payload); break;
      case "summary":         result = getSummary(); break;
      case "reservoir":       result = getReservoirs(); break;
      default:                result = { ok:false, error:"unknown action: "+action };
    }
  } catch(err) { result = { ok:false, error:err.toString() }; }
  return respond(result);
}

// ===== READ =====

function getStations(river) {
  const rows = sheetToObjects(ss().getSheetByName(SHEET_STATIONS));
  if (river) return rows.filter(r => String(r.river||"").indexOf(river)!==-1);
  return rows;
}

function getSummary() {
  const stations = getStations(), latest = getLatestWaterByStation();
  let normal=0,warn=0,crit=0;
  const merged = stations.map(st => {
    const w=latest[st.station_id]||{}, level=parseFloat(w.level);
    let status="ปกติ";
    if(!isNaN(level)){
      if(parseFloat(st.bank_level)&&level>=parseFloat(st.bank_level)) status="วิกฤติ";
      else if(parseFloat(st.warn_level)&&level>=parseFloat(st.warn_level)) status="เฝ้าระวัง";
    }
    if(status==="วิกฤติ") crit++; else if(status==="เฝ้าระวัง") warn++; else normal++;
    return Object.assign({},st,{current_level:isNaN(level)?null:level,flow:w.flow||null,status,last_update:w.date?(w.date+" "+(w.time||"")):null});
  });
  const rain=getRainfall(1);
  let avgRain=0;
  if(rain.length>0) avgRain=rain.reduce((a,r)=>a+(parseFloat(r.rain_24hr)||0),0)/rain.length;
  return {total:stations.length,normal,warn,crit,avg_rain_24hr:avgRain,stations:merged,updated:new Date().toISOString()};
}

function getRiverDashboard(riverKey) {
  const key = String(riverKey || "").toLowerCase();
  const stations = getStations().filter(st => {
    const id = String(st.station_id || "").toUpperCase();
    const river = String(st.river || "").toLowerCase();
    if (key === "paneang") return id.indexOf("PN") === 0 || river.indexOf("paneang") !== -1 || river.indexOf("พะเนียง") !== -1;
    if (key === "mong")    return id.indexOf("MG") === 0 || river.indexOf("mong") !== -1 || river.indexOf("โมง") !== -1;
    if (key === "mo")      return id.indexOf("MO") === 0 || river.indexOf("mo") === 0 || river.indexOf("ลำน้ำมอ") !== -1;
    if (key === "phuay")   return id.indexOf("PY") === 0 || river.indexOf("phuay") !== -1 || river.indexOf("พวย") !== -1;
    return true;
  });
  const latest = getLatestWaterByStation();
  let normal = 0, warn = 0, crit = 0;
  const merged = stations.map(st => {
    const w = latest[st.station_id] || {};
    const level = parseFloat(w.level);
    const bank = parseFloat(st.bank_level);
    const warnLevel = parseFloat(st.warn_level);
    let status = "ปกติ";
    if (!isNaN(level)) {
      if (!isNaN(bank) && level >= bank) status = "วิกฤติ";
      else if (!isNaN(warnLevel) && level >= warnLevel) status = "เฝ้าระวัง";
    }
    if (status === "วิกฤติ") crit++;
    else if (status === "เฝ้าระวัง") warn++;
    else normal++;
    return Object.assign({}, st, {
      id: st.station_id,
      current: isNaN(level) ? null : level,
      current_level: isNaN(level) ? null : level,
      flow: w.flow || null,
      status: status,
      last_update: w.date ? (w.date + " " + (w.time || "")) : null
    });
  });
  return {
    river: key,
    total: stations.length,
    normal: normal,
    warn: warn,
    crit: crit,
    stations: merged,
    updated: new Date().toISOString()
  };
}

function getWaterLevels(stationId,days) {
  const rows=sheetToObjects(ss().getSheetByName(SHEET_WATER));
  const cutoff=new Date(); cutoff.setDate(cutoff.getDate()-(days||7));
  return rows.filter(r=>{if(stationId&&r.station_id!==stationId)return false;const d=parseDate(r.date);return d&&d>=cutoff;}).sort((a,b)=>parseDate(a.date)-parseDate(b.date));
}

function getLatestWaterByStation() {
  const rows=sheetToObjects(ss().getSheetByName(SHEET_WATER)), latest={};
  rows.forEach(r=>{const sid=r.station_id;if(!sid)return;const d=parseDate(r.date);if(!latest[sid]||parseDate(latest[sid].date)<d)latest[sid]=r;});
  return latest;
}

function getRainfall(days) {
  const rows=sheetToObjects(ss().getSheetByName(SHEET_RAIN));
  const cutoff=new Date(); cutoff.setDate(cutoff.getDate()-(days||7));
  return rows.filter(r=>{const d=parseDate(r.date);return d&&d>=cutoff;}).sort((a,b)=>parseDate(b.date)-parseDate(a.date));
}

function getReservoirs() {
  const sheet=getOrCreateReservoirSheet(), rows=sheetToObjects(sheet), latest={};
  rows.forEach(r=>{const id=r.reservoir_id;if(!id)return;const d=parseDate(r.date);if(!latest[id]||parseDate(latest[id].date)<d)latest[id]=r;});
  return Object.values(latest);
}

function getHistory(limit) {
  const water=sheetToObjects(ss().getSheetByName(SHEET_WATER)).map(r=>Object.assign({type:"water"},r));
  const rain=sheetToObjects(ss().getSheetByName(SHEET_RAIN)).map(r=>Object.assign({type:"rain"},r));
  return water.concat(rain).sort((a,b)=>(parseDate(b.date)||new Date(0))-(parseDate(a.date)||new Date(0))).slice(0,limit||20);
}

function getDailyReport(targetDate) {
  const sheet=ss().getSheetByName("DailyReport");
  if(!sheet) return null;
  const rows=sheetToObjects(sheet);
  if(!rows.length) return null;
  if(targetDate) return rows.find(r=>String(r.date).slice(0,10)===targetDate)||null;
  rows.sort((a,b)=>parseDate(b.date)-parseDate(a.date));
  return rows[0];
}

// ===== WRITE =====

function saveWaterLevel(p) {
  const sheet = ss().getSheetByName(SHEET_WATER);
  const headers = getHeaders(sheet);
  // ใช้ != null แทน || "" เพื่อยอมรับค่า 0 / "0"
  const row = headers.map(h => {
    const v = p[h];
    return (v === undefined || v === null) ? "" : v;
  });
  sheet.appendRow(row);
  return {
    ok: true,
    message: "บันทึกข้อมูลระดับน้ำ " + (p.station_id||"") + " เรียบร้อย",
    station_id: p.station_id,
    recorded: { date:p.date, time:p.time, level:p.level, recorder:p.recorder }
  };
}

function saveRainfall(p) {
  const sheet = ss().getSheetByName(SHEET_RAIN);
  const headers = getHeaders(sheet);
  const row = headers.map(h => {
    const v = p[h];
    return (v === undefined || v === null) ? "" : v;
  });
  sheet.appendRow(row);
  return {
    ok: true,
    message: "บันทึกข้อมูลฝน " + (p.amphoe||"") + " เรียบร้อย",
    amphoe: p.amphoe,
    recorded: { date:p.date, rain_24hr:p.rain_24hr, recorder:p.recorder }
  };
}

function saveReservoir(p) {
  const sheet=getOrCreateReservoirSheet(), payload=normalizeReservoirPayload(p);
  if(!payload.reservoir_id&&!payload.reservoir_name) return {ok:false,error:"missing reservoir_id or reservoir_name"};
  const data=sheet.getDataRange().getValues(), headers=data[0];
  const idIdx=headers.indexOf("reservoir_id"), nmIdx=headerIndex(headers,["reservoir_name","name"]);
  const curIdx=headerIndex(headers,["current_volume","current"]), dateIdx=headers.indexOf("date"), updIdx=headers.indexOf("updated_at");
  const dateStr=String(payload.date||"").slice(0,10);

  for(let i=1;i<data.length;i++){
    const sameId  = idIdx>=0&&payload.reservoir_id  &&String(data[i][idIdx]).trim()===String(payload.reservoir_id).trim();
    const sameName= nmIdx>=0&&payload.reservoir_name&&String(data[i][nmIdx]).trim()===String(payload.reservoir_name).trim();
    const sameDate= dateIdx>=0&&String(data[i][dateIdx]).slice(0,10)===dateStr;
    if((sameId||sameName)&&sameDate){
      if(curIdx>=0) sheet.getRange(i+1,curIdx+1).setValue(payload.current_volume);
      if(updIdx>=0) sheet.getRange(i+1,updIdx+1).setValue(new Date());
      headers.forEach((h,col)=>{if(["current_volume","current","updated_at","date"].indexOf(h)!==-1)return;if(payload[h]!==undefined&&payload[h]!==null&&payload[h]!=="")sheet.getRange(i+1,col+1).setValue(payload[h]);});
      return {ok:true,message:"อัปเดต "+(payload.reservoir_name||payload.reservoir_id)+" วันที่ "+dateStr};
    }
  }
  const row=headers.map(h=>{if(h==="updated_at")return new Date();return payload[h]!==undefined?payload[h]:"";});
  sheet.appendRow(row);
  return {ok:true,message:"บันทึก "+(payload.reservoir_name||payload.reservoir_id)+" วันที่ "+dateStr};
}

function saveDailyReport(p) {
  const ss_obj=ss(); let sheet=ss_obj.getSheetByName("DailyReport");
  if(!sheet){
    sheet=ss_obj.insertSheet("DailyReport");
    sheet.appendRow(["date","reporter","dam_level","dam_use","dam_pct","dam_in","dam_out","dam_total","tmd_temp","tmd_cloud","tmd_rain_yest","tmd_pressure","tmd_humidity","tmd_wind","tmd_temp_min","tmd_visibility","tmd_rain_year","aqi_pm25","aqi_value","aqi_days_over","disaster_status","disaster_amphoe","disaster_note","saved_at"]);
  }
  const headers=getHeaders(sheet), dateStr=String(p.date||"").slice(0,10), dateIdx=headers.indexOf("date"), data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][dateIdx]).slice(0,10)===dateStr){
      const row=headers.map(h=>h==="saved_at"?new Date():(p[h]||data[i][headers.indexOf(h)]));
      sheet.getRange(i+1,1,1,headers.length).setValues([row]);
      return {ok:true,message:"อัปเดตรายงานวันที่ "+dateStr};
    }
  }
  sheet.appendRow(headers.map(h=>h==="saved_at"?new Date():(p[h]||"")));
  return {ok:true,message:"บันทึกรายงานวันที่ "+dateStr};
}

// ===== PIN MANAGEMENT =====
/** ระบบ PIN แยกตามสถานี/อำเภอ + Admin master PIN */

function ensureStationPinsSheet_() {
  const ss_obj = ss();
  let sh = ss_obj.getSheetByName(SHEET_PINS);
  if (!sh) {
    sh = ss_obj.insertSheet(SHEET_PINS);
    sh.appendRow(["station_id","pin","recorder_name","phone","last_changed","last_used","note"]);
    sh.getRange("A1:G1").setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#fff");
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(["station_id","pin","recorder_name","phone","last_changed","last_used","note"]);
  }
  return sh;
}

function ensureAmphoePinsSheet_() {
  const ss_obj = ss();
  let sh = ss_obj.getSheetByName(SHEET_AMPHOE_PINS);
  if (!sh) {
    sh = ss_obj.insertSheet(SHEET_AMPHOE_PINS);
    sh.appendRow(["amphoe","pin","recorder_name","phone","last_changed","last_used","note"]);
    sh.getRange("A1:G1").setFontWeight("bold").setBackground("#065f46").setFontColor("#fff");
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(["amphoe","pin","recorder_name","phone","last_changed","last_used","note"]);
  }
  return sh;
}

function getAdminPin_() {
  return PropertiesService.getScriptProperties().getProperty(ADMIN_PIN_KEY) || "";
}

function verifyAdminPin(pin) {
  const expected = getAdminPin_();
  if (!expected) return false;
  return String(pin || "").trim() === String(expected).trim();
}

// ===== SUB-ADMIN PINs (กรอกข้ามทุกจุดในประเภทของตน — ไม่ใช่ super admin) =====
const WATER_ADMIN_PIN_KEY     = "WATER_ADMIN_PIN";     // กรอกระดับน้ำได้ทุกสถานี
const RAIN_ADMIN_PIN_KEY      = "RAIN_ADMIN_PIN";      // กรอกฝนได้ทุกอำเภอ
const RESERVOIR_ADMIN_PIN_KEY = "RESERVOIR_ADMIN_PIN"; // กรอกอ่างได้ทุกอ่าง

function verifyWaterAdminPin(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty(WATER_ADMIN_PIN_KEY) || "";
  if (!expected) return false;
  return String(pin || "").trim() === String(expected).trim();
}
function verifyRainAdminPin(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty(RAIN_ADMIN_PIN_KEY) || "";
  if (!expected) return false;
  return String(pin || "").trim() === String(expected).trim();
}
function verifyReservoirAdminPin(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty(RESERVOIR_ADMIN_PIN_KEY) || "";
  if (!expected) return false;
  return String(pin || "").trim() === String(expected).trim();
}

/** ตรวจ PIN ทุกระดับสำหรับ savewater (station, water-admin, super-admin) */
function verifyWaterAccess(stationId, pin) {
  if (verifyAdminPin(pin)) return { ok:true, role:"super_admin" };
  if (verifyWaterAdminPin(pin)) return { ok:true, role:"water_admin" };
  if (verifyStationPin(stationId, pin)) return { ok:true, role:"station" };
  return { ok:false };
}
/** ตรวจ PIN สำหรับ saverain */
function verifyRainAccess(amphoe, pin) {
  if (verifyAdminPin(pin)) return { ok:true, role:"super_admin" };
  if (verifyRainAdminPin(pin)) return { ok:true, role:"rain_admin" };
  if (verifyAmphoePin(amphoe, pin)) return { ok:true, role:"amphoe" };
  return { ok:false };
}
/** ตรวจ PIN สำหรับ savereservoir */
function verifyReservoirAccess(reservoirId, pin) {
  if (verifyAdminPin(pin)) return { ok:true, role:"super_admin" };
  if (verifyReservoirAdminPin(pin)) return { ok:true, role:"reservoir_admin" };
  if (verifyReservoirPin(reservoirId, pin)) return { ok:true, role:"reservoir" };
  return { ok:false };
}

/** API: ตรวจ PIN ว่าเป็น sub-admin หรือไม่ — สำหรับ frontend เรียกตอน login */
function checkSubAdminRole(scope, pin) {
  // scope: "water" | "rain" | "reservoir"
  if (verifyAdminPin(pin)) return { ok:true, role:"super_admin", scope:scope };
  if (scope === "water" && verifyWaterAdminPin(pin))         return { ok:true, role:"water_admin", scope:"water" };
  if (scope === "rain"  && verifyRainAdminPin(pin))          return { ok:true, role:"rain_admin",  scope:"rain" };
  if (scope === "reservoir" && verifyReservoirAdminPin(pin)) return { ok:true, role:"reservoir_admin", scope:"reservoir" };
  return { ok:false };
}

function getStationPinRow_(stationId) {
  const sh = ensureStationPinsSheet_();
  if (sh.getLastRow() < 2) return null;
  const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  const target = String(stationId || "").toUpperCase().trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").toUpperCase().trim() === target) {
      return { row: i+2, station_id: target, pin: String(data[i][1]||""), recorder_name: data[i][2]||"", phone: data[i][3]||"", last_changed: data[i][4]||"", last_used: data[i][5]||"", note: data[i][6]||"" };
    }
  }
  return null;
}

function verifyStationPin(stationId, pin) {
  const r = getStationPinRow_(stationId);
  if (!r || !r.pin) return false;
  return String(pin||"").trim() === String(r.pin).trim();
}

function touchStationPinLastUsed(stationId) {
  try {
    const r = getStationPinRow_(stationId);
    if (r) ensureStationPinsSheet_().getRange(r.row, 6).setValue(new Date());
  } catch(e) {}
}

function getAmphoePinRow_(amphoe) {
  const sh = ensureAmphoePinsSheet_();
  if (sh.getLastRow() < 2) return null;
  const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  const target = String(amphoe || "").trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === target) {
      return { row: i+2, amphoe: target, pin: String(data[i][1]||""), recorder_name: data[i][2]||"", phone: data[i][3]||"", last_changed: data[i][4]||"", last_used: data[i][5]||"", note: data[i][6]||"" };
    }
  }
  return null;
}

function verifyAmphoePin(amphoe, pin) {
  const r = getAmphoePinRow_(amphoe);
  if (!r || !r.pin) return false;
  return String(pin||"").trim() === String(r.pin).trim();
}

function touchAmphoePinLastUsed(amphoe) {
  try {
    const r = getAmphoePinRow_(amphoe);
    if (r) ensureAmphoePinsSheet_().getRange(r.row, 6).setValue(new Date());
  } catch(e) {}
}

/** Master list ของสถานีทั้งจังหวัด — ใช้สำหรับ initStationPins
 *  จะทำงานได้แม้ Sheet Stations ยังว่าง (ไม่ต้องรัน runSetup ก่อน) */
const STATION_MASTER_LIST = [
  {station_id:"PN01",name:"วังปลาป้อม",       river:"ลำน้ำพะเนียง",amphoe:"นาวัง"},
  {station_id:"PN02",name:"โคกกระทอ",         river:"ลำน้ำพะเนียง",amphoe:"นาวัง"},
  {station_id:"PN03",name:"วังสามหาบ",        river:"ลำน้ำพะเนียง",amphoe:"นาวัง"},
  {station_id:"PN04",name:"บ้านหนองด่าน",     river:"ลำน้ำพะเนียง",amphoe:"นากลาง"},
  {station_id:"PN05",name:"บ้านฝั่งแดง",      river:"ลำน้ำพะเนียง",amphoe:"นากลาง"},
  {station_id:"PN06",name:"ปตร.หนองหว้าใหญ่", river:"ลำน้ำพะเนียง",amphoe:"เมืองหนองบัวลำภู"},
  {station_id:"PN07",name:"วังหมื่น",          river:"ลำน้ำพะเนียง",amphoe:"เมืองหนองบัวลำภู"},
  {station_id:"PN08",name:"ปตร.ปู่หลอด",       river:"ลำน้ำพะเนียง",amphoe:"เมืองหนองบัวลำภู"},
  {station_id:"PN09",name:"บ้านข้องโป้",       river:"ลำน้ำพะเนียง",amphoe:"เมืองหนองบัวลำภู"},
  {station_id:"PN10",name:"ปตร.หัวนา",         river:"ลำน้ำพะเนียง",amphoe:"เมืองหนองบัวลำภู"},
  {station_id:"MG01",name:"คลองบุญทัน",        river:"ลำน้ำโมง",    amphoe:"สุวรรณคูหา"},
  {station_id:"MG02",name:"บ้านโคก",           river:"ลำน้ำโมง",    amphoe:"สุวรรณคูหา"},
  {station_id:"MG03",name:"บ้านนาตาแหลว",     river:"ลำน้ำโมง",    amphoe:"สุวรรณคูหา"},
  {station_id:"MG04",name:"บ้านกุดผึ้ง",        river:"ลำน้ำโมง",    amphoe:"สุวรรณคูหา"},
  {station_id:"MO01",name:"อ่างเก็บน้ำมอ",      river:"ลำน้ำมอ",     amphoe:"ศรีบุญเรือง"},
  {station_id:"MO02",name:"บ้านวังคูณ",         river:"ลำน้ำมอ",     amphoe:"ศรีบุญเรือง"},
  {station_id:"MO03",name:"บ้านโนนสูงเปลือย",   river:"ลำน้ำมอ",     amphoe:"ศรีบุญเรือง"},
  {station_id:"PY01",name:"บ้านวังโปร่ง",       river:"ลำน้ำพวย",    amphoe:"ศรีบุญเรือง"},
  {station_id:"PY02",name:"บ้านทุ่งโพธิ์",      river:"ลำน้ำพวย",    amphoe:"ศรีบุญเรือง"},
  {station_id:"PY03",name:"บ้านโคกล่าม",       river:"ลำน้ำพวย",    amphoe:"ศรีบุญเรือง"}
];

/** รายการสถานีที่ใช้สำหรับจัดการ PIN — รวมข้อมูลจาก Sheet Stations (ถ้ามี) กับ master list
 *  เพื่อให้ initStationPins ทำงานได้แม้ Sheet ยังว่างเปล่า */
function getStationsForPinAdmin_() {
  const fromSheet = getStations(); // อ่านจาก Sheet Stations
  const byId = {};
  // master list มาก่อน (เป็น default)
  STATION_MASTER_LIST.forEach(s => { byId[s.station_id] = Object.assign({}, s); });
  // sheet override (กรณีเจ้าหน้าที่แก้ใน Sheet เอง)
  fromSheet.forEach(s => {
    const id = String(s.station_id||"").toUpperCase();
    if (id) byId[id] = Object.assign(byId[id]||{}, s, {station_id:id});
  });
  return Object.keys(byId).sort().map(k => byId[k]);
}

function listStationPins() {
  ensureStationPinsSheet_();
  const stations = getStationsForPinAdmin_();
  const out = [];
  stations.forEach(s => {
    const row = getStationPinRow_(s.station_id);
    out.push({
      station_id: s.station_id,
      name: s.name,
      river: s.river,
      amphoe: s.amphoe,
      pin: row ? row.pin : "",
      recorder_name: row ? row.recorder_name : "",
      phone: row ? row.phone : "",
      last_changed: row ? (row.last_changed ? Utilities.formatDate(new Date(row.last_changed), "Asia/Bangkok", "yyyy-MM-dd HH:mm") : "") : "",
      last_used: row ? (row.last_used ? Utilities.formatDate(new Date(row.last_used), "Asia/Bangkok", "yyyy-MM-dd HH:mm") : "") : "",
      note: row ? row.note : "",
      has_pin: !!(row && row.pin)
    });
  });
  return out;
}

function setStationPin_(stationId, newPin, recorderName) {
  if (!stationId) return { ok:false, error:"ไม่ระบุรหัสสถานี" };
  if (!newPin || String(newPin).trim().length < 3) return { ok:false, error:"PIN ต้องมีอย่างน้อย 3 ตัวอักษร" };
  const sh = ensureStationPinsSheet_();
  const row = getStationPinRow_(stationId);
  const now = new Date();
  if (row) {
    sh.getRange(row.row, 2).setValue(String(newPin).trim());
    if (recorderName !== undefined && recorderName !== null) sh.getRange(row.row, 3).setValue(recorderName);
    sh.getRange(row.row, 5).setValue(now);
  } else {
    sh.appendRow([String(stationId).toUpperCase(), String(newPin).trim(), recorderName||"", "", now, "", ""]);
  }
  return { ok:true, message:"ตั้ง PIN สำหรับ " + stationId + " เรียบร้อย", station_id: stationId };
}

function initStationPins() {
  ensureStationPinsSheet_();
  const stations = getStationsForPinAdmin_();
  if (!stations.length) {
    return { ok:false, error:"ไม่พบรายการสถานี (master list ว่าง — เป็นไปไม่ได้ในสภาวะปกติ)" };
  }
  const added = [];
  stations.forEach(s => {
    const existing = getStationPinRow_(s.station_id);
    if (!existing || !existing.pin) {
      const pin = String(Math.floor(1000 + Math.random()*9000));
      setStationPin_(s.station_id, pin, "");
      added.push({station_id: s.station_id, name: s.name, pin: pin});
    }
  });
  return { ok:true, message:"สร้าง PIN ใหม่ " + added.length + " สถานี (จากทั้งหมด " + stations.length + " สถานี — ที่เหลือมี PIN อยู่แล้ว)", added: added };
}

function listAmphoePins() {
  ensureAmphoePinsSheet_();
  return AMPHOES.map(am => {
    const row = getAmphoePinRow_(am);
    return {
      amphoe: am,
      pin: row ? row.pin : "",
      recorder_name: row ? row.recorder_name : "",
      phone: row ? row.phone : "",
      last_changed: row ? (row.last_changed ? Utilities.formatDate(new Date(row.last_changed), "Asia/Bangkok", "yyyy-MM-dd HH:mm") : "") : "",
      last_used: row ? (row.last_used ? Utilities.formatDate(new Date(row.last_used), "Asia/Bangkok", "yyyy-MM-dd HH:mm") : "") : "",
      has_pin: !!(row && row.pin)
    };
  });
}

function setAmphoePin_(amphoe, newPin, recorderName) {
  if (!amphoe) return { ok:false, error:"ไม่ระบุอำเภอ" };
  if (!newPin || String(newPin).trim().length < 3) return { ok:false, error:"PIN ต้องมีอย่างน้อย 3 ตัวอักษร" };
  const sh = ensureAmphoePinsSheet_();
  const row = getAmphoePinRow_(amphoe);
  const now = new Date();
  if (row) {
    sh.getRange(row.row, 2).setValue(String(newPin).trim());
    if (recorderName !== undefined && recorderName !== null) sh.getRange(row.row, 3).setValue(recorderName);
    sh.getRange(row.row, 5).setValue(now);
  } else {
    sh.appendRow([amphoe, String(newPin).trim(), recorderName||"", "", now, "", ""]);
  }
  return { ok:true, message:"ตั้ง PIN สำหรับอำเภอ " + amphoe + " เรียบร้อย", amphoe: amphoe };
}

function initAmphoePins() {
  ensureAmphoePinsSheet_();
  const added = [];
  AMPHOES.forEach(am => {
    const existing = getAmphoePinRow_(am);
    if (!existing || !existing.pin) {
      const pin = String(Math.floor(1000 + Math.random()*9000));
      setAmphoePin_(am, pin, "");
      added.push({amphoe: am, pin: pin});
    }
  });
  return { ok:true, message:"สร้าง PIN ใหม่ " + added.length + " อำเภอ", added: added };
}

// ===== RESERVOIR PIN MANAGEMENT =====
/** รายการอ่างเก็บน้ำในจังหวัด (master list) — sync กับ reservoir.html */
const RESERVOIR_LIST = [
  { id:"R01", name:"ห้วยยางเงาะ",      amphoe:"เมืองหนองบัวลำภู", capacity:0.400 },
  { id:"R02", name:"ห้วยซับม่วง",       amphoe:"ศรีบุญเรือง",       capacity:0.750 },
  { id:"R03", name:"ห้วยเหล่ายาง",      amphoe:"เมืองหนองบัวลำภู", capacity:2.469 },
  { id:"R04", name:"อ่างน้ำบอง",        amphoe:"โนนสัง",            capacity:20.800 },
  { id:"R05", name:"ห้วยสนามชัย",       amphoe:"นากลาง",            capacity:0.330 },
  { id:"R06", name:"ผาวัง",             amphoe:"นาวัง",             capacity:2.122 },
  { id:"R07", name:"ห้วยลาดกั่ว",       amphoe:"นาวัง",             capacity:0.842 },
  { id:"R08", name:"ห้วยโซ่",           amphoe:"สุวรรณคูหา",        capacity:1.430 },
  { id:"R09", name:"ห้วยไร่ 1",         amphoe:"นากลาง",            capacity:0.200 },
  { id:"R10", name:"ห้วยไร่ 2",         amphoe:"นากลาง",            capacity:0.695 },
  { id:"R11", name:"ห้วยลำใย",          amphoe:"นากลาง",            capacity:0.450 },
  { id:"R12", name:"ห้วยโป่งซาง",       amphoe:"นากลาง",            capacity:0.300 },
  { id:"R13", name:"ห้วยบ้านคลองเจริญ", amphoe:"สุวรรณคูหา",        capacity:0.623 },
  { id:"R14", name:"ผาจ้ำน้ำ",          amphoe:"นาวัง",             capacity:0.085 }
];

function ensureReservoirPinsSheet_() {
  const ss_obj = ss();
  let sh = ss_obj.getSheetByName(SHEET_RESERVOIR_PINS);
  if (!sh) {
    sh = ss_obj.insertSheet(SHEET_RESERVOIR_PINS);
    sh.appendRow(["reservoir_id","pin","recorder_name","phone","last_changed","last_used","note"]);
    sh.getRange("A1:G1").setFontWeight("bold").setBackground("#0c4a6e").setFontColor("#fff");
  }
  return sh;
}

function getReservoirPinRow_(reservoirId) {
  const sh = ensureReservoirPinsSheet_();
  if (sh.getLastRow() < 2) return null;
  const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  const target = String(reservoirId || "").toUpperCase().trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").toUpperCase().trim() === target) {
      return { row: i+2, reservoir_id: target, pin: String(data[i][1]||""), recorder_name: data[i][2]||"", phone: data[i][3]||"", last_changed: data[i][4]||"", last_used: data[i][5]||"", note: data[i][6]||"" };
    }
  }
  return null;
}

function verifyReservoirPin(reservoirId, pin) {
  const r = getReservoirPinRow_(reservoirId);
  if (!r || !r.pin) return false;
  return String(pin||"").trim() === String(r.pin).trim();
}

function touchReservoirPinLastUsed(reservoirId) {
  try {
    const r = getReservoirPinRow_(reservoirId);
    if (r) ensureReservoirPinsSheet_().getRange(r.row, 6).setValue(new Date());
  } catch(e) {}
}

function listReservoirPins() {
  ensureReservoirPinsSheet_();
  return RESERVOIR_LIST.map(r => {
    const row = getReservoirPinRow_(r.id);
    return {
      reservoir_id: r.id, name: r.name, amphoe: r.amphoe, capacity: r.capacity,
      pin: row ? row.pin : "",
      recorder_name: row ? row.recorder_name : "",
      phone: row ? row.phone : "",
      last_changed: row ? (row.last_changed ? Utilities.formatDate(new Date(row.last_changed), "Asia/Bangkok", "yyyy-MM-dd HH:mm") : "") : "",
      last_used: row ? (row.last_used ? Utilities.formatDate(new Date(row.last_used), "Asia/Bangkok", "yyyy-MM-dd HH:mm") : "") : "",
      note: row ? row.note : "",
      has_pin: !!(row && row.pin)
    };
  });
}

function setReservoirPin_(reservoirId, newPin, recorderName) {
  if (!reservoirId) return { ok:false, error:"ไม่ระบุรหัสอ่างเก็บน้ำ" };
  if (!newPin || String(newPin).trim().length < 3) return { ok:false, error:"PIN ต้องมีอย่างน้อย 3 ตัวอักษร" };
  const sh = ensureReservoirPinsSheet_();
  const row = getReservoirPinRow_(reservoirId);
  const now = new Date();
  if (row) {
    sh.getRange(row.row, 2).setValue(String(newPin).trim());
    if (recorderName !== undefined && recorderName !== null) sh.getRange(row.row, 3).setValue(recorderName);
    sh.getRange(row.row, 5).setValue(now);
  } else {
    sh.appendRow([String(reservoirId).toUpperCase(), String(newPin).trim(), recorderName||"", "", now, "", ""]);
  }
  return { ok:true, message:"ตั้ง PIN สำหรับ " + reservoirId + " เรียบร้อย", reservoir_id: reservoirId };
}

function initReservoirPins() {
  ensureReservoirPinsSheet_();
  const added = [];
  RESERVOIR_LIST.forEach(r => {
    const existing = getReservoirPinRow_(r.id);
    if (!existing || !existing.pin) {
      const pin = String(Math.floor(1000 + Math.random()*9000));
      setReservoirPin_(r.id, pin, "");
      added.push({reservoir_id: r.id, name: r.name, pin: pin});
    }
  });
  return { ok:true, message:"สร้าง PIN ใหม่ " + added.length + " อ่าง", added: added };
}

function getReservoirContext(reservoirId) {
  const meta = RESERVOIR_LIST.find(r => r.id === String(reservoirId).toUpperCase());
  if (!meta) return { error:"ไม่พบอ่างเก็บน้ำ " + reservoirId };
  const pinRow = getReservoirPinRow_(reservoirId);
  // Recent records from Reservoir sheet
  const all = getReservoirs() || [];
  const forThis = all.filter(r => String(r.reservoir_id||"").toUpperCase() === meta.id);
  // Sort by date desc, take 10 most recent
  forThis.sort((a,b) => String(b.date||"").localeCompare(String(a.date||"")));
  const recent = forThis.slice(0, 10);
  const last = recent.length ? recent[0] : null;
  const lastVol = last ? parseFloat(last.current_volume) : null;
  const pct = (lastVol != null && !isNaN(lastVol) && meta.capacity > 0) ? (lastVol / meta.capacity * 100) : null;
  let status = "ไม่มีข้อมูล", statusColor = "#94a3b8";
  if (pct != null) {
    if (pct >= 100)     { status = "ล้นความจุ";   statusColor = "#dc2626"; }
    else if (pct >= 80) { status = "น้ำมากเฝ้าระวัง"; statusColor = "#f59e0b"; }
    else if (pct >= 30) { status = "ปกติ";       statusColor = "#10b981"; }
    else                { status = "น้ำน้อย";     statusColor = "#0ea5e9"; }
  }
  return {
    reservoir_id: meta.id, name: meta.name, amphoe: meta.amphoe, capacity: meta.capacity,
    recorder_name: pinRow ? pinRow.recorder_name : "",
    last_volume: lastVol, last_date: last ? last.date : "", last_reporter: last ? last.reporter : "",
    last_pct: pct, status: status, status_color: statusColor,
    recent: recent
  };
}

/** ใช้ใน Login screen — รายชื่อสถานี + สถานะมี PIN หรือยัง */
function getStationListPublic() {
  const stations = getStations();
  ensureStationPinsSheet_();
  const pinMap = {};
  try {
    const sh = ensureStationPinsSheet_();
    if (sh.getLastRow() >= 2) {
      const data = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
      data.forEach(r => pinMap[String(r[0]||"").toUpperCase().trim()] = { has_pin: !!String(r[1]||"").trim(), recorder: r[2]||"" });
    }
  } catch(e) {}
  return stations.map(s => ({
    station_id: s.station_id, name: s.name, river: s.river, amphoe: s.amphoe,
    bank_level: s.bank_level, warn_level: s.warn_level,
    has_pin: !!(pinMap[s.station_id] && pinMap[s.station_id].has_pin),
    recorder_name: (pinMap[s.station_id] && pinMap[s.station_id].recorder) || ""
  }));
}

/** Context หลังล็อกอินสำเร็จ — ข้อมูลสถานี + ระดับน้ำล่าสุด + ประวัติ 7 บันทึก */
function getStationContext(stationId) {
  const stations = getStations();
  const station = stations.find(s => String(s.station_id||"").toUpperCase() === String(stationId).toUpperCase());
  if (!station) return { error:"ไม่พบสถานี " + stationId };
  const pinRow = getStationPinRow_(stationId);
  const recent = getWaterLevels(stationId, 7);
  const last = (recent && recent.length) ? recent[recent.length-1] : null;
  const lv = last ? parseFloat(last.level) : null;
  let status = "ไม่มีข้อมูล", statusColor = "#94a3b8";
  if (lv !== null && !isNaN(lv)) {
    if (lv >= parseFloat(station.bank_level)) { status = "วิกฤติ — ล้นตลิ่ง"; statusColor = "#dc2626"; }
    else if (lv >= parseFloat(station.warn_level)) { status = "เฝ้าระวัง"; statusColor = "#f59e0b"; }
    else { status = "ปกติ"; statusColor = "#10b981"; }
  }
  return {
    station_id: station.station_id, name: station.name, river: station.river,
    village: station.village, amphoe: station.amphoe,
    bank_level: station.bank_level, warn_level: station.warn_level, crit_level: station.crit_level,
    lat: station.lat, lon: station.lon,
    recorder_name: pinRow ? pinRow.recorder_name : "",
    last_level: lv, last_date: last ? last.date : "", last_time: last ? last.time : "",
    last_recorder: last ? last.recorder : "", last_remark: last ? last.remark : "",
    status: status, status_color: statusColor,
    recent: recent.slice(-7).reverse()
  };
}

function getAmphoeContext(amphoe) {
  const pinRow = getAmphoePinRow_(amphoe);
  const rain = getRainfall(7);
  const forAmphoe = rain.filter(r => String(r.amphoe||"") === String(amphoe));
  const last = forAmphoe.length ? forAmphoe[forAmphoe.length-1] : null;
  return {
    amphoe: amphoe,
    recorder_name: pinRow ? pinRow.recorder_name : "",
    last_rain_24hr: last ? last.rain_24hr : null,
    last_rain_7day: last ? last.rain_7day : null,
    last_rain_month: last ? last.rain_month : null,
    last_date: last ? last.date : "",
    last_recorder: last ? last.recorder : "",
    recent: forAmphoe.slice(-7).reverse()
  };
}

// ===== SETUP =====

/** เรียกครั้งเดียวจาก Apps Script Editor หลังอัปเดต Code.gs
 *  เพื่อ append สถานีใหม่ลง Sheet `Stations` โดยไม่กระทบของเดิม
 *  ตรวจตาม station_id ถ้ามีแล้วจะข้าม ถ้ายังไม่มีจะเพิ่ม */
function addMissingStations() {
  const ALL_STATIONS = [
    ["PN01","วังปลาป้อม","ลำน้ำพะเนียง","บ้านโคกเจริญ","นาวัง",17.42065,101.99304,290.0,289.5,290.0,true],
    ["PN02","โคกกระทอ","ลำน้ำพะเนียง","บ้านโคกกระทอ","นาวัง",17.34314,102.07167,266.0,265.5,266.0,true],
    ["PN03","วังสามหาบ","ลำน้ำพะเนียง","บ้านวังสามหาบ","นาวัง",17.30990,102.10789,258.0,257.5,258.0,true],
    ["PN04","บ้านหนองด่าน","ลำน้ำพะเนียง","บ้านหนองด่าน","นากลาง",17.27936,102.16552,249.0,248.5,249.0,true],
    ["PN05","บ้านฝั่งแดง","ลำน้ำพะเนียง","บ้านฝั่งแดง","นากลาง",17.26730,102.22728,237.0,236.5,237.0,true],
    ["PN06","ปตร.หนองหว้าใหญ่","ลำน้ำพะเนียง","บ้านหนองหว้าใหญ่","เมืองหนองบัวลำภู",17.17981,102.38617,216.0,215.5,216.0,true],
    ["PN07","วังหมื่น","ลำน้ำพะเนียง","บ้านวังหมื่น","เมืองหนองบัวลำภู",17.18317,102.43244,210.0,209.5,210.0,true],
    ["PN08","ปตร.ปู่หลอด","ลำน้ำพะเนียง","บ้านโนนคูณ","เมืองหนองบัวลำภู",17.11487,102.45435,203.0,202.5,203.0,true],
    ["PN09","บ้านข้องโป้","ลำน้ำพะเนียง","บ้านข้องโป้","เมืองหนองบัวลำภู",17.08217,102.45068,201.0,200.5,201.0,true],
    ["PN10","ปตร.หัวนา","ลำน้ำพะเนียง","บ้านดอนหัน","เมืองหนองบัวลำภู",17.00067,102.42400,191.0,190.5,191.0,true],
    ["MG01","คลองบุญทัน","ลำน้ำโมง","บ้านบุญทัน","สุวรรณคูหา",17.54512,102.16832,231.0,230.5,231.0,true],
    ["MG02","บ้านโคก","ลำน้ำโมง","บ้านโคก","สุวรรณคูหา",17.54952,102.20425,218.0,217.5,218.0,true],
    ["MG03","บ้านนาตาแหลว","ลำน้ำโมง","บ้านโคก","สุวรรณคูหา",17.57567,102.27326,202.0,201.5,202.0,true],
    ["MG04","บ้านกุดผึ้ง","ลำน้ำโมง","บ้านกุดผึ้ง","สุวรรณคูหา",17.56062,102.31572,192.0,191.5,192.0,true],
    ["MO01","อ่างเก็บน้ำมอ","ลำน้ำมอ","บ้านฝายหิน","ศรีบุญเรือง",17.16608,102.18177,242.0,241.5,242.0,true],
    ["MO02","บ้านวังคูณ","ลำน้ำมอ","บ้านวังคูณ","ศรีบุญเรือง",17.03214,102.24920,211.0,210.5,211.0,true],
    ["MO03","บ้านโนนสูงเปลือย","ลำน้ำมอ","บ้านโนนสูงเปลือย","ศรีบุญเรือง",16.96934,102.27002,202.0,201.5,202.0,true],
    ["PY01","บ้านวังโปร่ง","ลำน้ำพวย","บ้านวังโปร่ง","ศรีบุญเรือง",17.01415,102.19359,212.0,211.5,212.0,true],
    ["PY02","บ้านทุ่งโพธิ์","ลำน้ำพวย","บ้านทุ่งโพธิ์","ศรีบุญเรือง",16.97482,102.22344,197.0,196.5,197.0,true],
    ["PY03","บ้านโคกล่าม","ลำน้ำพวย","บ้านโคกล่าม","ศรีบุญเรือง",16.91317,102.23807,193.0,192.5,193.0,true],
  ];
  const ss_obj = ss();
  let sh = ss_obj.getSheetByName(SHEET_STATIONS);
  if (!sh) {
    sh = ss_obj.insertSheet(SHEET_STATIONS);
    sh.appendRow(["station_id","name","river","village","amphoe","lat","lon","bank_level","warn_level","crit_level","active"]);
  }
  const data = sh.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || "").trim().toUpperCase();
    if (id) existing[id] = true;
  }
  let added = 0, skipped = 0;
  const addedIds = [];
  ALL_STATIONS.forEach(row => {
    const id = String(row[0]).toUpperCase();
    if (existing[id]) { skipped++; return; }
    sh.appendRow(row);
    added++;
    addedIds.push(id);
  });
  const msg = "เพิ่มสถานีใหม่ " + added + " รายการ (ข้ามของเดิม " + skipped + " รายการ)" +
              (addedIds.length ? " — " + addedIds.join(", ") : "");
  Logger.log(msg);
  return msg;
}

function runSetup() {
  const ss_obj=ss();
  function ensure(name,hdr){let sh=ss_obj.getSheetByName(name);if(!sh){sh=ss_obj.insertSheet(name);}if(sh.getLastRow()===0)sh.appendRow(hdr);return sh;}
  ensure(SHEET_STATIONS,["station_id","name","river","village","amphoe","lat","lon","bank_level","warn_level","crit_level","active"]);
  ensure(SHEET_WATER,   ["station_id","date","time","level","flow","recorder","remark"]);
  ensure(SHEET_RAIN,    ["station_id","amphoe","date","rain_24hr","rain_7day","rain_month","recorder","remark"]);
  ensure(SHEET_RESERVOIR,RESERVOIR_HEADERS);
  ensure(SHEET_SETTINGS,["key","value"]);
  const stSh=ss_obj.getSheetByName(SHEET_STATIONS);
  if(stSh.getLastRow()<=1){
    [["PN01","วังปลาป้อม","ลำน้ำพะเนียง","บ้านโคกเจริญ","นาวัง",17.42065,101.99304,290.0,289.5,290.0,true],
     ["PN02","โคกกระทอ","ลำน้ำพะเนียง","บ้านโคกกระทอ","นาวัง",17.34314,102.07167,266.0,265.5,266.0,true],
     ["PN03","วังสามหาบ","ลำน้ำพะเนียง","บ้านวังสามหาบ","นาวัง",17.30990,102.10789,258.0,257.5,258.0,true],
     ["PN04","บ้านหนองด่าน","ลำน้ำพะเนียง","บ้านหนองด่าน","นากลาง",17.27936,102.16552,249.0,248.5,249.0,true],
     ["PN05","บ้านฝั่งแดง","ลำน้ำพะเนียง","บ้านฝั่งแดง","นากลาง",17.26730,102.22728,237.0,236.5,237.0,true],
     ["PN06","ปตร.หนองหว้าใหญ่","ลำน้ำพะเนียง","บ้านหนองหว้าใหญ่","เมืองหนองบัวลำภู",17.17981,102.38617,216.0,215.5,216.0,true],
     ["PN07","วังหมื่น","ลำน้ำพะเนียง","บ้านวังหมื่น","เมืองหนองบัวลำภู",17.18317,102.43244,210.0,209.5,210.0,true],
     ["PN08","ปตร.ปู่หลอด","ลำน้ำพะเนียง","บ้านโนนคูณ","เมืองหนองบัวลำภู",17.11487,102.45435,203.0,202.5,203.0,true],
     ["PN09","บ้านข้องโป้","ลำน้ำพะเนียง","บ้านข้องโป้","เมืองหนองบัวลำภู",17.08217,102.45068,201.0,200.5,201.0,true],
     ["PN10","ปตร.หัวนา","ลำน้ำพะเนียง","บ้านดอนหัน","เมืองหนองบัวลำภู",17.00067,102.42400,191.0,190.5,191.0,true],
     ["MG01","คลองบุญทัน","ลำน้ำโมง","บ้านบุญทัน","สุวรรณคูหา",17.54512,102.16832,231.0,230.5,231.0,true],
     ["MG02","บ้านโคก","ลำน้ำโมง","บ้านโคก","สุวรรณคูหา",17.54952,102.20425,218.0,217.5,218.0,true],
     ["MG03","บ้านนาตาแหลว","ลำน้ำโมง","บ้านโคก","สุวรรณคูหา",17.57567,102.27326,202.0,201.5,202.0,true],
     ["MG04","บ้านกุดผึ้ง","ลำน้ำโมง","บ้านกุดผึ้ง","สุวรรณคูหา",17.56062,102.31572,192.0,191.5,192.0,true],
     ["MO01","อ่างเก็บน้ำมอ","ลำน้ำมอ","บ้านฝายหิน","ศรีบุญเรือง",17.16608,102.18177,242.0,241.5,242.0,true],
     ["MO02","บ้านวังคูณ","ลำน้ำมอ","บ้านวังคูณ","ศรีบุญเรือง",17.03214,102.24920,211.0,210.5,211.0,true],
     ["MO03","บ้านโนนสูงเปลือย","ลำน้ำมอ","บ้านโนนสูงเปลือย","ศรีบุญเรือง",16.96934,102.27002,202.0,201.5,202.0,true],
     ["PY01","บ้านวังโปร่ง","ลำน้ำพวย","บ้านวังโปร่ง","ศรีบุญเรือง",17.01415,102.19359,212.0,211.5,212.0,true],
     ["PY02","บ้านทุ่งโพธิ์","ลำน้ำพวย","บ้านทุ่งโพธิ์","ศรีบุญเรือง",16.97482,102.22344,197.0,196.5,197.0,true],
     ["PY03","บ้านโคกล่าม","ลำน้ำพวย","บ้านโคกล่าม","ศรีบุญเรือง",16.91317,102.23807,193.0,192.5,193.0,true],
    ].forEach(r=>stSh.appendRow(r));
  }
  const resSh=ss_obj.getSheetByName(SHEET_RESERVOIR);
  if(resSh.getLastRow()<=1){
    const today=Utilities.formatDate(new Date(),"Asia/Bangkok","yyyy-MM-dd");
    [["R01","ห้วยยางเงาะ","เมืองหนองบัวลำภู",0.400,0.240],["R02","ห้วยซับม่วง","ศรีบุญเรือง",0.750,0.450],
     ["R03","ห้วยเหล่ายาง","เมืองหนองบัวลำภู",2.469,1.481],["R04","อ่างน้ำบอง","โนนสัง",20.800,9.984],
     ["R05","ห้วยสนามชัย","นากลาง",0.330,0.198],["R06","ผาวัง","นาวัง",2.122,1.273],
     ["R07","ห้วยลาดกั่ว","นาวัง",0.842,0.505],["R08","ห้วยโซ่","สุวรรณคูหา",1.430,0.858],
     ["R09","ห้วยไร่ 1","นากลาง",0.200,0.120],["R10","ห้วยไร่ 2","นากลาง",0.695,0.417],
     ["R11","ห้วยลำใย","นากลาง",0.450,0.270],["R12","ห้วยโป่งซาง","นากลาง",0.300,0.180],
     ["R13","ห้วยบ้านคลองเจริญ","สุวรรณคูหา",0.623,0.374],["R14","ผาจ้ำน้ำ","นาวัง",0.085,0.051],
    ].forEach(d=>resSh.appendRow([d[0],d[1],d[2],d[3],d[4],today,"ระบบ",new Date()]));
  }
  Logger.log("Setup complete ✅");
}

// ===== HELPERS =====
function getOrCreateReservoirSheet(){const ss_obj=ss();let sh=ss_obj.getSheetByName(SHEET_RESERVOIR);if(!sh){sh=ss_obj.insertSheet(SHEET_RESERVOIR);sh.appendRow(RESERVOIR_HEADERS);}if(sh.getLastRow()===0||sh.getLastColumn()===0)sh.appendRow(RESERVOIR_HEADERS);return sh;}
function normalizeReservoirPayload(p){return{reservoir_id:p.reservoir_id||p.id||"",reservoir_name:p.reservoir_name||p.name||"",amphoe:p.amphoe||"",capacity:toNum(p.capacity),current_volume:toNum(p.current_volume!==undefined?p.current_volume:p.current),date:p.date||Utilities.formatDate(new Date(),"Asia/Bangkok","yyyy-MM-dd"),reporter:p.reporter||"",updated_at:new Date()};}
function toNum(v){if(v===""||v===null||v===undefined)return "";const n=parseFloat(v);return isNaN(n)?"":n;}
function headerIndex(headers,names){for(let i=0;i<names.length;i++){const idx=headers.indexOf(names[i]);if(idx>=0)return idx;}return -1;}
function ss(){return SpreadsheetApp.getActiveSpreadsheet();}
function getHeaders(sheet){return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];}
function sheetToObjects(sheet){if(!sheet)return[];const data=sheet.getDataRange().getValues();if(data.length<2)return[];const headers=data[0];return data.slice(1).map(row=>{const obj={};headers.forEach((h,i)=>{let v=row[i];if(v instanceof Date)v=Utilities.formatDate(v,"Asia/Bangkok","yyyy-MM-dd");obj[h]=v;});return obj;}).filter(o=>Object.values(o).some(v=>v!==""&&v!==null&&v!==undefined));}
function parseDate(v){if(!v)return null;if(v instanceof Date)return v;const d=new Date(v);return isNaN(d.getTime())?null:d;}
function respond(data,callback){const json=JSON.stringify(data);if(callback)return ContentService.createTextOutput(callback+"("+json+");").setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);}
function getAppPin(){return String(PropertiesService.getScriptProperties().getProperty(PIN_PROPERTY_KEY)||"").trim();}
function installDefaultPinForSetup(){PropertiesService.getScriptProperties().setProperty(PIN_PROPERTY_KEY,"123456");Logger.log("ตั้งค่า APP_PIN เริ่มต้นแล้ว ควรเปลี่ยนก่อนใช้งานจริง");}

// ===== TEST =====
function testPin(){Logger.log(getAppPin()?"✅ ตั้งค่า APP_PIN แล้ว":"❌ ยังไม่ได้ตั้งค่า APP_PIN");}
function testGetSummary(){Logger.log(JSON.stringify(getSummary(),null,2));}
function testGetReservoirs(){Logger.log(JSON.stringify(getReservoirs(),null,2));}
