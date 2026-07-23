# คู่มือการติดตั้ง — ศูนย์บัญชาการข้อมูลน้ำ จ.หนองบัวลำภู

## 1. ไฟล์ที่ต้องอัปทับบน GitHub (repo เดิม branch ที่ใช้งาน)

| ไฟล์ | สถานะ |
|---|---|
| `ds.css` | **ไฟล์ใหม่** — Design System กลาง (สี/ฟอนต์/เลย์เอาต์) ต้องอัปด้วยเสมอ |
| `ds.js` | **ไฟล์ใหม่** — ระบบ Tooltip กลาง ต้องอัปด้วยเสมอ |
| `favicon.svg` | **ไฟล์ใหม่** — ไอคอนบนแท็บเบราว์เซอร์ |
| `index.html` | **หน้าปกนำเสนอ** — เปลี่ยนเป็นหน้าแรกของเว็บแล้ว |
| `dashboard.html` | **ชื่อใหม่** — แดชบอร์ดผู้บริหาร (เดิมคือ index.html) |
| `logo-nbp.png` | **แทนที่** — ตราจังหวัดความละเอียดใหม่ (200 px) |
| `logo-nbp-lg.png` | **ไฟล์ใหม่** — ตราจังหวัดขนาดใหญ่ สำหรับหน้าปก (420 px) |

| `daily_briefing.html` | แทนที่ของเดิม |
| `paneang.html` `mong.html` `mo.html` `phuay.html` | แทนที่ของเดิม |
| `rainfall.html` `reservoir.html` | แทนที่ของเดิม |
| `input.html` | แทนที่ของเดิม (ฟอร์มเดิม + ระบบออฟไลน์) |
| `sw.js` | แทนที่ของเดิม (VERSION = wnb-v7) |

ไฟล์เดิมที่ **ยังต้องมีอยู่ในโฟลเดอร์** (ไม่ต้องแก้): `theme.css`, `config.js`,
`manifest.json`, `logo-nbp.png`, `favicon.ico`, `favicon-32.png`,
`icon-192.png`, `icon-512.png`

`index_legacy_patchv2.html` = หน้า index เดิม เก็บไว้เผื่อย้อนกลับ (ไม่ต้องอัปก็ได้)

## 2. หลังอัปเสร็จ
กด `Ctrl + Shift + R` (ล้าง cache) — sw.js ขยับเป็น **wnb-v13** แล้ว
ระบบจะล้าง cache เก่าให้ผู้ใช้ทุกคนอัตโนมัติ

## 3. สิ่งที่เปลี่ยนใน sw.js
- `VERSION` : wnb-v6 → **wnb-v13**
- เพิ่ม `./ds.css`, `./ds.js`, `./favicon.svg` เข้า precache
- **ตัด** `nbp_water_lines.geojson` และ `nbp_water_points.geojson` ออกจาก precache
  (หน้าใหม่ฝังขอบเขตตำบลไว้ในไฟล์แล้ว ไม่ต้องโหลดไฟล์เส้นน้ำ ~5 MB อีก
  → เปิดเว็บครั้งแรกเร็วขึ้นมาก) ถ้าต้องการย้อนกลับ ให้เพิ่ม 2 บรรทัดนี้กลับเข้า `PRECACHE_URLS`

## 4. เพิ่ม action ที่ยังขาดใน Google Apps Script

หน้าลำน้ำเรียก `?action=paneang | mong | mo | phuay`
ถ้ามีแค่ `paneang` ให้เพิ่มอีก 3 ตัว โดยใช้ handler เดียวกันแล้วกรองตามลำน้ำ

```javascript
// ── ใน doGet(e) เพิ่มเงื่อนไขเหล่านี้ ─────────────────────────
var action = (e.parameter.action || '').toLowerCase();

var RIVER_OF = {
  paneang: 'ลำน้ำพะเนียง',
  mong:    'ลำน้ำโมง',
  mo:      'ลำน้ำมอ',
  phuay:   'ลำน้ำพวย'
};

if (RIVER_OF[action]) {
  return jsonOut(getRiverData(RIVER_OF[action]));
}

// ── ฟังก์ชันดึงข้อมูลรายลำน้ำ ────────────────────────────────
function getRiverData(riverName) {
  var sh   = SpreadsheetApp.getActive().getSheetByName('WaterLevel'); // ← ชื่อชีตจริง
  var rows = sh.getDataRange().getValues();
  var head = rows.shift();

  var iId    = head.indexOf('station_id');
  var iRiver = head.indexOf('river');
  var iLevel = head.indexOf('water_level');
  var iDate  = head.indexOf('date');

  var out = [];
  rows.forEach(function (r) {
    if (String(r[iRiver]).trim() !== riverName) return;
    if (r[iLevel] === '' || r[iLevel] === null) return;
    out.push({
      station_id : String(r[iId]).trim(),
      water_level: parseFloat(r[iLevel]),
      date       : Utilities.formatDate(new Date(r[iDate]),
                     'Asia/Bangkok', 'yyyy-MM-dd')
    });
  });
  return out;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

**รูปแบบข้อมูลที่หน้าเว็บต้องการ** (เรียงตามวันที่ ใช้ทำกราฟย้อนหลังได้ด้วย)

```json
[
  { "station_id": "PN01", "water_level": 288.20, "date": "2026-07-20" },
  { "station_id": "PN01", "water_level": 288.35, "date": "2026-07-21" }
]
```

> ถ้ายังไม่เพิ่ม action หน้าเว็บจะไม่พัง — จะขึ้นชิปเหลือง **"● ชุดตัวอย่าง"**
> มุมขวาบน และแสดงข้อมูลตัวอย่างแทน เพื่อไม่ให้เข้าใจผิดว่าเป็นสถานการณ์จริง

## 5. action อื่นที่หน้าเว็บใช้ (มีอยู่แล้ว)
`?action=hydro` · `?action=reservoir` · `?action=floodgate` · `?action=rain&days=7`

## 6. แก้โทนทั้งระบบทีเดียว
สี ฟอนต์ เงา ขอบมน อยู่ใน `ds.css` ส่วนบนสุด (`:root`)
แก้ตัวแปรที่นี่ที่เดียว เปลี่ยนทั้ง 9 หน้าพร้อมกัน

```css
:root{
  --navy:#0a1e3c;   /* สีหลัก (แถบข้าง/หัวข้อ) */
  --gold:#c9a24b;   /* สีเน้นแบบทางการ */
  --ok:#16a34a;     /* ปกติ */
  --warn:#f59e0b;   /* เฝ้าระวัง */
  --crit:#dc2626;   /* วิกฤติ */
}
```


## 7. ระบบ Tooltip กลาง (ds.js)

ทุกหน้ามีคำอธิบายเมื่อชี้เมาส์แล้ว — ถ้าจะเพิ่มจุดใหม่ ทำได้ 2 แบบ

```html
<!-- แบบข้อความธรรมดา -->
<div data-tip="คำอธิบายที่จะแสดงเมื่อชี้">...</div>
```

```javascript
/* แบบมีหลอดเกณฑ์/ตาราง — กำหนดเนื้อหาเป็น HTML */
window.TIPS['R01'] = '<div class="tth">...</div><div class="body">...</div>';
/* แล้วใส่ data-tipid="R01" ที่ element */
```

ไม่ต้องเขียนโค้ดเพิ่ม — `ds.js` จับ event ให้อัตโนมัติทั้งหน้า
รวมถึง element ที่สร้างขึ้นภายหลังด้วย JavaScript


## 8. หน้าปกนำเสนอ (index.html)

หน้าปกเป็น **หน้าแรกของเว็บ** แล้ว — เปิด URL หลักจะเจอปกก่อน กดปุ่ม
"เข้าสู่แดชบอร์ด" เพื่อไปหน้า `dashboard.html` และคลิกตราจังหวัดที่แถบข้าง
ของหน้าใดก็ได้เพื่อกลับมาหน้าปก

**สำคัญ — ต้องเปลี่ยนชื่อไฟล์เดิมในรีโปด้วย**
ไฟล์ `index.html` เดิม (แดชบอร์ด) ถูกเปลี่ยนชื่อเป็น `dashboard.html`
เวลาอัปโหลดให้อัป `index.html` (ปก) และ `dashboard.html` (แดชบอร์ด) ทั้งคู่

### ใส่ภาพพื้นหลังหน้าปก
วางไฟล์ไว้โฟลเดอร์เดียวกับ `index.html` ระบบจะเลือกให้เองตามลำดับ

| ลำดับ | ชื่อไฟล์ | ใช้เมื่อ |
|---|---|---|
| 1 | `cover-bg.gif` | ภาพเคลื่อนไหว |
| 2 | `cover-bg.webp` | ภาพเคลื่อนไหว (ไฟล์เล็กกว่า gif มาก — แนะนำ) |
| 3 | `cover-bg.jpg` | ภาพนิ่ง |
| 4 | `cover-bg.png` | ภาพนิ่ง |

ถ้าไม่มีไฟล์ใดเลย ปกจะใช้พื้นหลังไล่เฉด + ลายขอบเขตจังหวัดแทน — ไม่พัง

- แนะนำภาพแนวนอน อย่างน้อย 1920×1080 px
- ระบบคลุมชั้นสีเข้มโปร่งแสงให้อัตโนมัติ ตัวหนังสืออ่านออกเสมอ
- **GIF ควรมีขนาดไม่เกิน ~5 MB** ถ้าใหญ่กว่านั้นเปิดบนมือถือจะช้ามาก
  แนะนำแปลงเป็น `.webp` แบบเคลื่อนไหว จะเล็กลง 5–10 เท่าโดยคุณภาพใกล้เคียง

## 9. หน้าบันทึกข้อมูล (input.html)

- ตัดแท็บ "บันทึกฝนรายอำเภอ" ออกแล้ว — หน้าปริมาณฝนดึงข้อมูลจาก Open-Meteo
  อัตโนมัติ (ถ้ามีคนกรอกค่าตรวจวัดจริงในระบบ ระบบจะใช้ค่าตรวจวัดก่อนเสมอ)
- โค้ดส่วนกรอกฝนยังอยู่ครบ เปิดแท็บคืนได้โดยเพิ่มปุ่มใน `.mode-tabs`
  พร้อม `data-mode="rain"` — `setMode()` รองรับอยู่แล้ว
- แก้บั๊กสำคัญ: `ds.css` ล็อก `overflow:hidden` ทำให้ฟอร์มยาวเลื่อนไม่ได้
  หน้านี้จึง override ให้เลื่อนได้ และเพิ่มการรองรับแท็บเล็ต/มือถือ
