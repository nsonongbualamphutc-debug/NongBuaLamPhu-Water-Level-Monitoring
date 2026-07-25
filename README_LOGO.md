# Patch v3 — เพิ่มตราจังหวัดหนองบัวลำภู แทนไอคอนหยดน้ำ/emoji

## ที่มาโลโก้
ดึงจาก repo Drug-situation-dashboard-Nong-Bua-Lamphu-Province
(ไฟล์ต้นฉบับ: ตราหนองบัวลำภู.png) แล้วประมวลผลใหม่:
- ตัดขอบโปร่งใสให้พอดี, เพิ่มระยะขอบเล็กน้อย
- logo-nbp.png (96×96 โปร่งใส) → ใช้ในกล่องไอคอนหัวข้อทุกหน้า
- icon-192.png / icon-512.png (พื้นกรมท่า #0a1e3c) → PWA install icon
- favicon.ico + favicon-32.png → แท็บเบราว์เซอร์

## ไฟล์ใหม่ (6)
logo-nbp.png, icon-192.png, icon-512.png, favicon.ico, favicon-32.png
(icon-192/512 เขียนทับไฟล์เดิมที่เป็น SVG หยดน้ำ)

## ไฟล์ที่แก้ (9)
ทุกหน้า (index, paneang, mong, mo, phuay, reservoir, input) — แทนที่ SVG/emoji
ไอคอนเดิมในกล่องหัวข้อด้วย <img src="logo-nbp.png">, เพิ่ม <link rel="icon">
+ manifest.json (ชี้ไปไฟล์ png จริงแทน data:svg เดิม)
+ sw.js (VERSION wnb-v5, precache รูปใหม่ทั้งหมด)

## วิธีติดตั้ง
อัปทุกไฟล์ในชุดนี้ทับของเดิมบน GitHub Pages แล้ว Ctrl+Shift+R
(ถ้าเคยติดตั้งเป็น PWA ในมือถือ — ต้องถอนแล้วติดตั้งใหม่ไอคอนจึงจะเปลี่ยน
เพราะ OS แคชไอคอนแยกจาก Service Worker)

## วิธีถอด patch
เปลี่ยน <img src="logo-nbp.png" ...> กลับเป็น SVG/emoji เดิม (อยู่ใน git history)
หรือแทนที่ manifest.json ด้วยเวอร์ชัน data:svg เดิม
