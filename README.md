# ⚔️ War of Wannakhadi (มหาศึกวรรณคดี)

มหาศึกวรรณคดี (War of Wannakhadi) เกมเพื่อความสนุกและการเรียนรู้ รูปแบบบอร์ดเกม RPG ออนไลน์หลายผู้เล่น เชื่อมต่อข้อมูลผ่าน Google Sheets และปรับให้เล่นแบบเรียลไทม์ได้อย่างลื่นไหลทั้งบนคอมพิวเตอร์และมือถือ (รองรับ PWA)

---

## ✨ ฟีเจอร์หลัก
- **รองรับหลายผู้เล่นแบบเรียลไทม์:** ซิงค์สถานะห้องผ่าน Google Apps Script (`CacheService` + `LockService` + Version Delta Polling) ทำให้การอัปเดตเทิร์น สกิล และแอนิเมชันลื่นไหล ลดดีเลย์
- **เชื่อมต่อฐานข้อมูล Google Sheets:** ดึงข้อมูลกระดาน (`BoardConfig`), ไอเท็ม (`Items`), ตัวละคร (`Characters`), เหตุการณ์โลก (`Events`), คำถาม (`คำถาม`) และห้องเล่น (`Rooms`) โดยตรงจาก Google Sheets
- **รองรับมือถือและ PWA:** ติดตั้งลงหน้าจอโฮมบน iOS / Android ได้ทันทีผ่าน `manifest.json` และ `sw.js`

---

## 📁 โครงสร้างไฟล์
- `index.html` — หน้าเกมทั้งหมด (UI, 3D Board, Combat/Boss/Quiz/Shop Modals, GAS API Bridge, PWA)
- `Code.gs` — Google Apps Script Backend API (`doGet` / `doPost` JSON API + Cache & Lock)
- `appsscript.json` — การตั้งค่าโปรเจกต์ Google Apps Script
- `manifest.json` & `sw.js` — การตั้งค่า Progressive Web App (PWA)
- `.nojekyll` — ปิดการใช้งาน Jekyll บน GitHub Pages
