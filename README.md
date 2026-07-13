# BangYai BoxMaster

ระบบติดตามกล่องยา HAD / CPR — โรงพยาบาลบางใหญ่

## วิธีใช้กับ GitHub Pages
1. สร้าง repository ใหม่ แล้วอัพโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ (ให้ `index.html` อยู่ที่ root)
2. ไปที่ **Settings → Pages** → Source: **Deploy from a branch** → เลือก branch `main` / root → Save
3. รอสักครู่ แล้วเปิด URL ที่ GitHub แจ้ง (เช่น `https://username.github.io/repo-name/`)

## ไฟล์
- `index.html` — ตัวแอปทั้งหมด (ไฟล์เดียว ไม่ต้องพึ่งไฟล์อื่น)
- `Code.gs` — โค้ดสำหรับวางใน Google Apps Script เพื่อใช้ Google Sheets เป็นฐานข้อมูลกลาง (ดูวิธีตั้งค่าในหน้า "จัดการ" ของแอป)
