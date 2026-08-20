# iFind — ระบบค้นหาค่าสี & คุณภาพน้ำตาล (Mitr Lao Sugar)

แอปมือถือ (Google Apps Script Web App) สำหรับ **ค้นหาและคัดแยกน้ำตาลตามค่าสี (ICUMSA) และค่าคุณภาพ** ในคลังสินค้า
ใช้ **Google Sheets เป็นฐานข้อมูล** รองรับผู้ใช้ 2 บทบาทแยกด้วย PIN

| บทบาท | ทำอะไรได้ | PIN เริ่มต้น |
|---|---|---|
| **QC** | บันทึก/แก้ไขค่าคุณภาพ (สี, Pol, ความชื้น, อินเวิร์ต, M.A.) | `1111` |
| **คลังสินค้า** | ค้นหา Lot, คัดค่าที่เหมาะสม, จัดการสต๊อก (ตำแหน่ง+จำนวน) | `2222` |

> เปลี่ยน PIN/เกณฑ์ Spec ได้ในชีต `CONFIG` โดยไม่ต้องแก้โค้ด

---

## โครงสร้างโปรเจกต์

```
iFind/
├─ src/                 โค้ด Apps Script (โฟลเดอร์สำหรับ clasp)
│  ├─ appsscript.json   manifest (เขตเวลา, web app config)
│  ├─ Code.gs           backend: PIN, อ่าน/เขียน, ตรวจ Spec
│  ├─ Migrate.gs        ย้ายข้อมูล COA 52 แผ่น → ชีต DATA
│  └─ Index.html        หน้าจอแอป (Home, ค้นหา, บันทึก)
├─ demo/                ตัวอย่างเปิดในเบราว์เซอร์ (mock data)
├─ assets/              ภาพดีไซน์ (icon, Head, การ์ด ฯลฯ)
└─ docs/install-th.md   คู่มือติดตั้งภาษาไทยแบบละเอียด
```

ฐานข้อมูล: [Google Sheet “คลังสินค้า คาสี”](https://docs.google.com/spreadsheets/d/1Fhyt32uFb7eyerdXvM9zuYOB9sOGwsF4SX1my_936aM/edit)

---

## ติดตั้งครั้งแรก (ดูละเอียดใน `docs/install-th.md`)

1. เปิดชีต → **Extensions → Apps Script** → วางไฟล์ใน `src/` (Code.gs, Migrate.gs, Index)
2. Run `setup` → สร้างแท็บ `DATA` + `CONFIG`
3. Import ไฟล์ `.xlsx` เข้าชีต → Run `migrate` → ย้ายข้อมูลเข้า DATA
4. **Deploy → Web app** → Execute as *Me*, Access *Anyone* → เปิดลิงก์บนมือถือ

---

## อัปเดตโค้ดด้วย clasp (ไม่ต้องก๊อปวางมือ)

```bash
npm install -g @google/clasp     # ติดตั้งครั้งเดียว
clasp login                       # ล็อกอิน Google
cp .clasp.json.example .clasp.json   # แล้วใส่ scriptId ของโปรเจกต์
clasp push                        # อัปโค้ดใน src/ ขึ้น Apps Script
```

หา `scriptId` ได้จาก Apps Script → Project Settings → “IDs”

---

## เกณฑ์คุณภาพ (Spec)

สี Color ≤ 1,200 ICU · Pol ≥ 99.00 % · ความชื้น ≤ 0.20 % · อินเวิร์ต ≤ 0.80 % · M.A. 0.75–1.20 mm

สถานะ: 🟢 ผ่าน · 🟡 เฝ้าระวัง (เกิน 1 ค่า) · 🔴 ไม่ผ่าน (เกิน ≥2 ค่า)
