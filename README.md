# 🚀 LMDS Version 5.0: Logistics Master Data System

ระบบจัดการข้อมูลหลักด้านโลจิสติกส์ (Master Data Management) ที่ถูกออกแบบใหม่ทั้งหมด (Re-architecture) เพื่อแก้ปัญหาความซ้ำซ้อนของข้อมูล ชื่อสถานที่ และพิกัดที่ไม่ถูกต้อง โดยแยกโครงสร้างฐานข้อมูลออกเป็นเชิงสัมพันธ์ (Relational Model) บน Google Sheets

## 🌟 สิ่งที่ใหม่ใน Version 5.0
- **Separation of Concerns:** แยก "ตัวตน (Entity)" ออกจาก "สถานที่ (Location)" อย่างชัดเจน
- **Data Ingestion Pipeline:** ระบบรับข้อมูลดิบจาก `SCGนครหลวงJWDภูมิภาค` เพื่อทำความสะอาดก่อนเข้าสู่ฐานข้อมูลหลัก
- **Conflict Resolution UI:** ระบบจัดการข้อขัดแย้งผ่านหน้าจอ UI ที่ง่ายต่อการอนุมัติ
- **No AI Dependency:** ทำงานด้วยตรรกะและกฎเกณฑ์ที่ชัดเจน (Deterministic Logic) 100% มั่นใจได้ในความเสถียร
- **Modular Architecture:** แยกโค้ดเป็น 10 โมดูล ดูแลรักษาง่าย ขยายต่อได้สะดวก

## 🏗️ สถาปัตยกรรมระบบ
ระบบแบ่งการทำงานเป็น 2 ส่วนหลักที่ประสานกัน:

### ส่วนที่ 1: Daily Operations (ปฏิบัติงานรายวัน)
- **แหล่งข้อมูล:** SCG API
- **กระบวนการ:** Input -> API -> Data Sheet -> Match with Master -> Summary
- **ผลลัพธ์:** ชีต `Data` ที่มีพิกัดสะอาด (`LatLong_Actual`) และอีเมลพนักงาน พร้อมชีตสรุป

### ส่วนที่ 2: Master Data Engine (เครื่องยนต์ฐานข้อมูล)
- **แหล่งข้อมูล:** ชีตดิบ `SCGนครหลวงJWDภูมิภาค`
- **กระบวนการ:** Raw Data -> Clean/Normalize -> Resolve Entity/Location -> Handle Conflicts -> Master Sheets
- **ผลลัพธ์:** ฐานข้อมูลสะอาดในชีต `Entities`, `Locations`, `Entity_Loc_Map`

## 📂 โครงสร้างไฟล์ (Modules)
1. `V5_Config.gs`: การตั้งค่ากลาง
2. `V5_Setup.gs`: สร้างโครงสร้างชีต
3. `V5_Utils_Helper.gs`: ฟังก์ชันช่วยเหลือ (UUID, Distance, Normalize)
4. `V5_Service_Master_Core.gs`: สมองกลจับคู่ Entity/Location
5. `V5_Service_DataIngestion.gs`: รับข้อมูลดิบเข้าฐาน
6. `V5_Service_Daily_Match.gs`: จับคู่ข้อมูลรายวันกับฐานข้อมูลหลัก
7. `V5_Service_SCG_API.gs`: ดึงข้อมูลจาก API
8. `V5_Service_Summary.gs`: สร้างรายงานสรุป
9. `V5_Service_ConflictManager.gs`: จัดการข้อขัดแย้ง + UI
10. `V5_Menu.gs`: เมนูควบคุมระบบ

## 🚀 การเริ่มต้นใช้งาน
1. นำโค้ดทั้ง 10 ไฟล์ไปวางใน Google Apps Script
2. รันฟังก์ชัน `V5_InitAllSheets` จากเมนู `🚀 LMDS V5` > `⚙️ Setup & Config`
3. เริ่มนำเข้าข้อมูลดิบผ่านเมนู `ส่วนที่ 2: จัดการฐานข้อมูล`
4. เริ่มปฏิบัติงานรายวันผ่านเมนู `ส่วนที่ 1: ปฏิบัติการรายวัน`

## 📄 เอกสารประกอบ
- [BLUEPRINT_โปรเจกต์ LMDS Version 5.0.md](./BLUEPRINT_โปรเจกต์_LMDS_Version_5.0.md)
- [เอกสารคู่มือการใช้งาน_SOP_โปรเจกต์ LMDS Version 5.0.md](./เอกสารคู่มือการใช้งาน_SOP_โปรเจกต์_LMDS_Version_5.0.md)

---
*Developed for SCG/JWD Logistics Data Integrity*
