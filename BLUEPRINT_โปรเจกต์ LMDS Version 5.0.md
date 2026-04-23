# 📘 BLUEPRINT: LMDS Version 5.0 Architecture

## 1. วัตถุประสงค์
เพื่อสร้างฐานข้อมูลสถานที่และลูกค้า (Master Data) ที่ "สะอาด" "ไม่ซ้ำซ้อน" และ "ตรวจสอบได้" โดยแก้ปัญหา 8 ประการเรื่องชื่อซ้ำ ที่อยู่ซ้ำ และพิกัดผิดพลาด ผ่านการแยกโมเดลข้อมูลแบบ Entity-Location Relationship

## 2. โครงสร้างฐานข้อมูล (Database Schema)

### 2.1 ชีตหลักใหม่ (Part 2: Master Data)
| ชื่อชีต | หน้าที่ | คอลัมน์สำคัญ (Key Columns) |
| :--- | :--- | :--- |
| **Entities** | เก็บข้อมูล "ใคร" (ลูกค้า/ร้าน) | `Entity_ID` (PK), `Display_Name`, `Normalized_Name`, `Status` |
| **Locations** | เก็บข้อมูล "ที่ไหน" (พิกัด/ที่อยู่) | `Location_ID` (PK), `Lat`, `Lng`, `LatLong_Key`, `Full_Address` |
| **Entity_Loc_Map** | เชื่อมความสัมพันธ์ "ใคร-อยู่ที่ไหน" | `Map_ID`, `Entity_ID` (FK), `Location_ID` (FK), `Relation_Type`, `Is_Active` |
| **Conflict_Queue** | กักเก็บข้อมูลที่มีความขัดแย้ง | `Queue_ID`, `Incoming_Name`, `Suggested_Entity_ID`, `Conflict_Reason`, `Action_Required` |
| **NameMapping** | เก็บชื่อแปรผัน (Alias) | `Alias_Name`, `Target_Entity_ID`, `Province_Hint` |
| **Database** | View ภาพรวม (Report) | `Record_ID`, `Entity_ID`, `Location_ID`, `Display_Name`, `Address`, `Lat`, `Lng` |

### 2.2 ชีตปฏิบัติการ (Part 1: Daily Ops)
| ชื่อชีต | หน้าที่ | การเปลี่ยนแปลงใน V5 |
| :--- | :--- | :--- |
| **Input** | รับค่า Cookie/Shipment No | ไม่เปลี่ยน |
| **Data** | ข้อมูลรายวันจาก API | เพิ่มคอลัมน์ `Matched_Entity_ID`, `Matched_Location_ID` |
| **สรุป_*** | รายงานสรุป | Logic การคำนวณอ้างอิงจากข้อมูลที่ Match แล้วเท่านั้น |

### 2.3 ชีตแหล่งข้อมูลดิบ
| ชื่อชีต | หน้าที่ |
| :--- | :--- |
| **SCGนครหลวงJWDภูมิภาค** | แหล่งข้อมูลดิบสำหรับ Part 2 (มีคอลัมน์ `SYNC_STATUS`) |
| **PostalRef** | อ้างอิงรหัสไปรษณีย์และเขตตำบล |

## 3. ตรรกะการทำงานหลัก (Core Logic)

### 3.1 กระบวนการ Ingestion (ส่วนที่ 2)
1. **Read:** อ่านแถวจาก `SCGนครหลวงJWDภูมิภาค` ที่ยังไม่มีสถานะ `SYNCED`
2. **Normalize:** ทำความสะอาดชื่อ (ตัดคำฟุ่มเฟือย, ตัวพิมพ์เล็ก)
3. **Resolve Location:** 
   - ตรวจสอบพิกัดเดิมในระบบ (ระยะห่าง < 200 ม.) -> ใช้ `Location_ID` เดิม
   - ไม่เจอ -> สร้าง `Location_ID` ใหม่
4. **Resolve Entity:** 
   - ตรวจสอบชื่อตรง หรือ Alias -> ได้ `Entity_ID` เดิม
   - ไม่เจอ -> สร้าง `Entity_ID` ใหม่
5. **Check Relation:** 
   - ถ้า Entity นี้ เคยแมปกับ Location นี้แล้ว -> จบ (Matched)
   - ถ้าไม่เคย -> ตรวจสอบความขัดแย้ง (Conflict Detection)
     - *กรณีผิดปกติ:* เช่น ชื่อเดียวกันแต่พิกัดห่างกัน > 1 กม. -> ส่งเข้า `Conflict_Queue`
     - *ปกติ:* สร้างความสัมพันธ์ใหม่ใน `Entity_Loc_Map` (Auto-Link)
6. **Update Status:** บันทึกสถานะลงชีตดิบ

### 3.2 กระบวนการ Daily Match (ส่วนที่ 1)
1. **Read:** อ่านรายชื่อ `ShipToName` และ `LatLong_SCG` จากชีต `Data`
2. **Query:** ค้นหาใน `Entities` และ `Entity_Loc_Map` โดยใช้ชื่อและพิกัดเป็นตัวค้น
3. **Fetch:** ดึง `Lat`, `Lng` ที่ถูกต้องที่สุดจาก `Locations`
4. **Write:** เขียนลง `LatLong_Actual`, `Email พนักงาน`, และ ID ที่จับคู่ได้

### 3.3 การจัดการความขัดแย้ง (Conflict Resolution)
- ระบบจะไม่นำข้อมูลที่มีความเสี่ยงสูงเข้าฐานข้อมูลหลักโดยอัตโนมัติ
- Admin ต้องตรวจสอบผ่าน UI (`V5_ShowConflictQueueUI`) เพื่อเลือก:
  - **Approve:** ยืนยันการสร้างใหม่ หรือ การเชื่อมโยง
  - **Reject/Ignore:** ยกเลิกแถวนั้น หรือ แก้ไขข้อมูลก่อนบันทึก

## 4. การแก้ปัญหา 8 ข้อ (Solution Matrix)
| ปัญหา | กลไกแก้ไขใน V5 |
| :--- | :--- |
| 1. ชื่อบุคคลซ้ำ | แยก `Entity_ID` คนละตัว ใช้บริบท (ที่อยู่/จังหวัด) ช่วยตัดสิน |
| 2. ชื่อสถานที่ซ้ำ | ใช้ `Location_ID` กลาง หลาย Entity มาแชร์ที่เดียวกันได้ |
| 3. LatLong ซ้ำ | สร้าง `Location_ID` เดียว รองรับหลาย Entity |
| 4. ชื่อเขียนต่างกัน | ใช้ `NameMapping` ชี้ไป `Entity_ID` เดียวกัน |
| 5. คนละชื่อ ที่เดียวกัน | แยก `Entity` แต่ชี้ไปที่ `Location_ID` เดียวกัน |
| 6. ชื่อเดียวกัน ที่ต่างที่ | `Entity` เดียวกัน แต่มีหลาย Row ใน `Entity_Loc_Map` (Branch) |
| 7. ชื่อเดียวกัน พิกัดต่างไกล | ระบบตรวจจับระยะทาง -> ส่งเข้า `Conflict_Queue` |
| 8. คนละชื่อ พิกัดเดียวกัน | ระบบแจ้งเตือน -> Admin ตรวจสอบว่าเป็นคนละหน่วยงานจริงหรือไม่ |

## 5. เทคโนโลยีที่ใช้
- **Platform:** Google Apps Script + Google Sheets
- **Logic:** Deterministic Rules (Haversine Formula, String Normalization)
- **UI:** HTML Service (Dialog/Sidebar) สำหรับ Conflict Queue
- **No AI:** ไม่มีการใช้ Machine Learning เพื่อความแม่นยำที่ตรวจสอบได้
