/**
 * V5_Config.gs
 * Version: 5.0
 * Description: ศูนย์กลางการตั้งค่า (Central Configuration) กำหนดชื่อชีต ดัชนีคอลัมน์ และค่าคงที่ของระบบ
 *              อัปเดตตามโครงสร้างจริงของผู้ใช้งาน (Part 1 & Part 2)
 */

const V5_CONFIG = {
  // ---------------------------------------------------------------------------
  // 1. ชื่อชีตทั้งหมด (Sheet Names)
  // ---------------------------------------------------------------------------
  SHEETS: {
    // --- ส่วนที่ 2: Master Data Engine & Raw Data ---
    ENTITIES: "Entities",               // ชีตใหม่: เก็บตัวตน (ใคร)
    LOCATIONS: "Locations",             // ชีตใหม่: เก็บพิกัด (ที่ไหน)
    MAP: "Entity_Loc_Map",              // ชีตใหม่: ความสัมพันธ์ (ใคร-อยู่ที่ไหน)
    QUEUE: "Conflict_Queue",            // ชีตใหม่: จัดการข้อขัดแย้ง
    MAPPING: "NameMapping",             // ชีตเดิม: เก็บชื่อแปรผัน (Alias)
    DATABASE_VIEW: "Database",          // ชีตเดิม: รายงานภาพรวม (View)
    
    RAW_SCG: "SCGนครหลวงJWDภูมิภาค",     // แหล่งข้อมูลดิบหลัก
    POSTAL: "PostalRef",                // ข้อมูลไปรษณีย์
    
    // --- ส่วนที่ 1: Daily Operations ---
    INPUT: "Input",                     // รับค่า Cookie/Shipment
    DATA: "Data",                       // ผลลัพธ์การทำงานรายวัน
    EMPLOYEE: "ข้อมูลพนักงาน",            // ข้อมูลพนักงาน
    SUMMARY_SHIPMENT: "สรุป_Shipment",   // สรุปตาม Shipment
    SUMMARY_OWNER: "สรุป_เจ้าของสินค้า"  // สรุปตามเจ้าของสินค้า
  },

  // ---------------------------------------------------------------------------
  // 2. ดัชนีคอลัมน์ (Column Indices) - เริ่มนับจาก 1
  // ---------------------------------------------------------------------------
  COL: {
    // --- ชีต Entities (ใหม่) ---
    ENTITIES: {
      ID: 1, DISPLAY_NAME: 2, NORM_NAME: 3, TYPE: 4, PHONE: 5, TAX_ID: 6, 
      STATUS: 7, MERGED_TO: 8, CREATED_AT: 9, UPDATED_AT: 10
    },
    
    // --- ชีต Locations (ใหม่) ---
    LOCATIONS: {
      ID: 1, LAT: 2, LNG: 3, KEY: 4, ADDR: 5, PROVINCE: 6, DISTRICT: 7, 
      SUBDISTRICT: 8, POSTAL_CODE: 9, TYPE: 10, CONFIDENCE: 11, SOURCE: 12, LAST_VERIFIED: 13
    },
    
    // --- ชีต Entity_Loc_Map (ใหม่) ---
    MAP: {
      ID: 1, ENT_ID: 2, LOC_ID: 3, REL_TYPE: 4, IS_ACTIVE: 5, CONFIDENCE: 6, NOTES: 7, CREATED_AT: 8
    },
    
    // --- ชีต Conflict_Queue (ใหม่) ---
    QUEUE: {
      ID: 1, RECEIVED_AT: 2, IN_NAME: 3, IN_LATLNG: 4, IN_ADDR: 5, 
      SUG_ENT_ID: 6, SUG_LOC_ID: 7, REASON: 8, ACTION: 9, RESOLVED_BY: 10, RESOLVED_AT: 11, NOTE: 12
    },
    
    // --- ชีต NameMapping (ปรับปรุง) ---
    MAPPING: {
      ALIAS: 1, TARGET_ENT_ID: 2, PROV_HINT: 3, CONFIDENCE: 4, MAPPED_BY: 5, USAGE: 6, LAST_USED: 7
    },
    
    // --- ชีต Database (View) ---
    DATABASE_VIEW: {
      REC_ID: 1, ENT_ID: 2, LOC_ID: 3, NAME: 4, ADDR: 5, LAT: 6, LNG: 7, PROV: 8, STATUS: 9, UPDATED: 10
    },

    // --- ชีต Input (ส่วนที่ 1) ---
    INPUT: {
      COOKIE: 2,      // B1
      SHIPMENT_LIST: 1 // A4:A
    },

    // --- ชีต Data (ส่วนที่ 1) - 29 คอลัมน์ ---
    DATA: {
      ID_DAILY: 1, PlanDelivery: 2, InvoiceNo: 3, ShipmentNo: 4, DriverName: 5, 
      TruckLicense: 6, CarrierCode: 7, CarrierName: 8, SoldToCode: 9, SoldToName: 10, 
      ShipToName: 11, ShipToAddress: 12, LatLong_SCG: 13, MaterialName: 14, ItemQuantity: 15, 
      QuantityUnit: 16, ItemWeight: 17, DeliveryNo: 18, Count_System: 19, List_System: 20, 
      ScanStatus: 21, DeliveryStatus: 22, Email_Emp: 23, Total_Qty: 24, Total_Wgt: 25, 
      Count_Invoice: 26, LatLong_Actual: 27, Owner_Name: 28, ShopKey: 29,
      // คอลัมน์เสริมสำหรับ Part 2 (จะเพิ่มท้ายแถวอัตโนมัติหากไม่มี)
      Matched_Ent_ID: 30, Matched_Loc_ID: 31, Match_Confidence: 32
    },

    // --- ชีต ข้อมูลพนักงาน ---
    EMPLOYEE: {
      ID: 1, NAME: 2, PHONE: 3, ID_CARD: 4, LICENSE_PLATE: 5, VEHICLE_TYPE: 6, EMAIL: 7, ROLE: 8
    },

    // --- ชีต สรุป_Shipment ---
    SUMMARY_SHIPMENT: {
      KEY: 1, ShipmentNo: 2, TruckLicense: 3, PlanDelivery: 4, Total_Count: 5, EPOD_Count: 6, LastUpdated: 7
    },

    // --- ชีต สรุป_เจ้าของสินค้า ---
    SUMMARY_OWNER: {
      KEY: 1, SoldToName: 2, PlanDelivery: 3, Total_Count: 4, EPOD_Count: 5, LastUpdated: 6
    },

    // --- ชีต SCGนครหลวงJWDภูมิภาค (Raw Data) - 41+ คอลัมน์ ---
    RAW_SCG: {
      HEAD: 1, ID_SCG: 2, Date: 3, Time: 4, Dest_Point: 5, Name_Full: 6, License_Plate: 7, 
      Shipment_No: 8, Invoice_No: 9, Photo_Bill: 10, Cust_Code: 11, Owner_Name: 12, 
      Dest_Name: 13, Email_Emp: 14, LAT: 15, LNG: 16, Doc_Return_ID: 17, Warehouse: 18, 
      Address_Full: 19, Photo_Goods: 20, Photo_Shop: 21, Note: 22, Month: 23, Dist_KM: 24, 
      Addr_From_LatLong: 25, SM_Link: 26, Emp_ID: 27, Click_LatLong: 28, Start_Time: 29, 
      End_Time: 30, Move_Dist_M: 31, Duration_Min: 32, Speed_M_Min: 33, Check_Result: 34, 
      Issue: 35, Photo_Time: 36, SYNC_STATUS: 37
      // หมายเหตุ: หากมีคอลัมน์เพิ่ม สามารถระบุเลขถัดไปได้เลย
    },

    // --- ชีต PostalRef ---
    POSTAL: {
      CODE: 1, SUBDISTRICT: 2, DISTRICT: 3, PROVINCE: 4, PROV_CODE: 5, DIST_CODE: 6, LAT: 7, LNG: 8, NOTES: 9
    }
  },

  // ---------------------------------------------------------------------------
  // 3. SYSTEM CONSTANTS & THRESHOLDS
  // ---------------------------------------------------------------------------
  SYSTEM: {
    // เกณฑ์ความมั่นใจในการจับคู่อัตโนมัติ (0-100)
    CONFIDENCE_AUTO_ACCEPT: 90, 
    
    // เกณฑ์ระยะทาง (เมตร) ในการถือว่าอยู่ในจุดเดียวกัน
    DISTANCE_THRESHOLD_METERS: 150, 
    
    // จำนวนทศนิยมของพิกัดสำหรับการสร้าง Key (6 ตำแหน่ง ≈ 10 ซม.)
    LAT LNG_DECIMAL_PLACES: 6,
    
    // ค่าเริ่มต้นสถานะ
    DEFAULT_STATUS: "ACTIVE",
    DEFAULT_RELATION: "PRIMARY",
    
    // การตั้งค่า AI (ถ้ามีใช้ในอนาคต)
    AI_MODEL: "gemini-pro",
    AI_TIMEOUT_MS: 30000
  },

  // ---------------------------------------------------------------------------
  // 4. HELPER METHODS (ภายใน Config)
  // ---------------------------------------------------------------------------
  
  /**
   * ตรวจสอบว่าชีตสำคัญมีครบหรือไม่
   * @returns {boolean} true ถ้าครบ, false ถ้าขาด
   */
  validateSystemIntegrity: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = Object.values(this.SHEETS);
    let missing = [];
    
    for (let sheetName of requiredSheets) {
      if (!ss.getSheetByName(sheetName)) {
        missing.push(sheetName);
      }
    }
    
    if (missing.length > 0) {
      Logger.log("⚠️ Missing Sheets: " + missing.join(", "));
      return false;
    }
    return true;
  }
};
