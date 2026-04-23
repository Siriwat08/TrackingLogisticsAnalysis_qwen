/**
 * ==============================================================================
 * PROJECT: LMDS (Logistics Master Data System)
 * VERSION: 5.0 (Complete Rebuild)
 * FILE: V5_Config.gs
 * DESCRIPTION: Central Configuration Hub. Defines all Sheet Names, Column Indices, 
 *              and System Constants for both Part 1 (Daily Ops) and Part 2 (Master Data).
 * ==============================================================================
 */

const V5_CONFIG = {
  // ---------------------------------------------------------------------------
  // 1. SHEET NAMES DEFINITION
  // ---------------------------------------------------------------------------
  SHEETS: {
    // --- PART 2: MASTER DATA ENGINE (ข้อมูลดิบและฐานข้อมูลหลัก) ---
    RAW_SOURCE: "SCGนครหลวงJWDภูมิภาค",   // แหล่งข้อมูลดิบ
    POSTAL_REF: "PostalRef",             // ข้อมูลไปรษณีย์
    
    ENTITIES: "Entities",                // เก็บตัวตน (ใคร)
    LOCATIONS: "Locations",              // เก็บพิกัด/ที่อยู่ (ที่ไหน)
    ENTITY_LOC_MAP: "Entity_Loc_Map",    // ความสัมพันธ์ (ใคร อยู่ ที่ไหน)
    CONFLICT_QUEUE: "Conflict_Queue",    // คิวจัดการข้อขัดแย้ง
    NAME_MAPPING: "NameMapping",         // ชื่อแปรผัน/Alias
    DATABASE_VIEW: "Database",           // View ภาพรวมสำหรับรายงาน
    
    // --- PART 1: DAILY OPERATIONS (ปฏิบัติงานรายวัน) ---
    INPUT: "Input",                      // รับค่า Cookie/Shipment
    DATA: "Data",                        // ข้อมูลงานรายวันจาก API
    EMPLOYEE: "ข้อมูลพนักงาน",            // ข้อมูลพนักงานและอีเมล
    SUMMARY_SHIPMENT: "สรุป_Shipment",   // สรุปตาม Shipment
    SUMMARY_OWNER: "สรุป_เจ้าของสินค้า"   // สรุปตามเจ้าของสินค้า
  },

  // ---------------------------------------------------------------------------
  // 2. COLUMN INDICES (1-Based Index)
  // กำหนดลำดับคอลัมน์ของทุกชีตเพื่อให้โค้ดอื่นอ้างอิงได้ถูกต้อง
  // ---------------------------------------------------------------------------
  COL: {
    // --- ENTITIES (ชีต: Entities) ---
    ENTITIES: {
      ID: 1,               // Entity_ID (UUID)
      DISPLAY_NAME: 2,     // Display_Name
      NORMALIZED_NAME: 3,  // Normalized_Name
      TYPE: 4,             // Entity_Type (PERSON, SHOP, COMPANY)
      PHONE: 5,            // Phone
      TAX_ID: 6,           // Tax_ID
      STATUS: 7,           // Status (ACTIVE, MERGED, DELETED)
      MERGED_TO: 8,        // Merged_To_Entity_ID
      CREATED_AT: 9,       // Created_At
      UPDATED_AT: 10       // Updated_At
    },

    // --- LOCATIONS (ชีต: Locations) ---
    LOCATIONS: {
      ID: 1,               // Location_ID (UUID)
      LAT: 2,              // Latitude
      LNG: 3,              // Longitude
      LATLONG_KEY: 4,      // LatLong_Key (สำหรับการค้นหาเร็ว)
      FULL_ADDRESS: 5,     // Full_Address
      PROVINCE: 6,         // Province
      DISTRICT: 7,         // District
      SUB_DISTRICT: 8,     // SubDistrict
      POSTAL_CODE: 9,      // Postal_Code
      TYPE: 10,            // Location_Type
      CONFIDENCE: 11,      // Confidence_Score
      SOURCE: 12,          // Source (SCG_API, MANUAL, GPS)
      LAST_VERIFIED: 13    // Last_Verified
    },

    // --- ENTITY_LOC_MAP (ชีต: Entity_Loc_Map) ---
    MAP: {
      ID: 1,               // Map_ID (UUID)
      ENTITY_ID: 2,        // Entity_ID
      LOCATION_ID: 3,      // Location_ID
      RELATION_TYPE: 4,    // Relation_Type (PRIMARY, BRANCH, TEMP)
      IS_ACTIVE: 5,        // Is_Active (TRUE/FALSE)
      CONFIDENCE: 6,       // Confidence
      NOTES: 7,            // Notes
      CREATED_AT: 8        // Created_At
    },

    // --- CONFLICT_QUEUE (ชีต: Conflict_Queue) ---
    QUEUE: {
      ID: 1,               // Queue_ID
      RECEIVED_AT: 2,      // Received_At
      INCOMING_NAME: 3,    // Incoming_Name
      INCOMING_LATLNG: 4,  // Incoming_LatLong
      INCOMING_ADDR: 5,    // Incoming_Address
      SUGGESTED_ENT: 6,    // Suggested_Entity_ID
      SUGGESTED_LOC: 7,    // Suggested_Location_ID
      REASON: 8,           // Conflict_Reason
      ACTION_REQ: 9,       // Action_Required (REVIEW, MERGE, SPLIT)
      RESOLVED_BY: 10,     // Resolved_By
      RESOLVED_AT: 11,     // Resolved_At
      NOTE: 12             // Resolution_Note
    },

    // --- NAME_MAPPING (ชีต: NameMapping) ---
    MAPPING: {
      ALIAS: 1,            // Alias_Name
      TARGET_ENT_ID: 2,    // Target_Entity_ID
      PROV_HINT: 3,        // Province_Hint
      CONFIDENCE: 4,       // Confidence
      MAPPED_BY: 5,        // Mapped_By (AI, MANUAL, AUTO)
      USAGE_COUNT: 6,      // Usage_Count
      LAST_USED: 7         // Last_Used
    },

    // --- DATABASE VIEW (ชีต: Database) ---
    DATABASE: {
      RECORD_ID: 1,        // Record_ID
      ENTITY_ID: 2,        // Entity_ID
      LOCATION_ID: 3,      // Location_ID
      DISPLAY_NAME: 4,     // Display_Name
      FULL_ADDRESS: 5,     // Full_Address
      LAT: 6,              // Latitude
      LNG: 7,              // Longitude
      PROVINCE: 8,         // Province
      STATUS: 9,           // Status
      LAST_UPDATED: 10     // Last_Updated
    },

    // --- DATA SHEET (ชีต: Data - ส่วนที่ 1) ---
    // หมายเหตุ: ดัชนีเหล่านี้ต้องตรงกับโครงสร้าง 29 คอลัมน์เดิม + คอลัมน์ใหม่
    DATA: {
      ID_JOB: 1,
      PLAN_DELIVERY: 2,
      INVOICE_NO: 3,
      SHIPMENT_NO: 4,
      DRIVER_NAME: 5,
      TRUCK_LICENSE: 6,
      CARRIER_CODE: 7,
      CARRIER_NAME: 8,
      SOLD_TO_CODE: 9,
      SOLD_TO_NAME: 10,
      SHIP_TO_NAME: 11,    // *สำคัญ: ใช้จับคู่*
      SHIP_TO_ADDR: 12,
      LATLONG_SCG: 13,     // *สำคัญ: พิกัดต้นทาง*
      MATERIAL_NAME: 14,
      ITEM_QTY: 15,
      QTY_UNIT: 16,
      ITEM_WEIGHT: 17,
      DELIVERY_NO: 18,
      DEST_COUNT_SYS: 19,
      DEST_LIST_SYS: 20,
      SCAN_STATUS: 21,
      DELIVERY_STATUS: 22,
      EMP_EMAIL: 23,       // จะถูกเติมโดยระบบ
      TOTAL_QTY: 24,
      TOTAL_WEIGHT: 25,
      INV_COUNT: 26,
      LATLONG_ACTUAL: 27,  // *สำคัญ: พิกัดจริงที่ได้จาก Part 2*
      OWNER_NAME: 28,
      SHOP_KEY: 29,
      // คอลัมน์เพิ่มเติมสำหรับ Part 2 (จะเพิ่มท้ายแถว)
      MATCHED_ENT_ID: 30,  // Matched_Entity_ID
      MATCHED_LOC_ID: 31,  // Matched_Location_ID
      MATCH_CONFIDENCE: 32 // Match_Confidence
    },

    // --- RAW SOURCE (ชีต: SCGนครหลวงJWDภูมิภาค) ---
    // สมมติโครงสร้างพื้นฐาน (อาจปรับแก้ตามข้อมูลจริงในอนาคต)
    RAW: {
      ID: 1,
      DATE: 2,
      TIME: 3,
      DEST_NAME: 4,      // ชื่อปลายทาง
      DRIVER_NAME: 5,
      TRUCK: 6,
      SHIPMENT_NO: 7,
      INVOICE_NO: 8,
      CUSTOMER_CODE: 9,
      OWNER_NAME: 10,
      DEST_NAME_FULL: 11,// ชื่อปลายทางเต็ม
      EMAIL: 12,
      LAT: 13,           // LAT
      LNG: 14,           // LONG
      ADDRESS: 15,       // ที่อยู่
      SYNC_STATUS: 16,   // สถานะการ Sync
      // ... คอลัมน์อื่นๆ สามารถเพิ่มได้ตามต้องการ
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
