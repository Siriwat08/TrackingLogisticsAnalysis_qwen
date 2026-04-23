/**
 * V5_Setup.gs
 * Version: 5.0
 * Description: สร้างโครงสร้างชีตทั้งหมด (13 ชีต), กำหนด Headers, และเตรียมความพร้อมของระบบ
 */

function V5_InitAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log("🚀 เริ่มการสร้างโครงสร้างฐานข้อมูล V5...");
    
    // --- ส่วนที่ 2: Master Data Engine (สร้างใหม่ทั้งหมด) ---
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.ENTITIES, HEADERS.ENTITIES);
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.LOCATIONS, HEADERS.LOCATIONS);
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.MAP, HEADERS.MAP);
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.QUEUE, HEADERS.QUEUE);
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.MAPPING, HEADERS.MAPPING);
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.DATABASE_VIEW, HEADERS.DATABASE_VIEW);
    
    // --- ส่วนที่ 1: Daily Operations (ตรวจสอบและปรับปรุง) ---
    // Input
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.INPUT, HEADERS.INPUT);
    
    // Data (สำคัญ: ต้องเพิ่มคอลัมน์เชื่อมโยงใหม่)
    ensureSheetColumns(ss, V5_CONFIG.SHEETS.DATA, HEADERS.DATA);
    
    // ข้อมูลพนักงาน
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.EMPLOYEE, HEADERS.EMPLOYEE);
    
    // สรุปต่างๆ
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.SUMMARY_SHIPMENT, HEADERS.SUMMARY_SHIPMENT);
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.SUMMARY_OWNER, HEADERS.SUMMARY_OWNER);
    
    // แหล่งข้อมูลดิบ
    ensureSheetColumns(ss, V5_CONFIG.SHEETS.RAW_SCG, HEADERS.RAW_SCG);
    
    // PostalRef
    createSheetWithHeaders(ss, V5_CONFIG.SHEETS.POSTAL, HEADERS.POSTAL);
    
    // จัดรูปแบบเบื้องต้น (Format & Freeze)
    applyBasicFormatting(ss);
    
    Logger.log("✅ เสร็จสิ้นการสร้างโครงสร้าง!");
    ui.alert("✅ สำเร็จ!\n\nสร้าง/อัปเดต ครบทั้ง 13 ชีตเรียบร้อยแล้ว\nระบบพร้อมสำหรับการนำเข้าข้อมูลดิบ (Part 2) และการปฏิบัติงานรายวัน (Part 1)");
    
  } catch (e) {
    Logger.log("❌ เกิดข้อผิดพลาด: " + e.toString());
    ui.alert("❌ เกิดข้อผิดพลาดในการสร้างชีต:\n" + e.toString());
  }
}

/**
 * Helper: สร้างชีตใหม่พร้อม Headers ถ้ายังไม่มี
 */
function createSheetWithHeaders(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`✨ สร้างชีตใหม่: ${sheetName}`);
  } else {
    Logger.log(`ℹ️ ชีตมีอยู่แล้ว: ${sheetName} (ข้ามการสร้าง)`);
    // ถ้ามีแล้ว แต่อาจต้องการอัปเดต Header (ถ้ามีการเปลี่ยนแปลงในอนาคต)
    // ปัจจุบันจะตรวจสอบแค่ความยาว ถ้าไม่ตรงจะเตือน
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (currentHeaders.length !== headers.length) {
      Logger.log(`⚠️ คำเตือน: จำนวนคอลัมน์ของ ${sheetName} ไม่ตรงกัน (เดิม: ${currentHeaders.length}, ใหม่: ${headers.length})`);
    }
    return; 
  }
  
  // เขียน Headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // จัดรูปแบบ Header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  
  // ตรึงแถวแรก
  sheet.setFrozenRows(1);
  
  // ปรับความกว้างคอลัมน์อัตโนมัติ (จำกัดที่ 15-30 เพื่อความสวยงาม)
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Helper: ตรวจสอบชีตที่มีอยู่ และเพิ่มคอลัมน์ที่ขาดหาย (สำหรับชีตเดิมเช่น Data, SCG)
 */
function ensureSheetColumns(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // ถ้าไม่มีเลย ให้สร้างใหม่ (กรณีชีตเดิมถูกลบผิด)
    createSheetWithHeaders(ss, sheetName, expectedHeaders);
    return;
  }
  
  // ตรวจสอบ Header แถวที่ 1
  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  
  // เปรียบเทียบและเติมส่วนที่ขาด
  // หมายเหตุ: วิธีนี้สมมติว่าลำดับคอลัมน์ต้องตรงกันเป๊ะตาม Config
  // หากต้องการความยืดหยุ่นสูงมาก อาจต้องเช็คชื่อทีละคอลัมน์ แต่ในที่นี้ใช้วิธีเขียนทับหรือเติมให้ครบตามจำนวน
  if (currentHeaders.length < expectedHeaders.length) {
    Logger.log(`🔧 อัปเดตชีต ${sheetName}: เพิ่มคอลัมน์จาก ${currentHeaders.length} เป็น ${expectedHeaders.length}`);
    
    // เขียน Headers ที่ขาดลงไป (ต่อท้าย)
    const missingCount = expectedHeaders.length - currentHeaders.length;
    const missingHeaders = expectedHeaders.slice(currentHeaders.length);
    
    if (missingCount > 0) {
      sheet.getRange(1, currentHeaders.length + 1, 1, missingCount).setValues([missingHeaders]);
      const newHeaderRange = sheet.getRange(1, currentHeaders.length + 1, 1, missingCount);
      newHeaderRange.setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
      sheet.autoResizeColumns(currentHeaders.length + 1, missingCount);
    }
  } else if (currentHeaders.length > expectedHeaders.length) {
     Logger.log(`⚠️ ชีต ${sheetName} มีคอลัมน์มากกว่าที่กำหนด (${currentHeaders.length} vs ${expectedHeaders.length}) ไม่ทำการลบข้อมูล`);
  } else {
    Logger.log(`✅ ชีต ${sheetName} โครงสร้างถูกต้อง`);
  }
  
  // ตรึงแถวแรกเสมอ
  sheet.setFrozenRows(1);
}

/**
 * จัดรูปแบบพื้นฐานทั่วทั้งสเปรดชีต
 */
function applyBasicFormatting(ss) {
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    // ซ่อน Gridlines เพื่อให้ดูสะอาดตาขึ้น (Optional)
    // sheet.setHiddenGridlines(true); 
    
    // ฟอนต์มาตรฐาน
    const range = sheet.getRange("A1:Z1000"); // ประมาณการ
    range.setFontFamily("Sarabun"); // หรือฟอนต์ที่รองรับภาษาไทยดี ๆ
    range.setFontSize(10);
  });
}

// ============================================================================
// รายการ Headers ทั้งหมด (อ้างอิงจาก V5_Config และความต้องการของผู้ใช้)
// ============================================================================

const HEADERS = {
  // --- Part 2: Master Data ---
  ENTITIES: [
    "Entity_ID", "Display_Name", "Normalized_Name", "Entity_Type", 
    "Phone", "Tax_ID", "Status", "Merged_To_Entity_ID", "Created_At", "Updated_At"
  ],
  LOCATIONS: [
    "Location_ID", "Latitude", "Longitude", "LatLong_Key", 
    "Full_Address", "Province", "District", "SubDistrict", "Postal_Code", 
    "Location_Type", "Confidence_Score", "Source", "Last_Verified"
  ],
  MAP: [
    "Map_ID", "Entity_ID", "Location_ID", "Relation_Type", 
    "Is_Active", "Confidence", "Notes", "Created_At"
  ],
  QUEUE: [
    "Queue_ID", "Received_At", "Incoming_Name", "Incoming_LatLong", 
    "Incoming_Address", "Suggested_Entity_ID", "Suggested_Location_ID", 
    "Conflict_Reason", "Action_Required", "Resolved_By", "Resolved_At", "Resolution_Note"
  ],
  MAPPING: [
    "Alias_Name", "Target_Entity_ID", "Province_Hint", "Confidence", 
    "Mapped_By", "Usage_Count", "Last_Used"
  ],
  DATABASE_VIEW: [
    "Record_ID", "Entity_ID", "Location_ID", "Display_Name", 
    "Full_Address", "Latitude", "Longitude", "Province", "Status", "Last_Updated"
  ],

  // --- Part 1: Daily Ops (ตามที่ระบุ) ---
  INPUT: [
    "COOKIE", "ShipmentNos" 
    // หมายเหตุ: ตามโจทย์ Cookie อยู่ B1, Shipment เริ่ม A4 แต่เราทำ Header ไว้แถว 1 เพื่อความเป็นระเบียบ
    // หรือถ้าต้องการตามเดิมเป๊ะๆ (ไม่มี Header) สามารถปรับได้ แต่แนะนำให้มี Header
  ],
  DATA: [
    "ID_งานประจำวัน", "PlanDelivery", "InvoiceNo", "ShipmentNo", 
    "DriverName", "TruckLicense", "CarrierCode", "CarrierName", 
    "SoldToCode", "SoldToName", "ShipToName", "ShipToAddress", 
    "LatLong_SCG", "MaterialName", "ItemQuantity", "QuantityUnit", 
    "ItemWeight", "DeliveryNo", "จำนวนปลายทาง_System", "รายชื่อปลายทาง_System", 
    "ScanStatus", "DeliveryStatus", "Email พนักงาน", 
    "จำนวนสินค้ารวมของร้านนี้", "น้ำหนักสินค้ารวมของร้านนี้", 
    "จำนวน_Invoice_ที่ต้องสแกน", "LatLong_Actual", 
    "ชื่อเจ้าของสินค้า_Invoice_ที่ต้องสแกน", "ShopKey",
    // คอลัมน์ใหม่สำหรับเชื่อม Part 2
    "Matched_Entity_ID", "Matched_Location_ID", "Match_Confidence"
  ],
  EMPLOYEE: [
    "ID_พนักงาน", "ชื่อ - นามสกุล", "เบอร์โทรศัพท์", "เลขที่บัตรประชาชน", 
    "ทะเบียนรถ", "เลือกประเภทรถยนต์", "Email พนักงาน", "ROLE"
  ],
  SUMMARY_SHIPMENT: [
    "ShipmentKey", "ShipmentNo", "TruckLicense", "PlanDelivery", 
    "จำนวน_ทั้งหมด", "จำนวน_E-POD_ทั้งหมด", "LastUpdated"
  ],
  SUMMARY_OWNER: [
    "SummaryKey", "SoldToName", "PlanDelivery", 
    "จำนวน_ทั้งหมด", "จำนวน_E-POD_ทั้งหมด", "LastUpdated"
  ],
  
  // --- Raw Data & Ref ---
  RAW_SCG: [
    "head", "ID_SCGนครหลวงJWDภูมิภาค", "วันที่ส่งสินค้า", "เวลาที่ส่งสินค้า", 
    "จุดส่งสินค้าปลายทาง", "ชื่อ - นามสกุล", "ทะเบียนรถ", "Shipment No", 
    "Invoice No", "รูปถ่ายบิลส่งสินค้า", "รหัสลูกค้า", "ชื่อเจ้าของสินค้า", 
    "ชื่อปลายทาง", "Email พนักงาน", "LAT", "LONG", "ID_Doc_Return", 
    "คลังสินค้า...", "ที่อยู่ปลายทาง", "รูปสินค้าตอนส่ง", "รูปหน้าร้าน / บ้าน", 
    "หมายเหตุ", "เดือน", "ระยะทางจากคลัง_Km", "ชื่อที่อยู่จาก_LatLong", 
    "SM_Link_SCG", "ID_พนักงาน", "พิกัดตอนกดบันทึกงาน", "เวลาเริ่มกรอกงาน", 
    "เวลาบันทึกงานสำเร็จ", "ระยะขยับจากจุดเริ่มต้น_เมตร", "ระยะเวลาใช้งาน_นาที", 
    "ความเร็วการเคลื่อนที่_เมตร_นาที", "ผลการตรวจสอบงานส่ง", "เหตุผิดปกติที่ตรวจพบ", 
    "เวลาถ่ายรูปหน้าร้าน_หน้าบ้าน", "SYNC_STATUS"
    // อาจเพิ่มคอลัมน์สถานะการประมวลผล part 2 ในอนาคต เช่น "PART2_PROCESSED"
  ],
  POSTAL: [
    "postcode", "subdistrict", "district", "province", 
    "province_code", "district_code", "lat", "lng", "notes"
  ]
};
