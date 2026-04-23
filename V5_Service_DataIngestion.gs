/**
 * V5_Service_DataIngestion.gs
 * Version: 5.0
 * Description: Ingests raw data from 'SCGนครหลวงJWDภูมิภาค' sheet, processes each row through the Master Core,
 *              and updates the sync status in the raw sheet.
 */

function V5_IngestRawDataToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(V5_CONFIG.SHEETS.RAW_SCG);
  
  if (!rawSheet) {
    Browser.msgBox("❌ ไม่พบชีตข้อมูลดิบ 'SCGนครหลวงJWDภูมิภาค'");
    return;
  }

  // ดึงข้อมูลทั้งหมด (สมมติว่าแถวที่ 1 เป็น Header)
  // หมายเหตุ: ควรปรับช่วงการดึงข้อมูลให้เหมาะสมหากข้อมูลมีจำนวนมากมากๆ (ใช้ getValues แบบแบ่ง Chunk)
  const data = rawSheet.getDataRange().getValues();
  const headers = data[0];
  
  // ค้นหา Index ของคอลัมน์สำคัญในชีตดิบ (แบบ Dynamic ไม่ต้อง Hardcode เลขคอลัมน์)
  const colIndex = {
    name: headers.indexOf("ชื่อ - นามสกุล") + 1,
    lat: headers.indexOf("LAT") + 1,
    lng: headers.indexOf("LONG") + 1,
    address: headers.indexOf("ที่อยู่ปลายทาง") + 1,
    province: headers.indexOf("จังหวัด") > -1 ? headers.indexOf("จังหวัด") + 1 : headers.indexOf("คลังสินค้า...") + 1, // Fallback
    syncStatus: headers.indexOf("SYNC_STATUS") + 1
  };

  if (colIndex.name === 0 || colIndex.lat === 0 || colIndex.lng === 0) {
    Browser.msgBox("❌ ไม่พบคอลัมน์สำคัญ (ชื่อ, LAT, LONG) ในชีตข้อมูลดิบ กรุณาตรวจสอบชื่อคอลัมน์");
    return;
  }

  let processedCount = 0;
  let newEntityCount = 0;
  let newLocationCount = 0;
  let conflictCount = 0;
  let errorCount = 0;
  let skippedCount = 0;

  // แถวเริ่มต้น (ข้าม Header)
  const startRow = 2; 
  
  // ใช้ Array สำหรับรวบรวมผลลัพธ์เพื่อเขียนลง Sheet ครั้งเดียว (Batch Write) เพื่อความเร็ว
  // เราจะอัปเดตเฉพาะคอลัมน์ SYNC_STATUS และอาจเพิ่มคอลัมน์ Entity_ID/Location_ID หากต้องการ
  const statusUpdates = []; 

  Logger.log(`เริ่มประมวลผลข้อมูลดิบตั้งแต่แถวที่ ${startRow} ถึง ${data.length}`);

  for (let i = startRow; i <= data.length; i++) {
    const row = data[i-1]; // Array index starts at 0
    
    // ตรวจสอบว่าแถวนี้อัปเดทไปแล้วหรือยัง (ถ้ามีค่าใน SYNC_STATUS แล้วข้ามได้ หรือประมวลผลใหม่ก็ได้ตาม Logic)
    const currentSyncStatus = row[colIndex.syncStatus - 1];
    
    // ตัวอย่าง Logic: ถ้ามีคำว่า "SYNCED" อยู่แล้ว ให้ข้าม (สามารถเปลี่ยนเป็นประมวลผลใหม่ได้หากต้องการ)
    if (currentSyncStatus && currentSyncStatus.toString().includes("SYNCED")) {
      skippedCount++;
      continue;
    }

    const rawName = row[colIndex.name - 1];
    const lat = parseFloat(row[colIndex.lat - 1]);
    const lng = parseFloat(row[colIndex.lng - 1]);
    const address = row[colIndex.address - 1] || "";
    const province = row[colIndex.province - 1] || "";

    // ข้ามแถวที่ไม่มีข้อมูลสำคัญ
    if (!rawName || isNaN(lat) || isNaN(lng)) {
      statusUpdates.push({ row: i, status: "ERROR_MISSING_DATA" });
      errorCount++;
      continue;
    }

    // --- เรียกใช้สมอง (Master Core) ---
    const result = V5_ResolveIdentity(rawName, lat, lng, address, province);
    
    let finalStatus = "";
    
    if (result.status === "MATCHED_COMPLETE" || result.status === "AUTO_LINKED") {
      processedCount++;
      if (result.message.includes("NEW_ENTITY")) newEntityCount++;
      if (result.message.includes("NEW_LOCATION")) newLocationCount++;
      finalStatus = `SYNCED: ${result.status}`;
    } else if (result.status === "CONFLICT_PENDING") {
      conflictCount++;
      finalStatus = `CONFLICT: ${result.message}`;
    } else if (result.status === "ERROR_INVALID_DATA") {
      errorCount++;
      finalStatus = "ERROR: Invalid Data";
    } else {
      finalStatus = `PROCESSED: ${result.status}`;
    }

    statusUpdates.push({ row: i, status: finalStatus });
  }

  // --- เขียนผลลัพธ์กลับลงชีตดิบ (Batch Update) ---
  if (statusUpdates.length > 0) {
    const statusRange = rawSheet.getRange(startRow, colIndex.syncStatus, statusUpdates.length, 1);
    const statusValues = statusUpdates.map(u => [u.status]);
    statusRange.setValues(statusValues);
  }

  // --- สรุปผล ---
  const summary = `
✅ เสร็จสิ้นการนำเข้าข้อมูล!
---------------------------
📊 ประมวลผลทั้งหมด: ${processedCount + conflictCount + errorCount} แถว
🔄 ข้าม (เดิมมีอยู่แล้ว): ${skippedCount} แถว
✨ สร้าง Entity ใหม่: ${newEntityCount} รายการ
📍 สร้าง Location ใหม่: ${newLocationCount} รายการ
⚠️ ติดขัด (รอตรวจสอบ): ${conflictCount} รายการ
❌ ผิดพลาด: ${errorCount} รายการ

กรุณาตรวจสอบชีต 'Conflict_Queue' หากมีรายการติดขัด
  `;
  
  Logger.log(summary);
  Browser.msgBox(summary);
}

/**
 * ฟังก์ชันเสริม: รีเซ็ตสถานะ Sync ทั้งหมด (สำหรับการทดสอบหรือเริ่มใหม่)
 */
function V5_ResetSyncStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(V5_CONFIG.SHEETS.RAW_SCG);
  
  if (!rawSheet) return;
  
  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf("SYNC_STATUS") + 1;
  
  if (colIndex === 0) {
    Browser.msgBox("ไม่พบคอลัมน์ SYNC_STATUS");
    return;
  }
  
  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) return;
  
  if (Browser.msgBox("คุณต้องการล้างสถานะ SYNC ทั้งหมดหรือไม่?\n(ข้อมูลใน Entities/Locations จะไม่ถูกลบ)", Browser.Buttons.YES_NO) === Browser.Buttons.YES) {
    rawSheet.getRange(2, colIndex, lastRow - 1, 1).clearContent();
    Browser.msgBox("✅ ล้างสถานะ Sync เรียบร้อยแล้ว");
  }
}
