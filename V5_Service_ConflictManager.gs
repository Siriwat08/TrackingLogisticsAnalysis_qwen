/**
 * V5_Service_ConflictManager.gs
 * Version: 5.0
 * Description: Handles the Conflict Queue logic and serves data to the HTML UI for manual resolution.
 */

// --- Functions for HTML UI ---

/**
 * ดึงรายการทั้งหมดใน Conflict Queue เพื่อแสดงบน UI
 * @returns {Array} Array of objects representing conflict rows
 */
function V5_GetConflictQueueData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const result = [];

  // กรองเฉพาะแถวที่ยังไม่ได้รับการแก้ไข (Action_Required != '')
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[V5_CONFIG.COL.QUEUE.ACTION_REQUIRED - 1]; // Column J (index 9)
    
    if (status && status !== "" && status !== "RESOLVED") {
      result.push({
        rowIndex: i + 1, // เก็บเลขแถวจริงของ Google Sheets (เริ่มที่ 1)
        id: row[V5_CONFIG.COL.QUEUE.ID - 1],
        timestamp: row[V5_CONFIG.COL.QUEUE.RECEIVED_AT - 1],
        incomingName: row[V5_CONFIG.COL.QUEUE.INCOMING_NAME - 1],
        incomingLatLong: row[V5_CONFIG.COL.QUEUE.INCOMING_LATLNG - 1],
        incomingAddress: row[V5_CONFIG.COL.QUEUE.INCOMING_ADDRESS - 1],
        suggestedEntityId: row[V5_CONFIG.COL.QUEUE.SUGGESTED_ENTITY_ID - 1],
        suggestedLocationId: row[V5_CONFIG.COL.QUEUE.SUGGESTED_LOCATION_ID - 1],
        reason: row[V5_CONFIG.COL.QUEUE.CONFLICT_REASON - 1],
        action: status
      });
    }
  }
  return result;
}

/**
 * ประมวลผลการอนุมัติหรือปฏิเสธจาก UI
 * @param {number} rowIndex - แถวที่ต้องการแก้ไขในชีต
 * @param {string} action - 'APPROVE' หรือ 'REJECT'
 * @param {string} resolverName - ชื่อผู้แก้ไข
 */
function V5_ResolveConflictItem(rowIndex, action, resolverName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  
  if (!sheet) throw new Error("ไม่พบชีต Conflict_Queue");

  const range = sheet.getRange(rowIndex, 1, 1, 12); // อ่านทั้งแถว
  const rowData = range.getValues()[0];
  
  const entityId = rowData[V5_CONFIG.COL.QUEUE.SUGGESTED_ENTITY_ID - 1];
  const locationId = rowData[V5_CONFIG.COL.QUEUE.SUGGESTED_LOCATION_ID - 1];
  const now = new Date();

  if (action === 'APPROVE') {
    // 1. สร้างความสัมพันธ์ใน Entity_Loc_Map
    if (entityId && locationId) {
      // ตรวจสอบซ้ำก่อนสร้าง
      const existing = V5_FindMapRelation(mapSheet, entityId, locationId);
      if (!existing) {
        V5_CreateMapRelation(mapSheet, entityId, locationId, "VERIFIED_MANUAL");
      } else {
        // ถ้ามีอยู่แล้ว ให้เปิด Active
        const mapRowIndex = mapSheet.getDataRange().getValues().findIndex(r => r[0] === existing[0]) + 1;
        mapSheet.getRange(mapRowIndex, V5_CONFIG.COL.MAP.IS_ACTIVE).setValue(true);
      }
    }
    
    // 2. อัปเดตสถานะใน Queue
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.ACTION_REQUIRED).setValue("RESOLVED_APPROVED");
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.RESOLVED_BY).setValue(resolverName || "Admin");
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.RESOLVED_AT).setValue(now);
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.RESOLUTION_NOTE).setValue("Approved via UI");

    return { success: true, message: "อนุมัติเรียบร้อยแล้ว: สร้างความสัมพันธ์ใหม่" };

  } else if (action === 'REJECT') {
    // 1. เพียงแค่ตีตกไป ไม่สร้างความสัมพันธ์
    // อาจมีการทำ Soft Delete Entity/Location ที่สร้างชั่วคราวได้ในอนาคต (ขั้นสูง)
    
    // 2. อัปเดตสถานะใน Queue
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.ACTION_REQUIRED).setValue("RESOLVED_REJECTED");
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.RESOLVED_BY).setValue(resolverName || "Admin");
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.RESOLVED_AT).setValue(now);
    sheet.getRange(rowIndex, V5_CONFIG.COL.QUEUE.RESOLUTION_NOTE).setValue("Rejected via UI");

    return { success: true, message: "ปฏิเสธเรียบร้อยแล้ว: ไม่สร้างความสัมพันธ์" };
  }

  throw new Error("คำสั่งไม่ถูกต้อง");
}

/**
 * ฟังก์ชันเสริม: ล้างคิวที่แก้แล้วทั้งหมด (Archive)
 */
function V5_ArchiveResolvedConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let count = 0;
  
  // วนลูปจากล่างขึ้นบนเพื่อความปลอดภัยในการลบแถว (หรือจะซ่อนก็ได้)
  // ในที่นี้จะใช้วิธี Clear Content แทนการลบแถวเพื่อเก็บประวัติ
  for (let i = data.length - 1; i >= 1; i--) {
    const status = data[i][V5_CONFIG.COL.QUEUE.ACTION_REQUIRED - 1];
    if (status && status.toString().includes("RESOLVED")) {
      // ถ้าต้องการลบถาวรให้ใช้ sheet.deleteRow(i+1)
      // แต่ที่นี่เราจะแค่เคลียร์สีหรือทำเครื่องหมายว่าจบแล้ว
      count++;
    }
  }
  
  Browser.msgBox(`พบรายการที่แก้ไขแล้ว ${count} รายการ (ระบบเก็บไว้เป็นประวัติ)`);
}

/**
 * คำนวณสถิติคุณภาพข้อมูลสำหรับ Dashboard
 */
function V5_GetQualityStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entSheet = ss.getSheetByName(V5_CONFIG.SHEETS.ENTITIES);
  const locSheet = ss.getSheetByName(V5_CONFIG.SHEETS.LOCATIONS);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);

  const stats = {
    totalEntities: entSheet ? Math.max(0, entSheet.getLastRow() - 1) : 0,
    totalLocations: locSheet ? Math.max(0, locSheet.getLastRow() - 1) : 0,
    totalMappings: mapSheet ? Math.max(0, mapSheet.getLastRow() - 1) : 0,
    pendingConflicts: 0,
    resolvedConflicts: 0
  };

  if (queueSheet) {
    const qData = queueSheet.getDataRange().getValues();
    for (let i = 1; i < qData.length; i++) {
      const status = qData[i][V5_CONFIG.COL.QUEUE.ACTION_REQUIRED - 1];
      if (status && status.toString().includes("RESOLVED")) {
        stats.resolvedConflicts++;
      } else if (status && status !== "") {
        stats.pendingConflicts++;
      }
    }
  }

  // คำนวณความครอบคลุม (Coverage)
  stats.coverageRate = stats.totalEntities > 0 
    ? ((stats.totalMappings / stats.totalEntities) * 100).toFixed(2) + "%" 
    : "0%";

  return stats;
}
