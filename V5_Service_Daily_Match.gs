/**
 * V5_Service_Daily_Match.gs
 * Version: 5.0 (Final)
 * Description: Matches daily job data (from Data sheet) against the clean Master Data (Entities/Locations).
 *              Fills in LatLong_Actual, Email, and IDs.
 */

function V5_MatchDailyDataWithMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(V5_CONFIG.SHEETS.DATA);
  const empSheet = ss.getSheetByName(V5_CONFIG.SHEETS.EMPLOYEE);
  const entSheet = ss.getSheetByName(V5_CONFIG.SHEETS.ENTITIES);
  const locSheet = ss.getSheetByName(V5_CONFIG.SHEETS.LOCATIONS);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);

  if (!dataSheet || dataSheet.getLastRow() < 2) {
    Browser.msgBox("⚠️ ไม่มีข้อมูลในชีต Data");
    return;
  }

  // โหลดข้อมูล Master ทั้งหมดเข้า Memory (เพื่อความเร็ว)
  const entData = entSheet.getDataRange().getValues();
  const locData = locSheet.getDataRange().getValues();
  const mapData = mapSheet.getDataRange().getValues();
  const empData = empSheet ? empSheet.getDataRange().getValues() : [];

  // สร้าง Map สำหรับค้นหาเร็ว (Indexing)
  // Entity Map: Normalized_Name -> Entity_ID
  const entMap = {};
  for (let i = 1; i < entData.length; i++) {
    const normName = entData[i][V5_CONFIG.COL.ENTITIES.NORM_NAME - 1];
    const id = entData[i][V5_CONFIG.COL.ENTITIES.ID - 1];
    if (normName) entMap[normName] = id;
  }

  // Location Map: Location_ID -> {lat, lng}
  const locMap = {};
  for (let i = 1; i < locData.length; i++) {
    const id = locData[i][V5_CONFIG.COL.LOCATIONS.ID - 1];
    const lat = locData[i][V5_CONFIG.COL.LOCATIONS.LAT - 1];
    const lng = locData[i][V5_CONFIG.COL.LOCATIONS.LNG - 1];
    locMap[id] = { lat, lng };
  }

  // Map Relation: Entity_ID -> Location_ID (เอาอันที่เป็น Primary และ Active)
  const entityToLocMap = {};
  for (let i = 1; i < mapData.length; i++) {
    const eId = mapData[i][V5_CONFIG.COL.MAP.ENTITY_ID - 1];
    const lId = mapData[i][V5_CONFIG.COL.MAP.LOCATION_ID - 1];
    const isActive = mapData[i][V5_CONFIG.COL.MAP.IS_ACTIVE - 1];
    if (isActive && !entityToLocMap[eId]) {
      entityToLocMap[eId] = lId;
    }
  }

  // Employee Map: DriverName/Truck -> Email
  const empMap = {};
  if (empData.length > 1) {
    const headers = empData[0];
    const colName = headers.indexOf("ชื่อ - นามสกุล") + 1;
    const colEmail = headers.indexOf("Email พนักงาน") + 1;
    const colTruck = headers.indexOf("ทะเบียนรถ") + 1;
    
    for (let i = 1; i < empData.length; i++) {
      const name = empData[i][colName - 1];
      const truck = empData[i][colTruck - 1];
      const email = empData[i][colEmail - 1];
      if (name) empMap[V5_NormalizeText(name)] = email;
      if (truck) empMap[truck] = email; // เผื่อใช้ทะเบียนรถหา
    }
  }

  // --- เริ่มประมวลผลชีต Data ---
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  
  const colShipToName = headers.indexOf("ShipToName") + 1;
  const colDriver = headers.indexOf("DriverName") + 1;
  const colTruck = headers.indexOf("TruckLicense") + 1;
  const colLatLongActual = headers.indexOf("LatLong_Actual") + 1;
  const colEmail = headers.indexOf("Email พนักงาน") + 1;
  const colEntId = headers.indexOf("Matched_Entity_ID") + 1;
  const colLocId = headers.indexOf("Matched_Location_ID") + 1;

  let matchCount = 0;
  let updateRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawName = row[colShipToName - 1];
    if (!rawName) continue;

    const cleanName = V5_NormalizeText(rawName);
    let matchedEntityId = null;
    let matchedLocationId = null;
    let finalLat = "";
    let finalLng = "";
    let finalEmail = "";

    // 1. หา Entity
    if (entMap[cleanName]) {
      matchedEntityId = entMap[cleanName];
      
      // 2. หา Location จาก Entity นั้น
      if (entityToLocMap[matchedEntityId]) {
        matchedLocationId = entityToLocMap[matchedEntityId];
        const coords = locMap[matchedLocationId];
        if (coords) {
          finalLat = coords.lat;
          finalLng = coords.lng;
        }
      }
    }

    // 3. หา Email พนักงาน
    const driverName = row[colDriver - 1];
    const truck = row[colTruck - 1];
    if (driverName && empMap[V5_NormalizeText(driverName)]) {
      finalEmail = empMap[V5_NormalizeText(driverName)];
    } else if (truck && empMap[truck]) {
      finalEmail = empMap[truck];
    }

    // บันทึกผลลัพธ์เพื่ออัปเดต
    if (matchedEntityId && matchedLocationId) {
      matchCount++;
      updateRows.push({
        row: i + 1,
        entId: matchedEntityId,
        locId: matchedLocationId,
        latLong: `${finalLat},${finalLng}`,
        email: finalEmail
      });
    }
  }

  // เขียนผลลัพธ์ลง Sheet (Batch Write)
  if (updateRows.length > 0) {
    // เตรียม Array สำหรับแต่ละคอลัมน์
    const entIds = updateRows.map(r => [r.entId]);
    const locIds = updateRows.map(r => [r.locId]);
    const latLongs = updateRows.map(r => [r.latLong]);
    const emails = updateRows.map(r => [r.email]);
    const rowsNumbers = updateRows.map(r => r.row);

    // ใช้ Loop เขียนทีละคอลัมน์เพราะแถวไม่ต่อเนื่องกันอาจจะยาก ถ้าเขียนเป็นช่วงต่อเนื่องได้ง่ายกว่า
    // แต่ที่นี่เราเขียนทีละแถวเพื่อความแม่นยำ (หรือใช้ Range แยก)
    // วิธีเร็วสุด: เขียนทีละคอลัมน์แบบเป็นช่วงถ้าแถวเรียงกัน แต่ที่นี่อาจไม่เรียง
    
    // วิธีที่ปลอดภัยและเร็วพอสำหรับหลักพันแถว:
    updateRows.forEach(r => {
      dataSheet.getRange(r.row, colEntId).setValue(r.entId);
      dataSheet.getRange(r.row, colLocId).setValue(r.locId);
      dataSheet.getRange(r.row, colLatLongActual).setValue(r.latLong);
      if (r.email) dataSheet.getRange(r.row, colEmail).setValue(r.email);
    });
    
    // ทาสีเขียวให้แถวที่ Match สำเร็จ
    updateRows.forEach(r => {
      dataSheet.getRange(r.row, 1, 1, headers.length).setBackground("#e6f4ea");
    });
  }

  Browser.msgBox(`✅ จับคู่ข้อมูลเสร็จสิ้น!\n\n✔️ จับคู่สำเร็จ: ${matchCount} รายการ\n❌ ไม่พบข้อมูล: ${data.length - 1 - matchCount} รายการ\n\nรายการที่สำเร็จจะถูกเติมพิกัดและอีเมลให้ทันที`);
}
