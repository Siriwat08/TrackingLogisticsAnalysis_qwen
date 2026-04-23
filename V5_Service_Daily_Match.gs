/**
 * V5_Service_Daily_Match.gs
 * Version: 5.0
 * Description: The bridge between Part 1 (Daily Data) and Part 2 (Clean Master Data).
 *              Reads 'Data' sheet, matches names/coords against Entities/Locations, 
 *              and fills in LatLong_Actual and Employee Email.
 */

function V5_MatchDailyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(V5_CONFIG.SHEETS.DATA);
  const empSheet = ss.getSheetByName(V5_CONFIG.SHEETS.EMPLOYEE);
  
  if (!dataSheet || !empSheet) {
    Browser.msgBox("❌ ไม่พบชีต Data หรือ ข้อมูลพนักงาน");
    return;
  }

  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  
  // หาตำแหน่งคอลัมน์ในชีต Data แบบ Dynamic
  const colMap = {
    shipToName: headers.indexOf("ShipToName") + 1,
    latLongSCG: headers.indexOf("LatLong_SCG") + 1,
    address: headers.indexOf("ShipToAddress") + 1,
    driverName: headers.indexOf("DriverName") + 1, // ใช้หา Email พนักงาน
    truckLicense: headers.indexOf("TruckLicense") + 1,
    
    // คอลัมน์ที่จะเขียนผลลัพธ์
    latLongActual: headers.indexOf("LatLong_Actual") + 1,
    employeeEmail: headers.indexOf("Email พนักงาน") + 1,
    matchedEntityId: headers.indexOf("Matched_Entity_ID") + 1,
    matchedLocationId: headers.indexOf("Matched_Location_ID") + 1,
    matchConfidence: headers.indexOf("Match_Confidence") + 1 // เพิ่มคอลัมน์นี้ถ้ายังไม่มี
  };

  // ตรวจสอบว่าคอลัมน์สำคัญมีหรือไม่
  if (colMap.shipToName === 0 || colMap.latLongSCG === 0) {
    Browser.msgBox("❌ ไม่พบคอลัมน์ ShipToName หรือ LatLong_SCG ในชีต Data");
    return;
  }

  // โหลดข้อมูลพนักงานเข้า Memory (Map) เพื่อความเร็ว
  const empData = empSheet.getDataRange().getValues();
  const empMap = {}; // Key: ทะเบียนรถ หรือ ชื่อคนขับ, Value: Email
  // สมมติว่าค้นหารายการพนักงานจาก "ทะเบียนรถ" หรือ "ชื่อ-นามสกุล"
  // ในที่นี้ขอสมมติว่าใช้ "ทะเบียนรถ" เป็นหลักในการจับคู่ email
  const empColRegis = empData[0].indexOf("ทะเบียนรถ") + 1;
  const empColEmail = empData[0].indexOf("Email พนักงาน") + 1;
  
  for (let i = 1; i < empData.length; i++) {
    const regis = empData[i][empColRegis - 1];
    const email = empData[i][empColEmail - 1];
    if (regis) empMap[regis.toString().trim()] = email;
  }

  let matchCount = 0;
  let conflictCount = 0;
  let noMatchCount = 0;
  
  // เตรียม Array สำหรับเขียนกลับแบบ Batch (เพื่อความเร็ว)
  // เราจะอัปเดตหลายคอลัมน์พร้อมกัน
  const updates = []; 

  Logger.log(`เริ่มจับคู่ข้อมูลรายวัน แถวที่ 2 ถึง ${data.length}`);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const shipToName = row[colMap.shipToName - 1];
    const latLongStr = row[colMap.latLongSCG - 1];
    const address = row[colMap.address - 1] || "";
    const driverRegis = row[colMap.truckLicense - 1] ? row[colMap.truckLicense - 1].toString().trim() : "";

    // ถ้าไม่มีชื่อปลายทาง ข้ามไป
    if (!shipToName) {
      updates.push({ row: i + 1, updates: [] });
      continue;
    }

    let resultEntityId = "";
    let resultLocationId = "";
    let resultLatLong = "";
    let resultConfidence = 0;
    let statusLog = "";

    // --- 1. ค้นหาจากฐานข้อมูลใหม่ (Part 2) ---
    // แยก Lat/Long จากสตริง
    let lat = null, lng = null;
    if (latLongStr && latLongStr.includes(",")) {
      const parts = latLongStr.split(",");
      lat = parseFloat(parts[0]);
      lng = parseFloat(parts[1]);
    }

    if (lat && lng) {
      // เรียกใช้ Core Logic เพื่อค้นหา (แต่ไม่สร้างใหม่ ถ้าไม่มีในฐานข้อมูลหลัก)
      // หมายเหตุ: V5_ResolveIdentity จะสร้างใหม่ถ้าไม่เจอ ซึ่งอาจจะไม่เหมาะสำหรับ Daily Job ที่ต้องการแค่ "ดึงของที่มีอยู่"
      // ดังนั้น เราควรสร้างฟังก์ชันเฉพาะสำหรับการ "Search Only" ในไฟล์ Master_Core หรือแก้ไขพฤติกรรมที่นี่
      // เพื่อให้เข้าใจง่าย ในที่นี้เราจะใช้วิธี Search โดยตรงจากชีต Locations/Entities แทนการ Resolve ที่อาจสร้างขยะ
      
      const searchResult = V5_SearchExistingMaster(shipToName, lat, lng, address);
      
      if (searchResult.found) {
        resultEntityId = searchResult.entityId;
        resultLocationId = searchResult.locationId;
        resultLatLong = `${searchResult.lat},${searchResult.lng}`;
        resultConfidence = searchResult.confidence;
        statusLog = "MATCHED";
        matchCount++;
      } else {
        statusLog = "NO_MATCH_IN_MASTER";
        noMatchCount++;
        // กรณีไม่เจอ: อาจจะทิ้งว่าง หรือ ใช้พิกัด SCG เดิมชั่วคราว (แล้วแต่ความต้องการ)
        // ในที่นี้ขอทิ้งว่าง LatLong_Actual เพื่อให้รู้ว่ายังไม่ผ่านการตรวจสอบ
        resultLatLong = ""; 
      }
    } else {
      statusLog = "MISSING_COORDS";
      noMatchCount++;
    }

    // --- 2. ค้นหา Email พนักงาน ---
    let empEmail = "";
    if (driverRegis && empMap[driverRegis]) {
      empEmail = empMap[driverRegis];
    }

    // --- จัดเก็บผลลัพธ์เพื่อเขียนทีหลัง ---
    // ลำดับคอลัมน์ต้องตรงกับที่ต้องการจะเขียน
    // สมมติจะเขียน: LatLong_Actual, Email, Entity_ID, Location_ID, Confidence
    const rowDataUpdates = [
      resultLatLong, 
      empEmail, 
      resultEntityId, 
      resultLocationId, 
      resultConfidence > 0 ? resultConfidence : ""
    ];

    updates.push({ row: i + 1, updates: rowDataUpdates });
  }

  // --- เขียนผลลัพธ์ลงชีต Data แบบ Batch ---
  if (updates.length > 0 && colMap.latLongActual > 0) {
    // สร้าง Range ใหญ่ทีเดียว
    // คอลัมน์เริ่มต้นคือ LatLong_Actual, จำนวน 5 คอลัมน์
    const startRow = updates[0].row;
    // ต้องระวังเรื่องแถวที่กระโดดกัน อาจต้องใช้วิธีเขียนทีละช่วง หรือเขียนทั้งคอลัมน์
    // วิธีที่ง่ายและเร็วที่สุดสำหรับ Google Sheets คือการสร้าง Array ให้ครบทุกแถวแล้วเขียนทีเดียว
    
    // สร้าง Array ว่างขนาดเท่าจำนวนแถวข้อมูล
    const outputArray = new Array(data.length - 1).fill([]).map(() => ["", "", "", "", ""]);
    
    updates.forEach(u => {
      const rowIndex = u.row - 2; // ปรับ index ให้ตรงกับ Array (เริ่มที่ 0)
      if (rowIndex >= 0 && rowIndex < outputArray.length) {
        outputArray[rowIndex] = u.updates;
      }
    });

    const targetRange = dataSheet.getRange(2, colMap.latLongActual, outputArray.length, 5);
    targetRange.setValues(outputArray);
  }

  const summary = `
✅ เสร็จสิ้นการจับคู่ข้อมูลรายวัน!
---------------------------
📊 ทั้งหมด: ${data.length - 1} แถว
✨ จับคู่สำเร็จ: ${matchCount} รายการ
⚠️ ไม่พบในฐานข้อมูลหลัก: ${noMatchCount} รายการ (ตรวจสอบ Conflict_Queue หรือ รอ Admin อนุมัติ)

ข้อมูลพิกัดที่สะอาดแล้วและอีเมลพนักงานถูกเติมลงในชีต Data เรียบร้อยแล้ว
  `;
  
  Logger.log(summary);
  Browser.msgBox(summary);
}

/**
 * ฟังก์ชันเสริมสำหรับค้นหาเท่านั้น (ไม่สร้างใหม่)
 * ใช้สำหรับ Daily Job ที่ต้องการดึงข้อมูลที่มีอยู่แล้ว
 */
function V5_SearchExistingMaster(name, lat, lng, address) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entSheet = ss.getSheetByName(V5_CONFIG.SHEETS.ENTITIES);
  const locSheet = ss.getSheetByName(V5_CONFIG.SHEETS.LOCATIONS);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  
  const cleanName = V5_NormalizeText(name);
  
  // 1. หา Entity จากชื่อ
  let entityId = null;
  const entData = entSheet.getDataRange().getValues();
  for (let i = 1; i < entData.length; i++) {
    if (entData[i][V5_CONFIG.COL.ENTITIES.NORM_NAME - 1] === cleanName && 
        entData[i][V5_CONFIG.COL.ENTITIES.STATUS - 1] === "ACTIVE") {
      entityId = entData[i][V5_CONFIG.COL.ENTITIES.ID - 1];
      break;
    }
  }
  
  // 2. หา Location จากพิกัด (ในระยะ 200 ม.)
  let locationId = null;
  let bestLat = null;
  let bestLng = null;
  let minDist = 999999;
  
  const locData = locSheet.getDataRange().getValues();
  for (let i = 1; i < locData.length; i++) {
    const rLat = locData[i][V5_CONFIG.COL.LOCATIONS.LAT - 1];
    const rLng = locData[i][V5_CONFIG.COL.LOCATIONS.LNG - 1];
    if (rLat && rLng) {
      const dist = V5_CalculateDistanceMeters(lat, lng, rLat, rLng);
      if (dist <= 200 && dist < minDist) {
        minDist = dist;
        locationId = locData[i][V5_CONFIG.COL.LOCATIONS.ID - 1];
        bestLat = rLat;
        bestLng = rLng;
      }
    }
  }
  
  // 3. ถ้าเจอทั้งคู่ เช็คความสัมพันธ์ใน Map
  if (entityId && locationId) {
    const mapData = mapSheet.getDataRange().getValues();
    for (let i = 1; i < mapData.length; i++) {
      if (mapData[i][V5_CONFIG.COL.MAP.ENTITY_ID - 1] === entityId && 
          mapData[i][V5_CONFIG.COL.MAP.LOCATION_ID - 1] === locationId &&
          mapData[i][V5_CONFIG.COL.MAP.IS_ACTIVE - 1] === true) {
        
        return {
          found: true,
          entityId: entityId,
          locationId: locationId,
          lat: bestLat,
          lng: bestLng,
          confidence: 100
        };
      }
    }
  }
  
  // กรณีเจอแค่อย่างใดอย่างหนึ่ง อาจจะต้องใช้ Logic ย่อยๆ เพิ่มเติม (เช่น ใช้ชื่อหาพิกัดโดยประมาณ)
  // แต่ในที่นี้ขอคืนค่า Not Found ไปก่อนเพื่อความชัวร์ของข้อมูล
  return { found: false };
}
