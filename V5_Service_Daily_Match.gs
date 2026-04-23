/**
 * V5_Service_Daily_Match.gs
 * Version: 5.0
 * Description: The bridge between Part 1 (Daily Data) and Part 2 (Clean Master Data).
 *              Matches incoming shipment data against Entities/Locations to fill LatLong_Actual and Employee Email.
 */

function V5_ProcessDailyMatching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(V5_CONFIG.SHEETS.DATA);
  const empSheet = ss.getSheetByName(V5_CONFIG.SHEETS.EMPLOYEE);
  
  if (!dataSheet || !empSheet) {
    Browser.msgBox("❌ ไม่พบชีต 'Data' หรือ 'ข้อมูลพนักงาน'");
    return;
  }

  // ดึงข้อมูลทั้งหมดจากชีต Data
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  
  // หาตำแหน่งคอลัมน์ในชีต Data (Dynamic)
  const colMap = {
    shipToName: headers.indexOf("ShipToName") + 1,
    latLongSCG: headers.indexOf("LatLong_SCG") + 1,
    address: headers.indexOf("ShipToAddress") + 1,
    driverName: headers.indexOf("DriverName") + 1, // สำหรับหาอีเมลพนักงาน
    latLongActual: headers.indexOf("LatLong_Actual") + 1,
    emailCol: headers.indexOf("Email พนักงาน") + 1,
    // คอลัมน์ใหม่สำหรับเก็บ ID (ถ้ามี)
    entityIdCol: headers.indexOf("Matched_Entity_ID") + 1,
    locationIdCol: headers.indexOf("Matched_Location_ID") + 1
  };

  if (colMap.shipToName === 0 || colMap.latLongSCG === 0) {
    Browser.msgBox("❌ ไม่พบคอลัมน์ 'ShipToName' หรือ 'LatLong_SCG' ในชีต Data");
    return;
  }

  // โหลดข้อมูลพนักงานเข้า Memory (Map) เพื่อความเร็ว
  const employeeMap = V5_BuildEmployeeMap(empSheet);

  // โหลดข้อมูล Master Data เข้า Memory (Cache) เพื่อความเร็วในการค้นหา
  // หมายเหตุ: ในระบบจริงที่มีข้อมูลหลักแสนแถว อาจต้องใช้วิธีค้นหาแบบทีละน้อยหรือใช้ Cache Service
  // แต่สำหรับ Google Sheets ปกติ การโหลดมาทั้ง Sheet ยังพอทำได้หากไม่เกิน 5-10 พันแถว
  const entSheet = ss.getSheetByName(V5_CONFIG.SHEETS.ENTITIES);
  const locSheet = ss.getSheetByName(V5_CONFIG.SHEETS.LOCATIONS);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  const aliasSheet = ss.getSheetByName(V5_CONFIG.SHEETS.NAME_MAPPING);
  
  const masterCache = {
    entities: entSheet.getDataRange().getValues(),
    locations: locSheet.getDataRange().getValues(),
    maps: mapSheet.getDataRange().getValues(),
    aliases: aliasSheet ? aliasSheet.getDataRange().getValues() : []
  };

  let matchCount = 0;
  let noMatchCount = 0;
  const updates = []; // รวบรวมผลเพื่อเขียนทีเดียว

  Logger.log(`เริ่มจับคู่ข้อมูลรายวัน จำนวน ${data.length - 1} แถว...`);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const shipToName = row[colMap.shipToName - 1];
    const latLongStr = row[colMap.latLongSCG - 1];
    const address = row[colMap.address - 1] || "";
    const driverName = row[colMap.driverName - 1] || "";

    let finalLatLong = "";
    let matchedEntityId = "";
    let matchedLocationId = "";
    let isMatched = false;

    // 1. พยายามจับคู่กับ Master Data ใหม่ (Part 2)
    if (shipToName) {
      const latLngParts = latLongStr ? latLongStr.split(',') : [null, null];
      const lat = latLngParts[0] ? parseFloat(latLngParts[0]) : null;
      const lng = latLngParts[1] ? parseFloat(latLngParts[1]) : null;

      // เรียกฟังก์ชันค้นหา (ไม่สร้างใหม่ แค่ค้นหา)
      const searchResult = V5_SearchBestMatch(shipToName, lat, lng, address, masterCache);
      
      if (searchResult.found) {
        matchedEntityId = searchResult.entityId;
        matchedLocationId = searchResult.locationId;
        finalLatLong = `${searchResult.lat},${searchResult.lng}`;
        isMatched = true;
        matchCount++;
      } else {
        noMatchCount++;
        // ถ้าไม่เจอในระบบใหม่ อาจจะลองตกกลับไปใช้ Logic เก่า (ถ้ามี) หรือปล่อยว่างไว้ก่อน
        // ในที่นี้เราเน้นใช้ระบบใหม่เป็นหลัก
      }
    }

    // 2. หาอีเมลพนักงาน
    let employeeEmail = "";
    if (driverName && employeeMap[driverName]) {
      employeeEmail = employeeMap[driverName];
    }

    // เตรียมข้อมูลสำหรับการอัปเดต (เฉพาะแถวที่มีการเปลี่ยนแปลงหรือต้องการเติมข้อมูล)
    // เราเขียนเฉพาะคอลัมน์ที่สำคัญ: LatLong_Actual, Email, Entity_ID, Location_ID
    updates.push([
      finalLatLong,             // LatLong_Actual
      employeeEmail,            // Email พนักงาน
      matchedEntityId,          // Matched_Entity_ID
      matchedLocationId         // Matched_Location_ID
    ]);
  }

  // เขียนผลลัพธ์ลงชีต Data (Batch Write)
  if (updates.length > 0) {
    // กำหนดช่วงที่จะเขียน: เริ่มที่แถว 2, คอลัมน์แรกของกลุ่มที่ต้องการเขียน
    // ลำดับคอลัมน์ที่ต้องเขียนต้องตรงกับลำดับใน updates array
    // สมมติว่าเราเขียน 4 คอลัมน์ติดกัน หรือต้องระบุทีละช่วงหากคอลัมน์ไม่ติดกัน
    
    // วิธีที่ปลอดภัยที่สุดคือเขียนทีละคอลัมน์หากคอลัมน์ไม่อยู่ติดกัน
    // แต่เพื่อความรวดเร็ว เราจะเขียนเป็นบล็อกหากคอลัมน์อยู่ติดกัน
    // ในที่นี้สมมติว่าคอลัมน์เหล่านี้可能會อยู่กระจายกัน ดังนั้นควรเขียนแยก
    
    // เขียน LatLong_Actual
    if (colMap.latLongActual > 0) {
      const colData = updates.map(r => [r[0]]);
      dataSheet.getRange(2, colMap.latLongActual, updates.length, 1).setValues(colData);
    }
    // เขียน Email
    if (colMap.emailCol > 0) {
      const colData = updates.map(r => [r[1]]);
      dataSheet.getRange(2, colMap.emailCol, updates.length, 1).setValues(colData);
    }
    // เขียน Entity ID (ถ้ามีคอลัมน์)
    if (colMap.entityIdCol > 0) {
      const colData = updates.map(r => [r[2]]);
      dataSheet.getRange(2, colMap.entityIdCol, updates.length, 1).setValues(colData);
    }
    // เขียน Location ID (ถ้ามีคอลัมน์)
    if (colMap.locationIdCol > 0) {
      const colData = updates.map(r => [r[3]]);
      dataSheet.getRange(2, colMap.locationIdCol, updates.length, 1).setValues(colData);
    }
  }

  const summary = `
✅ เสร็จสิ้นการจับคู่ข้อมูลรายวัน!
---------------------------
📊 ทั้งหมด: ${data.length - 1} แถว
✅ จับคู่สำเร็จ: ${matchCount} รายการ
⚠️ ไม่พบข้อมูล: ${noMatchCount} รายการ (อาจต้องตรวจสอบใน Conflict Queue หรือเพิ่มใหม่)

ระบบได้เติม LatLong_Actual และ Email พนักงานเรียบร้อยแล้ว
  `;
  
  Logger.log(summary);
  Browser.msgBox(summary);
}

// --- Helper Functions for Matching ---

function V5_BuildEmployeeMap(sheet) {
  const data = sheet.getDataRange().getValues();
  const map = {};
  // สมมติคอลัมน์: ชื่อ-นามสกุล (2), Email (7) - เช็คจาก Config หรือ Hardcode ตามโครงสร้าง
  // ใช้ Index จาก Config ถ้ามี หรือนับเอง
  const colName = 2; 
  const colEmail = 7;
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][colName - 1];
    const email = data[i][colEmail - 1];
    if (name && email) {
      map[name.trim()] = email;
      // เพิ่มвариants ของชื่อถ้าจำเป็น (เช่น ตัดคำนำหน้า)
    }
  }
  return map;
}

function V5_SearchBestMatch(name, lat, lng, address, cache) {
  const cleanName = V5_NormalizeText(name);
  
  // 1. ค้นหาจาก Entity Name (Direct Match)
  const entColNorm = V5_CONFIG.COL.ENTITIES.NORM_NAME - 1;
  const entColId = V5_CONFIG.COL.ENTITIES.ID - 1;
  
  let candidateEntityId = null;
  
  // หา Entity ที่ชื่อตรงกัน
  for (let i = 1; i < cache.entities.length; i++) {
    if (cache.entities[i][entColNorm] === cleanName) {
      candidateEntityId = cache.entities[i][entColId];
      break;
    }
  }
  
  // ถ้าไม่เจอชื่อตรง ลองหาจาก Alias (NameMapping)
  if (!candidateEntityId && cache.aliases) {
    const aliasColName = V5_CONFIG.COL.MAPPING.ALIAS - 1;
    const aliasColTarget = V5_CONFIG.COL.MAPPING.TARGET_ID - 1;
    for (let i = 1; i < cache.aliases.length; i++) {
      if (V5_NormalizeText(cache.aliases[i][aliasColName]) === cleanName) {
        candidateEntityId = cache.aliases[i][aliasColTarget];
        break;
      }
    }
  }

  // 2. ถ้าเจอ Entity แล้ว ให้หา Location ที่เชื่อมอยู่
  if (candidateEntityId) {
    const mapColEnt = V5_CONFIG.COL.MAP.ENTITY_ID - 1;
    const mapColLoc = V5_CONFIG.COL.MAP.LOCATION_ID - 1;
    const mapColActive = V5_CONFIG.COL.MAP.IS_ACTIVE - 1;
    
    for (let i = 1; i < cache.maps.length; i++) {
      if (cache.maps[i][mapColEnt] === candidateEntityId && cache.maps[i][mapColActive] === true) {
        const locId = cache.maps[i][mapColLoc];
        
        // ดึงพิกัดจาก Location ID
        const locCoords = V5_GetCoordsFromCache(locId, cache.locations);
        if (locCoords) {
          return {
            found: true,
            entityId: candidateEntityId,
            locationId: locId,
            lat: locCoords.lat,
            lng: locCoords.lng,
            confidence: 100
          };
        }
      }
    }
  }

  // 3. Fallback: ถ้าไม่เจอชื่อ แต่มีพิกัด ให้ลองหา Location ที่ใกล้ที่สุด (ภายในระยะที่กำหนด)
  if (lat && lng) {
    const locColLat = V5_CONFIG.COL.LOCATIONS.LAT - 1;
    const locColLng = V5_CONFIG.COL.LOCATIONS.LNG - 1;
    const locColId = V5_CONFIG.COL.LOCATIONS.ID - 1;
    
    let bestDist = 999999;
    let bestLocId = null;
    let bestCoords = null;
    
    // หาระยะทางกับทุก Location (อาจช้าถ้าข้อมูลเยอะมาก ควรทำ Index หรือจำกัดขอบเขต)
    // สำหรับการเริ่มต้น วนลูปได้เลย
    for (let i = 1; i < cache.locations.length; i++) {
      const rLat = cache.locations[i][locColLat];
      const rLng = cache.locations[i][locColLng];
      
      if (rLat && rLng) {
        const dist = V5_CalculateDistanceMeters(lat, lng, rLat, rLng);
        if (dist < 50 && dist < bestDist) { // ระยะห่างไม่เกิน 50 เมตร ถือว่าใกล้ enough สำหรับ Fallback
          bestDist = dist;
          bestLocId = cache.locations[i][locColId];
          bestCoords = { lat: rLat, lng: rLng };
        }
      }
    }
    
    if (bestLocId) {
      return {
        found: true,
        entityId: "", // ไม่รู้ Entity แน่นอน
        locationId: bestLocId,
        lat: bestCoords.lat,
        lng: bestCoords.lng,
        confidence: 80 // ความมั่นใจลดลงเพราะจับจากพิกัดล้วนๆ
      };
    }
  }

  return { found: false };
}

function V5_GetCoordsFromCache(locationId, locationsData) {
  const colId = V5_CONFIG.COL.LOCATIONS.ID - 1;
  const colLat = V5_CONFIG.COL.LOCATIONS.LAT - 1;
  const colLng = V5_CONFIG.COL.LOCATIONS.LNG - 1;
  
  for (let i = 1; i < locationsData.length; i++) {
    if (locationsData[i][colId] === locationId) {
      return { lat: locationsData[i][colLat], lng: locationsData[i][colLng] };
    }
  }
  return null;
}
