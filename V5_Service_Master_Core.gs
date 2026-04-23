/**
 * V5_Service_Master_Core.gs
 * Version: 5.0
 * Description: The Core Brain of Part 2. Handles logic for resolving Entities, Locations, and their Relationships.
 *              Detects conflicts and pushes them to the Queue for manual review.
 */

// --- MAIN RESOLVER FUNCTION ---

/**
 * ฟังก์ชันหลักในการระบุตัวตนและพิกัด
 * @param {string} rawName - ชื่อต้นทางจากข้อมูลดิบ
 * @param {number} lat - ละติจูด
 * @param {number} lng - ลองจิจูด
 * @param {string} address - ที่อยู่เต็ม (ถ้ามี)
 * @param {string} province - จังหวัด (ถ้ามี)
 * @returns {object} ผลลัพธ์ { entityId, locationId, status, message }
 */
function V5_ResolveIdentity(rawName, lat, lng, address, province) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // โหลดชีตที่เกี่ยวข้อง (ใช้ Cache เพื่อความเร็วหากเรียกซ้ำในลูป)
  const entSheet = ss.getSheetByName(V5_CONFIG.SHEETS.ENTITIES);
  const locSheet = ss.getSheetByName(V5_CONFIG.SHEETS.LOCATIONS);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  const aliasSheet = ss.getSheetByName(V5_CONFIG.SHEETS.NAME_MAPPING);

  // 1. ทำความสะอาดข้อมูลเบื้องต้น
  const cleanName = V5_NormalizeText(rawName);
  const latKey = V5_CreateLatLongKey(lat, lng);
  
  if (!cleanName || !lat || !lng) {
    return { entityId: null, locationId: null, status: "ERROR_INVALID_DATA", message: "ข้อมูลไม่ครบถ้วน" };
  }

  // 2. ค้นหาหรือสร้าง Location (สถานที่)
  // ใช้เกณฑ์ระยะทาง 200 เมตร ในการพิจารณาว่าเป็นที่เดียวกัน
  let locationId = V5_FindExistingLocation(locSheet, lat, lng, 200);
  let locStatus = "";
  
  if (!locationId) {
    locationId = V5_CreateNewLocation(locSheet, lat, lng, address, province);
    locStatus = "NEW_LOCATION";
  } else {
    locStatus = "EXISTING_LOCATION";
  }

  // 3. ค้นหาหรือสร้าง Entity (ตัวตน)
  // ค้นหาจากชื่อที่ทำความสะอาดแล้ว และ Alias
  let entityId = V5_FindExistingEntity(entSheet, aliasSheet, cleanName, province);
  let entStatus = "";

  if (!entityId) {
    entityId = V5_CreateNewEntity(entSheet, rawName, cleanName, province);
    entStatus = "NEW_ENTITY";
  } else {
    entStatus = "EXISTING_ENTITY";
  }

  // 4. ตรวจสอบความสัมพันธ์ (Relationship Check)
  // เช็คว่า Entity นี้ เคยถูกแมปกับ Location นี้มาก่อนหรือไม่?
  const existingMap = V5_FindMapRelation(mapSheet, entityId, locationId);
  
  if (existingMap) {
    // มีความสัมพันธ์อยู่แล้ว
    if (existingMap[V5_CONFIG.COL.MAP.IS_ACTIVE - 1] === true) {
      return { 
        entityId: entityId, 
        locationId: locationId, 
        status: "MATCHED_COMPLETE", 
        message: "พบข้อมูลที่ตรงกันทั้งคู่" 
      };
    } else {
      // เคยมีแต่ถูก Deactivate (อาจเคย Merge ไป) -> ต้องตรวจสอบ
      return { 
        entityId: entityId, 
        locationId: locationId, 
        status: "CONFLICT_INACTIVE_MAP", 
        message: "ความสัมพันธ์เดิมถูกปิดใช้งาน ต้องตรวจสอบ" 
      };
    }
  }

  // 5. ตรวจสอบความขัดแย้งก่อนสร้างความสัมพันธ์ใหม่ (Conflict Detection Logic)
  const conflictCheck = V5_DetectConflicts(entSheet, locSheet, mapSheet, entityId, locationId, cleanName, lat, lng);

  if (conflictCheck.hasConflict) {
    // พบความขัดแย้ง -> ส่งเข้า Queue
    V5_PushToConflictQueue(queueSheet, rawName, lat, lng, address, entityId, locationId, conflictCheck.reason, conflictCheck.details);
    return { 
      entityId: entityId, 
      locationId: locationId, 
      status: "CONFLICT_PENDING", 
      message: conflictCheck.reason 
    };
  }

  // 6. สร้างความสัมพันธ์ใหม่ (Auto-Link)
  V5_CreateMapRelation(mapSheet, entityId, locationId, "PRIMARY");
  
  return { 
    entityId: entityId, 
    locationId: locationId, 
    status: "AUTO_LINKED", 
    message: `สร้างใหม่: ${entStatus} + ${locStatus}` 
  };
}

// --- HELPER FUNCTIONS FOR LOCATIONS ---

function V5_FindExistingLocation(sheet, lat, lng, thresholdMeters) {
  const data = sheet.getDataRange().getValues();
  const colLat = V5_CONFIG.COL.LOCATIONS.LAT;
  const colLng = V5_CONFIG.COL.LOCATIONS.LNG;
  const colId = V5_CONFIG.COL.LOCATIONS.ID;

  // วนลูปหาตำแหน่งที่ใกล้ที่สุดในเกณฑ์ที่กำหนด
  for (let i = 1; i < data.length; i++) {
    const rLat = data[i][colLat - 1];
    const rLng = data[i][colLng - 1];
    
    if (rLat && rLng) {
      const dist = V5_CalculateDistanceMeters(lat, lng, rLat, rLng);
      if (dist <= thresholdMeters) {
        return data[i][colId - 1]; // เจอแล้ว คืนค่า ID
      }
    }
  }
  return null; // ไม่เจอ
}

function V5_CreateNewLocation(sheet, lat, lng, address, province) {
  const newId = V5_GenerateUUID();
  const latKey = V5_CreateLatLongKey(lat, lng);
  const now = new Date();
  
  sheet.appendRow([
    newId,          // Location_ID
    lat,            // Latitude
    lng,            // Longitude
    latKey,         // LatLong_Key
    address || "",  // Full_Address
    province || "", // Province
    "",             // District (รอ Geocode หรือใส่ทีหลัง)
    "",             // SubDistrict
    "",             // Postal_Code
    "UNKNOWN",      // Location_Type
    85,             // Confidence_Score (เริ่มต้นสูงเพราะมาจาก GPS โดยตรง)
    "SCG_RAW",      // Source
    now             // Last_Verified
  ]);
  return newId;
}

// --- HELPER FUNCTIONS FOR ENTITIES ---

function V5_FindExistingEntity(entSheet, aliasSheet, cleanName, province) {
  // 1. หาจากชื่อตรงใน Entities
  const entData = entSheet.getDataRange().getValues();
  const colNorm = V5_CONFIG.COL.ENTITIES.NORM_NAME;
  const colId = V5_CONFIG.COL.ENTITIES.ID;
  
  for (let i = 1; i < entData.length; i++) {
    if (entData[i][colNorm - 1] === cleanName) {
      return entData[i][colId - 1];
    }
  }

  // 2. หาจาก NameMapping (Alias)
  // TODO: เพิ่ม Logic การค้นหา Alias แบบละเอียดในอนาคต
  // ปัจจุบันเช็คคร่าวๆ ก่อน
  
  return null;
}

function V5_CreateNewEntity(sheet, displayName, cleanName, province) {
  const newId = V5_GenerateUUID();
  const now = new Date();
  
  sheet.appendRow([
    newId,              // Entity_ID
    displayName,        // Display_Name
    cleanName,          // Normalized_Name
    "SHOP/PERSON",      // Entity_Type (Default)
    "",                 // Phone
    "",                 // Tax_ID
    "ACTIVE",           // Status
    "",                 // Merged_To_Entity_ID
    now,                // Created_At
    now                 // Updated_At
  ]);
  return newId;
}

// --- HELPER FUNCTIONS FOR MAPPING & CONFLICTS ---

function V5_FindMapRelation(sheet, entityId, locationId) {
  const data = sheet.getDataRange().getValues();
  const colEnt = V5_CONFIG.COL.MAP.ENTITY_ID;
  const colLoc = V5_CONFIG.COL.MAP.LOCATION_ID;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colEnt - 1] === entityId && data[i][colLoc - 1] === locationId) {
      return data[i];
    }
  }
  return null;
}

function V5_CreateMapRelation(sheet, entityId, locationId, relationType) {
  const newId = V5_GenerateUUID();
  const now = new Date();
  
  sheet.appendRow([
    newId,              // Map_ID
    entityId,           // Entity_ID
    locationId,         // Location_ID
    relationType,       // Relation_Type
    true,               // Is_Active
    95,                 // Confidence
    "",                 // Notes
    now                 // Created_At
  ]);
}

function V5_DetectConflicts(entSheet, locSheet, mapSheet, newEntityId, newLocationId, cleanName, lat, lng) {
  // Logic ตรวจความขัดแย้ง 8 ข้อ (แบบย่อสำหรับ Core)
  
  // 1. เช็ค Entity นี้ มี Location อื่นอีกไหม? (กรณี 1 ชื่อ หลายที่)
  const existingMapsForEntity = V5_GetMapsByEntity(mapSheet, newEntityId);
  if (existingMapsForEntity.length > 0) {
    // มี Location อื่นอยู่แล้ว เช็คระยะทาง
    for (const mapRow of existingMapsForEntity) {
      const otherLocId = mapRow[V5_CONFIG.COL.MAP.LOCATION_ID - 1];
      const otherLocCoords = V5_GetLocationCoords(locSheet, otherLocId);
      
      if (otherLocCoords) {
        const dist = V5_CalculateDistanceMeters(lat, lng, otherLocCoords.lat, otherLocCoords.lng);
        if (dist > 1000) { // ถ้าห่างเกิน 1 กม. ถือว่าน่าสงสัย (Branch หรือ ผิดพลาด?)
          return {
            hasConflict: true,
            reason: "ONE_ENTITY_MANY_LOCATIONS_FAR",
            details: `ชื่อนี้มีที่อยู่เดิมแล้ว แต่ห่างจากจุดใหม่ ${Math.round(dist)} เมตร`
          };
        }
      }
    }
  }

  // 2. เช็ค Location นี้ มี Entity อื่นอีกไหม? (กรณี 1 ที่ หลายชื่อ)
  const existingMapsForLocation = V5_GetMapsByLocation(mapSheet, newLocationId);
  if (existingMapsForLocation.length > 0) {
    // มีชื่ออื่นอยู่แล้วที่พิกัดนี้
    // อันนี้มักจะเป็น "ตลาด" หรือ "ตึกสำนักงาน" ซึ่งอาจไม่ใช่ Conflict ร้ายแรงเสมอไป
    // แต่เราจะแจ้งเตือนเพื่อให้ตรวจสอบว่าเป็นคนละบริษัทจริงหรือไม่
    /* 
       หมายเหตุ: ในเวอร์ชันแรก เราอาจจะยังไม่ถือเป็นเรื่องขัดแย้งร้ายแรง 
       แต่ถ้าต้องการเข้มงวด ให้เปิดโค้ดบรรทัดล่างนี้
    */
    /*
    return {
      hasConflict: true,
      reason: "ONE_LOCATION_MANY_ENTITIES",
      details: "พิกัดนี้มีเจ้าของอื่นอยู่แล้ว ตรวจสอบว่าเป็นคนละหน่วยงานหรือไม่"
    };
    */
  }

  return { hasConflict: false, reason: "", details: "" };
}

function V5_GetMapsByEntity(sheet, entityId) {
  const data = sheet.getDataRange().getValues();
  const colEnt = V5_CONFIG.COL.MAP.ENTITY_ID;
  const colActive = V5_CONFIG.COL.MAP.IS_ACTIVE;
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colEnt - 1] === entityId && data[i][colActive - 1] === true) {
      results.push(data[i]);
    }
  }
  return results;
}

function V5_GetMapsByLocation(sheet, locationId) {
  const data = sheet.getDataRange().getValues();
  const colLoc = V5_CONFIG.COL.MAP.LOCATION_ID;
  const colActive = V5_CONFIG.COL.MAP.IS_ACTIVE;
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colLoc - 1] === locationId && data[i][colActive - 1] === true) {
      results.push(data[i]);
    }
  }
  return results;
}

function V5_GetLocationCoords(sheet, locationId) {
  const data = sheet.getDataRange().getValues();
  const colId = V5_CONFIG.COL.LOCATIONS.ID;
  const colLat = V5_CONFIG.COL.LOCATIONS.LAT;
  const colLng = V5_CONFIG.COL.LOCATIONS.LNG;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colId - 1] === locationId) {
      return { lat: data[i][colLat - 1], lng: data[i][colLng - 1] };
    }
  }
  return null;
}

function V5_PushToConflictQueue(sheet, rawName, lat, lng, address, entId, locId, reason, details) {
  const newId = V5_GenerateUUID();
  const now = new Date();
  
  sheet.appendRow([
    newId,                    // Queue_ID
    now,                      // Received_At
    rawName,                  // Incoming_Name
    `${lat},${lng}`,          // Incoming_LatLong
    address || "",            // Incoming_Address
    entId,                    // Suggested_Entity_ID
    locId,                    // Suggested_Location_ID
    reason,                   // Conflict_Reason
    "REVIEW_REQUIRED",        // Action_Required
    "",                       // Resolved_By
    "",                       // Resolved_At
    details                   // Resolution_Note
  ]);
}
