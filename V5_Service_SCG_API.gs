/**
 * V5_Service_SCG_API.gs
 * Version: 5.0
 * Description: Fetches raw shipment data from SCG API based on Input sheet configurations 
 *              and writes it to the 'Data' sheet. Does NOT perform matching or cleaning.
 */

function V5_FetchRawDataFromAPI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(V5_CONFIG.SHEETS.INPUT);
  const dataSheet = ss.getSheetByName(V5_CONFIG.SHEETS.DATA);
  
  if (!inputSheet || !dataSheet) {
    Browser.msgBox("❌ ไม่พบชีต 'Input' หรือ 'Data' กรุณาตรวจสอบชื่อชีต");
    return;
  }

  // 1. อ่านค่าการตั้งค่าจากชีต Input
  const cookie = inputSheet.getRange("B1").getValue();
  const shipmentNosRange = inputSheet.getRange("A4:A100"); // สมมติใส่สูงสุด 100 แถว
  const shipmentNosValues = shipmentNosRange.getValues();
  
  // กรองแถวที่ว่างออก
  const shipmentNos = shipmentNosValues.flat().filter(String).map(s => s.toString().trim());

  if (!cookie) {
    Browser.msgBox("❌ กรุณาใส่ Cookie ในช่อง Input!B1 ก่อนดำเนินการ");
    return;
  }

  if (shipmentNos.length === 0) {
    Browser.msgBox("❌ ไม่พบรายการ Shipment Number ในช่อง Input!A4 ลงไป");
    return;
  }

  // แสดงยืนยัน
  const confirmMsg = `ต้องการดึงข้อมูลสำหรับ ${shipmentNos.length} Shipment(s) ใช่หรือไม่?`;
  if (Browser.msgBox(confirmMsg, Browser.Buttons.YES_NO) !== Browser.Buttons.YES) {
    return;
  }

  try {
    // 2. เรียก API (จำลองการเรียก - คุณต้องใส่ Logic การยิง API จริงของคุณที่นี่)
    // หมายเหตุ: เนื่องจากผมไม่เห็น Code เดิมของ fetchDataFromSCGJWD() แบบเต็ม 
    // ส่วนนี้คือโครงสร้างที่คุณต้องนำ Logic ยิง API เดิมของคุณมาใส่
    
    Logger.log(`กำลังดึงข้อมูลสำหรับ Shipment: ${shipmentNos.join(", ")}`);
    
    // --- จุดที่ต้องนำ Code เดิมของคุณมาใส่ ---
    // สมมติว่าฟังก์ชันเดิมชื่อ fetchDataFromSCGJWD(cookie, shipmentNos)
    // และคืนค่าเป็น Array of Objects ที่ตรงกับคอลัมน์ของชีต Data
    const apiResults = V5_CallSCGAPI_Internal(cookie, shipmentNos); 
    // -------------------------------------

    if (!apiResults || apiResults.length === 0) {
      Browser.msgBox("⚠️ ไม่สามารถดึงข้อมูลจาก API ได้ หรือข้อมูลว่างเปล่า");
      return;
    }

    // 3. เตรียม Header สำหรับชีต Data (ตามโครงสร้างที่คุณระบุ)
    const headers = [
      "ID_งานประจำวัน", "PlanDelivery", "InvoiceNo", "ShipmentNo", "DriverName", "TruckLicense", 
      "CarrierCode", "CarrierName", "SoldToCode", "SoldToName", "ShipToName", "ShipToAddress", 
      "LatLong_SCG", "MaterialName", "ItemQuantity", "QuantityUnit", "ItemWeight", "DeliveryNo", 
      "จำนวนปลายทาง_System", "รายชื่อปลายทาง_System", "ScanStatus", "DeliveryStatus", 
      "Email พนักงาน", "จำนวนสินค้ารวมของร้านนี้", "น้ำหนักสินค้ารวมของร้านนี้", 
      "จำนวน_Invoice_ที่ต้องสแกน", "LatLong_Actual", "ชื่อเจ้าของสินค้า_Invoice_ที่ต้องสแกน", 
      "ShopKey", "Matched_Entity_ID", "Matched_Location_ID" // เพิ่มคอลัมน์ใหม่สำหรับ Part 2
    ];

    // 4. เขียนข้อมูลลงชีต Data
    // ล้างข้อมูลเก่า (ถ้าต้องการ) หรือเขียนต่อท้าย
    // ที่นี่เลือกแบบล้างแล้วเขียนใหม่เพื่อความปลอดภัยของ Daily Job
    dataSheet.clearContents();
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    dataSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e0e0");

    // แปลงผลลัพธ์จาก API ให้เป็น Row Array
    const rowsToWrite = apiResults.map(item => [
      item.id || V5_GenerateUUID(),
      item.planDelivery,
      item.invoiceNo,
      item.shipmentNo,
      item.driverName,
      item.truckLicense,
      item.carrierCode,
      item.carrierName,
      item.soldToCode,
      item.soldToName,
      item.shipToName,
      item.shipToAddress,
      item.latLongSCG, // พิกัดจาก API (ยังไม่สะอาด)
      item.materialName,
      item.itemQuantity,
      item.quantityUnit,
      item.itemWeight,
      item.deliveryNo,
      item.destCountSystem,
      item.destNamesSystem,
      item.scanStatus,
      item.deliveryStatus,
      "", // Email พนักงาน (ว่างไว้รอจับคู่)
      item.totalQtyPerShop,
      item.totalWeightPerShop,
      item.invoiceCountPerShop,
      "", // LatLong_Actual (ว่างไว้รอจับคู่)
      item.ownerNamePerInvoice,
      item.shopKey,
      "", // Matched_Entity_ID (ว่างไว้รอจับคู่)
      ""  // Matched_Location_ID (ว่างไว้รอจับคู่)
    ]);

    if (rowsToWrite.length > 0) {
      dataSheet.getRange(2, 1, rowsToWrite.length, headers.length).setValues(rowsToWrite);
      
      // จัดรูปแบบเบื้องต้น
      dataSheet.setFrozenRows(1);
      dataSheet.setColumnWidth(11, 200); // ShipToName
      dataSheet.setColumnWidth(12, 250); // ShipToAddress
      
      Browser.msgBox(`✅ ดึงข้อมูลสำเร็จ!\nพบข้อมูลทั้งหมด ${rowsToWrite.length} รายการ\n\nขั้นตอนถัดไป:\nกรุณาใช้เมนู "🔄 จับคู่ข้อมูลสะอาด (Part 2)" เพื่อเติมพิกัดและอีเมล`);
    } else {
      Browser.msgBox("⚠️ ดึงข้อมูลได้แต่ไม่พบรายการที่นำไปบันทึกได้");
    }

  } catch (e) {
    Logger.log(`Error fetching  ${e.toString()}`);
    Browser.msgBox(`❌ เกิดข้อผิดพลาดในการดึงข้อมูล:\n${e.message}`);
  }
}

/**
 * ฟังก์ชันภายในสำหรับเรียก API (Placeholder)
 * คุณต้องนำ Logic การยิง API จาก Service_SCG.gs เดิมของคุณมาใส่ในส่วนนี้
 */
function V5_CallSCGAPI_Internal(cookie, shipmentNos) {
  // --- นำ Code เดิมจาก Service_SCG.gs มาใส่ที่นี่ ---
  // ตัวอย่างโครงสร้างการคืนค่าที่ต้องทำให้ได้:
  /*
  return [
    {
      planDelivery: "...",
      invoiceNo: "...",
      shipmentNo: "...",
      shipToName: "...",
      latLongSCG: "13.7563,100.5018",
      // ... อื่นๆ ครบตามคอลัมน์
    },
    ...
  ];
  */
  
  // จำลองข้อมูลทดสอบ (ลบออกเมื่อใส่ของจริง)
  Logger.log("⚠️ Warning: Using Mock Data. Please implement real API logic in V5_CallSCGAPI_Internal");
  return shipmentNos.map(no => ({
    planDelivery: new Date(),
    invoiceNo: `INV-${no}-01`,
    shipmentNo: no,
    driverName: "คนขับทดสอบ",
    truckLicense: "กข-1234",
    carrierCode: "C001",
    carrierName: "ขนส่งทดสอบ",
    soldToCode: "S001",
    soldToName: "ลูกค้าทดสอบ",
    shipToName: "ร้านทดสอบ สาขาใหญ่", // ชื่อนี้จะถูก拿去จับคู่กับ Part 2
    shipToAddress: "ถนนทดสอบ เขตทดสอบ กรุงเทพฯ",
    latLongSCG: "13.7563,100.5018", // พิกัดจาก API
    materialName: "สินค้า A",
    itemQuantity: 10,
    quantityUnit: "ชิ้น",
    itemWeight: 50,
    deliveryNo: `D-${no}`,
    destCountSystem: 1,
    destNamesSystem: "ร้านทดสอบ",
    scanStatus: "Pending",
    deliveryStatus: "Pending",
    totalQtyPerShop: 10,
    totalWeightPerShop: 50,
    invoiceCountPerShop: 1,
    ownerNamePerInvoice: "เจ้าของสินค้า A",
    shopKey: `${no}|ร้านทดสอบ`
  }));
}

/**
 * ฟังก์ชันเสริม: ล้างชีต Data และ Input
 */
function V5_ClearDailySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(V5_CONFIG.SHEETS.INPUT);
  const dataSheet = ss.getSheetByName(V5_CONFIG.SHEETS.DATA);
  
  if (Browser.msgBox("ต้องการล้างข้อมูลในชีต Input และ Data ทั้งหมดหรือไม่?", Browser.Buttons.YES_NO) === Browser.Buttons.YES) {
    if (inputSheet) inputSheet.getRange("A4:A100").clearContent();
    if (dataSheet) dataSheet.clearContents();
    Browser.msgBox("✅ ล้างข้อมูลเรียบร้อยแล้ว");
  }
}
