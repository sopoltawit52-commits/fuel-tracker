function doGet(e) {
  var action = e.parameter.action;
  var sheet = getSheet();

  if (action === "getHistory") {
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return responseJSON({ success: true, rows: [] });
    }
    
    var headers = data[0];
    var rows = [];
    // นำ 20 รายการล่าสุด (ย้อนหลังจากแถวล่าสุด)
    for (var i = data.length - 1; i >= Math.max(1, data.length - 20); i--) {
      var item = {};
      for (var j = 0; j < headers.length; j++) {
        item[headers[j]] = data[i][j];
      }
      rows.push(item);
    }
    return responseJSON({ success: true, rows: rows });
  }

  return responseJSON({ success: true, message: "Fuel Tracker API is running" });
}

function doPost(e) {
  try {
    var sheet = getSheet();
    var payload;
    
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }

    var recordDate = payload.date || new Date().toISOString();
    var fuelType = payload.fuelType || "-";
    var totalCost = Number(payload.totalCost) || 0;
    var liters = Number(payload.liters) || 0;
    var pricePerLiter = liters > 0 ? (totalCost / liters).toFixed(2) : (payload.pricePerLiter || 0);
    var isFullTank = payload.isFullTank ? "ใช่ (เต็มถัง)" : "ไม่ใช่";
    var estDistance = Number(payload.estDistance) || 0; // ระยะทางโดยประมาณจาก Google Maps หรือเส้นทางประจำ
    var tripDesc = payload.tripDesc || "-"; // เช่น บ้าน-ที่ทำงาน, วิ่งส่งของในเมือง
    var costPerKm = (estDistance > 0 && totalCost > 0) ? (totalCost / estDistance).toFixed(2) : "-";
    var kmPerLiter = (estDistance > 0 && liters > 0) ? (estDistance / liters).toFixed(2) : "-";
    var note = payload.note || "";
    var timestamp = new Date();

    // บันทึกแถวใหม่
    sheet.appendRow([
      timestamp,
      recordDate,
      fuelType,
      totalCost,
      liters,
      pricePerLiter,
      isFullTank,
      estDistance,
      costPerKm,
      kmPerLiter,
      tripDesc,
      note
    ]);

    return responseJSON({
      success: true,
      message: "บันทึกข้อมูลเรียบร้อยแล้ว",
      data: {
        totalCost: totalCost,
        liters: liters,
        costPerKm: costPerKm,
        kmPerLiter: kmPerLiter
      }
    });

  } catch (error) {
    return responseJSON({
      success: false,
      error: error.toString()
    });
  }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("บันทึกน้ำมัน");
  
  if (!sheet) {
    sheet = ss.insertSheet("บันทึกน้ำมัน");
    // สร้างหัวตาราง (Headers) อัตโนมัติหากยังไม่มี
    sheet.appendRow([
      "Timestamp",
      "วันที่เติม",
      "ประเภทน้ำมัน",
      "ยอดเงิน (บาท)",
      "จำนวนลิตร",
      "ราคาต่อลิตร (บาท)",
      "เติมเต็มถัง",
      "ระยะทางประมาณการ (กม.)",
      "ต้นทุนต่อกม. (บาท/กม.)",
      "อัตราสิ้นเปลือง (กม./ลิตร)",
      "เส้นทาง/การใช้งาน",
      "หมายเหตุ"
    ]);
    sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#E2E8F0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
