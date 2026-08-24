/**
 * GOOGLE APPS SCRIPT BACKEND – VĂN MẪU NVN (Chỉ chuyên xử lý Văn mẫu)
 * Spreadsheet ID: 1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g
 * Tên sheet: "Văn mẫu - NVN"
 */

var SPREADSHEET_ID = "1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g";
var VAN_MAU_SHEET  = "Văn mẫu - NVN";

function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(VAN_MAU_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(VAN_MAU_SHEET);
    sheet.appendRow(["STT", "Tiếng anh", "Tiếng Việt", "Note"]);
  }
  return sheet;
}

// ─── GET Router ───────────────────────────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action;

  // Hỗ trợ Fallback GET Bypass CORS
  if (e.parameter.data) {
    var params;
    try {
      params = JSON.parse(decodeURIComponent(e.parameter.data));
    } catch (err) {
      return responseJSON({ status: "error", message: "Dữ liệu JSON fallback không hợp lệ" });
    }
    return handlePostActions(action, params);
  }

  if (action === "read") {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var rows = [];
    if (data.length > 1) {
      rows = data.slice(1).map(function(row, index) {
        return { 
          id: index + 2, 
          english: row[1] || "", 
          vietnamese: row[2] || "", 
          note: row[3] || "" 
        };
      });
    }
    return responseJSON({ status: "success", data: rows });
  }

  return responseJSON({ status: "error", message: "Action không hợp lệ: " + action });
}

// ─── POST Router ──────────────────────────────────────────────────────────────
function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return responseJSON({ status: "error", message: "JSON payload không hợp lệ" });
  }

  var action = params.action;
  return handlePostActions(action, params);
}

// ─── Xử lý chung các thao tác ghi dữ liệu ──────────────────────────────────────
function handlePostActions(action, params) {
  var sheet = getSheet();

  if (action === "create") {
    sheet.appendRow(["", params.english, params.vietnamese, params.note || ""]);
    return responseJSON({ status: "success" });
  }

  if (action === "update") {
    var rowIndex = parseInt(params.id);
    if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
    sheet.getRange(rowIndex, 2).setValue(params.english);
    sheet.getRange(rowIndex, 3).setValue(params.vietnamese);
    sheet.getRange(rowIndex, 4).setValue(params.note || "");
    return responseJSON({ status: "success" });
  }

  if (action === "delete") {
    var rowIndex = parseInt(params.id);
    if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
    sheet.deleteRow(rowIndex);
    return responseJSON({ status: "success" });
  }

  return responseJSON({ status: "error", message: "Action không được hỗ trợ: " + action });
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
