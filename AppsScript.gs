/**
 * Google Apps Script API Backend cho tính năng "Văn mẫu NVN"
 * Spreadsheet ID: 1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g
 * Tên sheet: "Văn mẫu - NVN"
 */

function getSheet() {
  var spreadsheetId = "1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g";
  var sheetName = "Văn mẫu - NVN";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Tạo tiêu đề nếu sheet chưa có dữ liệu
    sheet.appendRow(["STT", "Tiếng anh", "Tiếng Việt", "Note"]);
  }
  return sheet;
}

function doGet(e) {
  var action = e.parameter.action;
  var sheet = getSheet();
  
  if (action === "read") {
    var data = sheet.getDataRange().getValues();
    // Bỏ qua dòng tiêu đề (row 1), lấy dữ liệu từ row 2
    var rows = [];
    if (data.length > 1) {
      rows = data.slice(1).map(function(row, index) {
        return { 
          id: index + 2, // Dòng thực tế trong Google Sheet (hàng đầu tiên là index + 2 vì dòng 1 là tiêu đề)
          stt: row[0] || (index + 1),
          english: row[1] || "", 
          vietnamese: row[2] || "", 
          note: row[3] || "" 
        };
      });
    }
    return responseJSON({ status: "success", data: rows });
  }
  
  return responseJSON({ status: "error", message: "Yêu cầu không hợp lệ" });
}

function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return responseJSON({ status: "error", message: "JSON payload không đúng định dạng" });
  }
  
  var action = params.action;
  var sheet = getSheet();

  if (action === "create") {
    // Thêm một dòng mới: cột 1 là trống/STT, cột 2 là english, cột 3 là vietnamese, cột 4 là note
    sheet.appendRow(["", params.english, params.vietnamese, params.note]);
    return responseJSON({ status: "success" });
  }

  if (action === "update") {
    var rowIndex = params.id; // Hàng cần sửa (id)
    if (!rowIndex || rowIndex < 2) {
      return responseJSON({ status: "error", message: "ID dòng sửa không hợp lệ" });
    }
    sheet.getRange(rowIndex, 2).setValue(params.english);
    sheet.getRange(rowIndex, 3).setValue(params.vietnamese);
    sheet.getRange(rowIndex, 4).setValue(params.note);
    return responseJSON({ status: "success" });
  }

  if (action === "delete") {
    var rowIndex = params.id; // Hàng cần xóa
    if (!rowIndex || rowIndex < 2) {
      return responseJSON({ status: "error", message: "ID dòng xóa không hợp lệ" });
    }
    sheet.deleteRow(rowIndex);
    return responseJSON({ status: "success" });
  }
  
  return responseJSON({ status: "error", message: "Action không được hỗ trợ" });
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
