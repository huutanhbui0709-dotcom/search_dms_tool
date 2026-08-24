/**
 * GOOGLE APPS SCRIPT BACKEND UNIFIED – NVN SYSTEM
 * Spreadsheet ID: 1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g
 * 
 * Hợp nhất cả hai tính năng:
 * 1. Task Pending (Sheet: "Công việc Pending")
 * 2. Văn mẫu (Sheet: "Văn mẫu - NVN")
 */

var SPREADSHEET_ID = "1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g";
var TASK_SHEET     = "Công việc Pending";
var VAN_MAU_SHEET  = "Văn mẫu - NVN";

// ─── Helpers: Lấy Sheet ────────────────────────────────────────────────────────
function getTaskSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TASK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TASK_SHEET);
    var headers = ["ID", "Hạn cuối", "Issue", "Ghi chú", "Người phụ trách",
                   "Trạng thái", "", "", "Thành viên", "URL Hệ thống", "Tên Hệ thống"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1e1b4b")
      .setFontColor("#ffffff");
  }
  return sheet;
}

function getVanMauSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(VAN_MAU_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(VAN_MAU_SHEET);
    sheet.appendRow(["STT", "Tiếng anh", "Tiếng Việt", "Note"]);
  }
  return sheet;
}

// ─── Helpers: đọc danh sách URL hệ thống từ Cột J & K ──────────────────────────
function readSystemUrls(sheet) {
  var lastRow    = sheet.getLastRow();
  var systemUrls = [];
  if (lastRow >= 3) {
    var colJK = sheet.getRange(3, 10, lastRow - 2, 2).getValues(); // Cột J(10) & K(11)
    colJK.forEach(function (row, idx) {
      var url  = String(row[0] || "").trim();
      var name = String(row[1] || "").trim();
      if (url || name) {
        systemUrls.push({ rowIndex: idx + 3, url: url, name: name });
      }
    });
  }
  return systemUrls;
}

// ─── Helper: tìm dòng trống đầu tiên của danh sách task (Cột B và C) ──────────
function getNextTaskRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  var colBC = sheet.getRange(2, 2, lastRow - 1, 2).getValues(); // Đọc cột B (Hạn cuối) và C (Issue)
  for (var i = 0; i < colBC.length; i++) {
    var dateVal = colBC[i][0];
    var issueVal = colBC[i][1];
    if (!dateVal && !issueVal) {
      return i + 2; // trả về rowIndex 1-based (2 + i)
    }
  }
  return lastRow + 1;
}

// ─── GET Router ───────────────────────────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action;

  // Hỗ trợ Fallback GET Bypass CORS (mã hóa JSON qua data)
  if (e.parameter.data) {
    var params;
    try {
      params = JSON.parse(decodeURIComponent(e.parameter.data));
    } catch (err) {
      return responseJSON({ status: "error", message: "Dữ liệu JSON fallback không hợp lệ" });
    }
    return handlePostActions(action, params);
  }

  // ── TASK: Đọc danh sách Task ───────────────────────────────────────────────
  if (action === "read_tasks") {
    var sheet      = getTaskSheet();
    var data       = sheet.getDataRange().getValues();
    var rows       = [];
    var systemUrls = readSystemUrls(sheet);

    if (data.length > 1) {
      data.slice(1).forEach(function (row, idx) {
        if (!row[2] && !row[1]) return;
        var dateStr = "";
        if (row[1]) {
          try {
            dateStr = (row[1] instanceof Date)
              ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd")
              : String(row[1]).split("T")[0];
          } catch (e) {
            dateStr = String(row[1]).split("T")[0];
          }
        }
        rows.push({
          id:        idx + 2,
          deadline:  dateStr,
          dueDate:   dateStr,
          issue:     row[2] || "",
          note:      row[3] || "",
          assignees: row[4]
            ? String(row[4]).split(",").map(function (s) { return s.trim(); }).filter(Boolean)
            : [],
          status:    row[5] || "Pending"
        });
      });
    }

    var savedUrl = "";
    systemUrls.forEach(function (s) {
      if (s.name.toLowerCase().indexOf("task") !== -1 || s.name.toLowerCase().indexOf("pending") !== -1) {
        savedUrl = s.url;
      }
    });

    return responseJSON({
      status:     "success",
      data:       rows,
      tasks:      rows,
      systemUrls: systemUrls,
      savedUrl:   savedUrl
    });
  }

  // ── TASK: Đọc danh sách Members từ Cột I ────────────────────────────────────
  if (action === "read_members") {
    var sheet   = getTaskSheet();
    var lastRow = sheet.getLastRow();
    var members = [];
    if (lastRow >= 3) {
      var colI = sheet.getRange(3, 9, lastRow - 2, 1).getValues();
      colI.forEach(function (row, idx) {
        var name = String(row[0] || "").trim();
        if (name) members.push({ rowIndex: idx + 3, name: name });
      });
    }
    return responseJSON({ status: "success", data: members });
  }

  // ── TASK: Đọc danh sách System URLs ──────────────────────────────────────────
  if (action === "read_system_urls") {
    var sheet      = getTaskSheet();
    var systemUrls = readSystemUrls(sheet);
    return responseJSON({ status: "success", data: systemUrls });
  }

  // ── VĂN MẪU: Đọc danh sách Văn Mẫu ───────────────────────────────────────────
  if (action === "read") {
    var sheet = getVanMauSheet();
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
  return handlePostActions(params.action, params);
}

// ─── Xử lý chung các thao tác ghi dữ liệu ─────────────────────────────────────
function handlePostActions(action, params) {
  try {
    // ── TASK ACTIONS ─────────────────────────────────────────────────────────
    if (action === "save_task") {
      var sheet = getTaskSheet();
      var assigneesStr = Array.isArray(params.assignees) ? params.assignees.join(", ") : (params.assignees || "");
      var status   = params.status   || "Pending";
      var deadline = params.deadline || params.dueDate || "";

      if (params.id) {
        var rowIndex = parseInt(params.id);
        if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
        sheet.getRange(rowIndex, 2).setValue(deadline);
        sheet.getRange(rowIndex, 3).setValue(params.issue || "");
        sheet.getRange(rowIndex, 4).setValue(params.note  || "");
        sheet.getRange(rowIndex, 5).setValue(assigneesStr);
        sheet.getRange(rowIndex, 6).setValue(status);
      } else {
        var nextRow = getNextTaskRow(sheet);
        sheet.getRange(nextRow, 1, 1, 6).setValues([["", deadline, params.issue || "", params.note || "", assigneesStr, status]]);
      }
      return responseJSON({ status: "success" });
    }

    if (action === "delete_task") {
      var sheet = getTaskSheet();
      var rowIndex = parseInt(params.id);
      if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
      
      var lastRow = sheet.getLastRow();
      if (rowIndex < lastRow) {
        var values = sheet.getRange(rowIndex + 1, 1, lastRow - rowIndex, 8).getValues();
        sheet.getRange(rowIndex, 1, lastRow - rowIndex, 8).setValues(values);
      }
      sheet.getRange(lastRow, 1, 1, 8).clearContent();
      return responseJSON({ status: "success" });
    }

    if (action === "add_member") {
      var sheet = getTaskSheet();
      var name = (params.name || "").trim();
      if (!name) return responseJSON({ status: "error", message: "Tên không được để trống" });
      var lastRow = sheet.getLastRow();
      if (lastRow >= 3) {
        var colI = sheet.getRange(3, 9, lastRow - 2, 1).getValues();
        for (var i = 0; i < colI.length; i++) {
          if (String(colI[i][0]).trim().toLowerCase() === name.toLowerCase()) {
            return responseJSON({ status: "error", message: "Thành viên đã tồn tại" });
          }
        }
      }
      var writeRow = Math.max(lastRow + 1, 3);
      if (lastRow >= 3) {
        var colIAll = sheet.getRange(3, 9, Math.max(lastRow - 2, 1), 1).getValues();
        writeRow = 3;
        for (var j = 0; j < colIAll.length; j++) {
          if (String(colIAll[j][0]).trim() !== "") writeRow = j + 4;
        }
      }
      sheet.getRange(writeRow, 9).setValue(name);
      return responseJSON({ status: "success" });
    }

    if (action === "delete_member") {
      var sheet = getTaskSheet();
      var name = (params.name || "").trim();
      if (!name) return responseJSON({ status: "error", message: "Tên không được để trống" });
      var lastRow = sheet.getLastRow();
      if (lastRow < 3) return responseJSON({ status: "error", message: "Không tìm thấy thành viên" });
      var colI = sheet.getRange(3, 9, lastRow - 2, 1).getValues();
      var found = false;
      for (var k = 0; k < colI.length; k++) {
        if (String(colI[k][0]).trim().toLowerCase() === name.toLowerCase()) {
          sheet.getRange(k + 3, 9).clearContent();
          found = true;
          break;
        }
      }
      if (!found) return responseJSON({ status: "error", message: "Không tìm thấy thành viên: " + name });
      return responseJSON({ status: "success" });
    }

    if (action === "save_url") {
      var sheet = getTaskSheet();
      var sysName = (params.name || "").trim();
      var sysUrl  = (params.url  || "").trim();
      if (!sysName) return responseJSON({ status: "error", message: "Thiếu tên hệ thống" });

      var lastRow = sheet.getLastRow();
      var targetRow = -1;

      if (lastRow >= 3) {
        var colK = sheet.getRange(3, 11, lastRow - 2, 1).getValues();
        for (var m = 0; m < colK.length; m++) {
          var existingName = String(colK[m][0] || "").trim().toLowerCase();
          if (existingName === sysName.toLowerCase()) {
            targetRow = m + 3;
            break;
          }
        }
      }

      if (targetRow === -1) {
        if (lastRow >= 3) {
          var colKAll = sheet.getRange(3, 11, Math.max(lastRow - 2, 1), 1).getValues();
          targetRow = 3;
          for (var n = 0; n < colKAll.length; n++) {
            if (String(colKAll[n][0]).trim() !== "") targetRow = n + 4;
          }
        } else {
          targetRow = 3;
        }
        sheet.getRange(targetRow, 11).setValue(sysName);
      }

      sheet.getRange(targetRow, 10).setValue(sysUrl);
      return responseJSON({ status: "success" });
    }

    // ── VĂN MẪU ACTIONS ──────────────────────────────────────────────────────
    if (action === "create") {
      var sheet = getVanMauSheet();
      sheet.appendRow(["", params.english, params.vietnamese, params.note || ""]);
      return responseJSON({ status: "success" });
    }

    if (action === "update") {
      var sheet = getVanMauSheet();
      var rowIndex = parseInt(params.id);
      if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
      sheet.getRange(rowIndex, 2).setValue(params.english);
      sheet.getRange(rowIndex, 3).setValue(params.vietnamese);
      sheet.getRange(rowIndex, 4).setValue(params.note || "");
      return responseJSON({ status: "success" });
    }

    if (action === "delete") {
      var sheet = getVanMauSheet();
      var rowIndex = parseInt(params.id);
      if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
      sheet.deleteRow(rowIndex);
      return responseJSON({ status: "success" });
    }

    return responseJSON({ status: "error", message: "Action không được hỗ trợ: " + action });

  } catch (err) {
    return responseJSON({ status: "error", message: "Lỗi thực thi Apps Script: " + err.toString() });
  }
}

// ─── Helper: response JSON ────────────────────────────────────────────────────
function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
