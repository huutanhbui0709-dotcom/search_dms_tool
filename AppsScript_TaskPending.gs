/**
 * GOOGLE APPS SCRIPT BACKEND – TASK PENDING NVN (Chỉ chuyên xử lý Tasks & Members)
 * Spreadsheet ID: 1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g
 * Tên sheet: "Công việc Pending"
 *   - Tasks: Cột A-H (hàng 2+)
 *   - Members: Cột I (I3:I)
 */

var SPREADSHEET_ID = "1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g";
var TASK_SHEET     = "Công việc Pending";

function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TASK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TASK_SHEET);
    var headers = ["ID", "Hạn cuối", "Issue", "Ghi chú", "Người phụ trách", "Trạng thái", "", "", "Thành viên"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1e1b4b")
      .setFontColor("#ffffff");
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

  // ── Task Pending: đọc danh sách task ─────────────────────────────────────
  if (action === "read_tasks") {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();
    var rows  = [];
    if (data.length > 1) {
      data.slice(1).forEach(function (row, idx) {
        // Chỉ lấy dòng có nội dung Issue hoặc Ngày hạn cuối
        if (!row[2] && !row[1]) return;
        var dateStr = "";
        if (row[1]) {
          try {
            if (row[1] instanceof Date) {
              dateStr = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd");
            } else {
              dateStr = String(row[1]).split("T")[0];
            }
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
          assignees: row[4] ? String(row[4]).split(",").map(function (s) { return s.trim(); }).filter(Boolean) : [],
          status:    row[5] || "Pending"
        });
      });
    }
    return responseJSON({ status: "success", data: rows, tasks: rows });
  }

  // ── Task Pending: đọc danh sách thành viên từ Cột I (I3:I) ──────────────
  if (action === "read_members") {
    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();
    var members = [];

    if (lastRow >= 3) {
      var colI = sheet.getRange(3, 9, lastRow - 2, 1).getValues();
      colI.forEach(function (row, idx) {
        var name = String(row[0] || "").trim();
        if (name) {
          members.push({ rowIndex: idx + 3, name: name });
        }
      });
    }
    return responseJSON({ status: "success", data: members });
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
  try {
    var sheet = getSheet();

    if (action === "save_task") {
      var assigneesStr = Array.isArray(params.assignees)
        ? params.assignees.join(", ")
        : (params.assignees || "");
      var status       = params.status || "Pending";
      var deadline     = params.deadline || params.dueDate || "";

      if (params.id) {
        // UPDATE – ghi đè dòng có id
        var rowIndex = parseInt(params.id);
        if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
        sheet.getRange(rowIndex, 2).setValue(deadline);
        sheet.getRange(rowIndex, 3).setValue(params.issue || "");
        sheet.getRange(rowIndex, 4).setValue(params.note || "");
        sheet.getRange(rowIndex, 5).setValue(assigneesStr);
        sheet.getRange(rowIndex, 6).setValue(status);
      } else {
        // CREATE – thêm dòng mới
        sheet.appendRow(["", deadline, params.issue || "", params.note || "", assigneesStr, status]);
      }
      return responseJSON({ status: "success" });
    }

    if (action === "delete_task") {
      var rowIndex = parseInt(params.id);
      if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
      sheet.deleteRow(rowIndex);
      return responseJSON({ status: "success" });
    }

    // ── MEMBER MANAGEMENT – Cột I (I3:I) của sheet "Công việc Pending" ───────
    if (action === "add_member") {
      var name = (params.name || "").trim();
      if (!name) return responseJSON({ status: "error", message: "Tên không được để trống" });

      var lastRow = sheet.getLastRow();

      // Kiểm tra trùng tên trong cột I
      if (lastRow >= 3) {
        var colI = sheet.getRange(3, 9, lastRow - 2, 1).getValues();
        for (var i = 0; i < colI.length; i++) {
          if (String(colI[i][0]).trim().toLowerCase() === name.toLowerCase()) {
            return responseJSON({ status: "error", message: "Thành viên đã tồn tại" });
          }
        }
      }

      // Tìm ô trống tiếp theo trong cột I (bắt đầu từ I3)
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

    return responseJSON({ status: "error", message: "Action không được hỗ trợ: " + action });
  } catch (err) {
    return responseJSON({ status: "error", message: "Lỗi thực thi Apps Script: " + err.toString() });
  }
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
