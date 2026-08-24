/**
 * GOOGLE APPS SCRIPT BACKEND – TASK PENDING NVN
 * Spreadsheet ID: 1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g
 * Tên sheet: "Công việc Pending"
 *   - Tasks   : Cột A-H (hàng 2+)  – lấy theo điều kiện có Issue hoặc Ngày
 *   - Members : Cột I  (I3:I)       – danh sách thành viên
 *   - URL     : Cột J  (J3:J)       – URL của từng hệ thống
 *   - Tên HT  : Cột K  (K3:K)       – Tên hệ thống (Solution, Calllog, ...)
 */

var SPREADSHEET_ID = "1ODSyKIHRqFVGsnCnKJYgEWYE6zI52ooZvwRX5rnvF1g";
var TASK_SHEET     = "Công việc Pending";

// ─── Helper: lấy sheet, tự tạo nếu chưa có ────────────────────────────────────
function getSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
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

// ─── Helper: đọc danh sách URL hệ thống từ Cột J & K ──────────────────────────
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

// ─── GET Router ───────────────────────────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action;

  // Fallback GET Bypass CORS: nếu có tham số data thì chuyển sang POST handler
  if (e.parameter.data) {
    var params;
    try {
      params = JSON.parse(decodeURIComponent(e.parameter.data));
    } catch (err) {
      return responseJSON({ status: "error", message: "Dữ liệu JSON fallback không hợp lệ" });
    }
    return handlePostActions(action, params);
  }

  // ── Đọc danh sách Task ───────────────────────────────────────────────────
  if (action === "read_tasks") {
    var sheet      = getSheet();
    var data       = sheet.getDataRange().getValues();
    var rows       = [];
    var systemUrls = readSystemUrls(sheet);

    if (data.length > 1) {
      data.slice(1).forEach(function (row, idx) {
        if (!row[2] && !row[1]) return;           // bỏ dòng trống
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

    // Tìm URL của "Task pending" để trả về riêng để Frontend tự cập nhật
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

  // ── Đọc danh sách Members từ Cột I ──────────────────────────────────────
  if (action === "read_members") {
    var sheet   = getSheet();
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

  // ── Đọc danh sách System URLs ────────────────────────────────────────────
  if (action === "read_system_urls") {
    var sheet      = getSheet();
    var systemUrls = readSystemUrls(sheet);
    return responseJSON({ status: "success", data: systemUrls });
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
    var sheet = getSheet();

    // ── Lưu / Cập nhật Task ─────────────────────────────────────────────────
    if (action === "save_task") {
      var assigneesStr = Array.isArray(params.assignees)
        ? params.assignees.join(", ")
        : (params.assignees || "");
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
        sheet.appendRow(["", deadline, params.issue || "", params.note || "", assigneesStr, status]);
      }
      return responseJSON({ status: "success" });
    }

    // ── Xóa Task ────────────────────────────────────────────────────────────
    if (action === "delete_task") {
      var rowIndex = parseInt(params.id);
      if (!rowIndex || rowIndex < 2) return responseJSON({ status: "error", message: "ID không hợp lệ" });
      sheet.deleteRow(rowIndex);
      return responseJSON({ status: "success" });
    }

    // ── Thêm Thành viên vào Cột I ───────────────────────────────────────────
    if (action === "add_member") {
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

    // ── Xóa Thành viên khỏi Cột I ──────────────────────────────────────────
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

    // ── Lưu / Cập nhật URL Hệ thống vào Cột J & K ─────────────────────────
    if (action === "save_url") {
      var sysName = (params.name || "").trim();
      var sysUrl  = (params.url  || "").trim();
      if (!sysName) return responseJSON({ status: "error", message: "Thiếu tên hệ thống" });

      var lastRow = sheet.getLastRow();
      var targetRow = -1;

      // Tìm dòng có tên hệ thống khớp ở Cột K
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
        // Tìm ô trống đầu tiên trong cột K từ K3
        if (lastRow >= 3) {
          var colKAll = sheet.getRange(3, 11, Math.max(lastRow - 2, 1), 1).getValues();
          targetRow = 3;
          for (var n = 0; n < colKAll.length; n++) {
            if (String(colKAll[n][0]).trim() !== "") targetRow = n + 4;
          }
        } else {
          targetRow = 3;
        }
        // Ghi tên hệ thống vào Cột K
        sheet.getRange(targetRow, 11).setValue(sysName);
      }

      // Ghi URL vào Cột J
      sheet.getRange(targetRow, 10).setValue(sysUrl);
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
