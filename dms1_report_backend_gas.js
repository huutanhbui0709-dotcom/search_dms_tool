/**
 * Google Apps Script Backend for DMS1 Report (CRUD for .xlsx file on Google Drive)
 * 
 * HƯỚNG DẪN CẤU HÌNH:
 * 1. Bật Dịch vụ: Services (+) -> Drive API -> Version: v2 -> Identifier: Drive -> Save (như hình bạn vừa mở)
 * 2. Thay FILE_ID bên dưới bằng ID file 00.DMS1-Report.xlsx trên Google Drive của bạn
 * 3. Deploy (Triển khai):
 *    - Chọn: Web App
 *    - Execute as (Thực thi dưới dạng): Me (Tôi)
 *    - Who has access (Ai có quyền truy cập): Anyone (Bất kỳ ai)
 *    - Chọn "New version" mỗi khi chỉnh sửa code
 */

const FILE_ID = "1MLI9GNQ-FDeWLwdIOPuPV40VAvzabd_6"; // ID File Excel trên Google Drive
const SHEET_NAME_OR_INDEX = 0; // Tên sheet (ví dụ: "Sheet1") hoặc số thứ tự (0 là sheet đầu tiên)

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = (params.action || "read").toLowerCase();

  try {
    if (!FILE_ID || FILE_ID === "YOUR_EXCEL_FILE_ID_HERE") {
      return createJsonResponse({
        status: "error",
        message: "Vui lòng điền đúng FILE_ID của file Excel trên Google Drive!"
      });
    }

    // 1. Lấy file Excel gốc và chuyển sang Google Sheet tạm
    const excelFile = DriveApp.getFileById(FILE_ID);
    const excelName = excelFile.getName();
    const excelBlob = excelFile.getBlob();

    const tempResource = {
      title: "Temp_DMS1_" + Date.now(),
      mimeType: "application/vnd.google-apps.spreadsheet"
    };

    // Tạo file Google Sheet tạm qua Drive API v2
    const tempFile = Drive.Files.insert(tempResource, excelBlob, { convert: true });
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    
    // Tìm sheet làm việc
    let sheet;
    if (typeof SHEET_NAME_OR_INDEX === "string") {
      sheet = tempSpreadsheet.getSheetByName(SHEET_NAME_OR_INDEX) || tempSpreadsheet.getSheets()[0];
    } else {
      sheet = tempSpreadsheet.getSheets()[SHEET_NAME_OR_INDEX || 0];
    }

    let isModified = false;

    // 2. Thực hiện thao tác CRUD
    if (action === "add") {
      const lastRow = sheet.getLastRow();
      const nextNo = params.colA || params.no || (lastRow >= 2 ? lastRow : 1);
      const colB = params.colB || params.nameVi || "";
      const colC = params.colC || params.nameEn || "";
      const colD = params.colD || params.path || "";
      const colE = params.colE || params.meaning || "";

      sheet.appendRow([nextNo, colB, colC, colD, colE]);
      isModified = true;

    } else if (action === "edit") {
      const rowIndex = parseInt(params.rowIndex || params.row, 10);
      if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
        throw new Error("Vị trí dòng (rowIndex=" + params.rowIndex + ") không hợp lệ!");
      }

      const colA = params.colA !== undefined ? params.colA : sheet.getRange(rowIndex, 1).getValue();
      const colB = params.colB !== undefined ? params.colB : sheet.getRange(rowIndex, 2).getValue();
      const colC = params.colC !== undefined ? params.colC : sheet.getRange(rowIndex, 3).getValue();
      const colD = params.colD !== undefined ? params.colD : sheet.getRange(rowIndex, 4).getValue();
      const colE = params.colE !== undefined ? params.colE : sheet.getRange(rowIndex, 5).getValue();

      sheet.getRange(rowIndex, 1, 1, 5).setValues([[colA, colB, colC, colD, colE]]);
      isModified = true;

    } else if (action === "delete") {
      const rowIndex = parseInt(params.rowIndex || params.row, 10);
      if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
        throw new Error("Vị trí dòng (rowIndex=" + params.rowIndex + ") không hợp lệ!");
      }

      sheet.deleteRow(rowIndex);
      isModified = true;
    }

    // 3. Nếu có chỉnh sửa: Xuất lại XLSX và ghi đè file gốc
    if (isModified) {
      SpreadsheetApp.flush();
      Utilities.sleep(600);

      // Xuất temp spreadsheet ra blob XLSX
      const exportUrl = "https://docs.google.com/spreadsheets/d/" + tempFile.id + "/export?format=xlsx";
      const exportResponse = UrlFetchApp.fetch(exportUrl, {
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken()
        },
        muteHttpExceptions: true
      });

      if (exportResponse.getResponseCode() !== 200) {
        throw new Error("Không thể xuất XLSX từ Google Sheet tạm: HTTP " + exportResponse.getResponseCode());
      }

      const updatedBlob = exportResponse.getBlob();
      updatedBlob.setName(excelName);
      updatedBlob.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

      // Ghi đè file Excel gốc trên Google Drive
      Drive.Files.update(
        { title: excelName, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        FILE_ID,
        updatedBlob
      );
    }

    // 4. Đọc toàn bộ danh sách để trả về frontend
    const lastRow = sheet.getLastRow();
    const rows = [];
    if (lastRow >= 2) {
      const dataValues = sheet.getRange(1, 1, lastRow, 5).getValues();
      for (let i = 1; i < dataValues.length; i++) {
        const row = dataValues[i];
        if (row.every(cell => String(cell).trim() === "")) continue;

        rows.push({
          rowIndex: i + 1,
          "No": row[0],
          "Tên báo cáo (Tiếng Việt)": row[1],
          "Tên báo cáo (Tiếng Anh)": row[2],
          "Đường dẫn": row[3],
          "Ý nghĩa báo cáo": row[4]
        });
      }
    }

    // 5. Xóa file tạm để giải phóng dung lượng Drive
    try {
      Drive.Files.remove(tempFile.id);
    } catch (cleanErr) {
      console.warn("Lỗi xóa file tạm: " + cleanErr.message);
    }

    return createJsonResponse({
      status: "success",
      action: action,
      total: rows.length,
      data: rows
    });

  } catch (err) {
    console.error("Lỗi DMS1 Report:", err);
    return createJsonResponse({
      status: "error",
      message: err.message || String(err)
    });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
