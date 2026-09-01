/**
 * Google Apps Script Backend for DMS1 Report (CRUD for .xlsx file on Google Drive)
 *
 * Instructions:
 * 1. Open Google Drive -> create a new Google Apps Script project.
 * 2. Enable Advanced Google Service: Services (+) -> Drive API (v2 or v3).
 * 3. Set the FILE_ID below to your 00.DMS1-Report.xlsx Google Drive File ID.
 * 4. Deploy -> New deployment -> Select type: Web App.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy Web App URL and paste into DMS1 Report frontend.
 */

const FILE_ID = "YOUR_EXCEL_FILE_ID_HERE"; // Thay thế bằng ID file 00.DMS1-Report.xlsx trên Google Drive

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = (params.action || "read").toLowerCase();

  try {
    if (!FILE_ID || FILE_ID === "YOUR_EXCEL_FILE_ID_HERE") {
      return createJsonResponse({
        status: "error",
        message: "Vui lòng cấu hình FILE_ID của file Excel trên Google Drive trong Apps Script!"
      });
    }

    // 1. Convert .xlsx file to temporary Google Spreadsheet
    const excelFile = DriveApp.getFileById(FILE_ID);
    const blob = excelFile.getBlob();
    
    // Create temp Google Sheet using Drive API
    const tempFileResource = {
      title: "Temp_DMS1_Report_" + Date.now(),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
    // Drive API v2
    const tempFile = Drive.Files.insert(tempFileResource, blob, { convert: true });
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const sheet = tempSpreadsheet.getSheets()[0];
    
    let isModified = false;

    // 2. Perform requested CRUD operation
    if (action === "add") {
      // Calculate next STT / No if not given
      const lastRow = sheet.getLastRow();
      const nextNo = params.colA || params.no || (lastRow >= 2 ? (lastRow) : 1);
      const colB = params.colB || params.nameVi || "";
      const colC = params.colC || params.nameEn || "";
      const colD = params.colD || params.path || "";
      const colE = params.colE || params.meaning || "";

      sheet.appendRow([nextNo, colB, colC, colD, colE]);
      isModified = true;

    } else if (action === "edit") {
      const rowIndex = parseInt(params.rowIndex || params.row, 10);
      if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
        throw new Error("rowIndex không hợp lệ: " + params.rowIndex);
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
        throw new Error("rowIndex không hợp lệ: " + params.rowIndex);
      }

      sheet.deleteRow(rowIndex);
      isModified = true;
    }

    // 3. If modified, export back to .xlsx format and overwrite original file on Drive
    if (isModified) {
      SpreadsheetApp.flush();
      Utilities.sleep(500);

      // Export temp spreadsheet as XLSX using OAuth Token
      const url = "https://docs.google.com/spreadsheets/d/" + tempFile.id + "/export?format=xlsx";
      const response = UrlFetchApp.fetch(url, {
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken()
        },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        throw new Error("Xuất file XLSX thất bại: HTTP " + response.getResponseCode());
      }

      const updatedBlob = response.getBlob().setName(excelFile.getName());

      // Overwrite original .xlsx file using Drive API v2
      Drive.Files.update({ mimeType: MimeType.MICROSOFT_EXCEL }, FILE_ID, updatedBlob);
    }

    // 4. Read all data rows from temporary spreadsheet to return updated JSON
    const dataValues = sheet.getDataRange().getValues();
    const headers = dataValues[0] || ["No", "Tên báo cáo (Tiếng Việt)", "Tên báo cáo (Tiếng Anh)", "Đường dẫn", "Ý nghĩa báo cáo"];
    const rows = [];

    for (let i = 1; i < dataValues.length; i++) {
      const row = dataValues[i];
      // Skip completely empty rows
      if (row.every(cell => String(cell).trim() === "")) continue;

      rows.push({
        rowIndex: i + 1, // 1-based index in sheet
        "No": row[0],
        "Tên báo cáo (Tiếng Việt)": row[1],
        "Tên báo cáo (Tiếng Anh)": row[2],
        "Đường dẫn": row[3],
        "Ý nghĩa báo cáo": row[4]
      });
    }

    // 5. Clean up temporary Google Sheet
    try {
      Drive.Files.remove(tempFile.id);
    } catch (cleanErr) {
      console.warn("Lỗi khi xóa file tạm: " + cleanErr.message);
    }

    return createJsonResponse({
      status: "success",
      action: action,
      total: rows.length,
      data: rows
    });

  } catch (err) {
    console.error("DMS1 Report Error:", err);
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
