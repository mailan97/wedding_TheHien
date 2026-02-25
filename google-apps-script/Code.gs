/**
 * Google Apps Script - Wedding RSVP Form Handler
 * ================================================
 * Script này nhận dữ liệu từ form trên website cưới
 * và tự động ghi vào Google Sheets.
 *
 * HƯỚNG DẪN CÀI ĐẶT:
 * 1. Mở Google Sheets → Extensions → Apps Script
 * 2. Xóa toàn bộ code mặc định, paste toàn bộ code này vào
 * 3. Nhấn "Deploy" → "New deployment"
 * 4. Chọn Type: "Web app"
 * 5. Cấu hình:
 *    - Description: "Wedding RSVP"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 6. Nhấn "Deploy" → Copy URL
 * 7. Paste URL vào biến scriptURL trong wedding.html
 */

// ===== CẤU HÌNH =====
// Lấy ID từ URL Google Sheets:
// https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
const SPREADSHEET_ID = "THAY_SPREADSHEET_ID_CUA_BAN_VAO_DAY";
const SHEET_NAME = "RSVP"; // Tên sheet tab chứa dữ liệu

/**
 * Xử lý request GET (test connection)
 */
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok", message: "Wedding RSVP API is running!" }),
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Xử lý request POST từ form wedding
 * Nhận dữ liệu JSON và ghi vào Google Sheets
 */
function doPost(e) {
  try {
    // Parse dữ liệu từ request
    const data = JSON.parse(e.postData.contents);

    // Mở spreadsheet bằng ID
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Nếu sheet chưa tồn tại, tạo mới với header
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // Tạo header row
      const headers = [
        "Thời gian",
        "Họ và tên",
        "Lời chúc",
        "Tham dự",
        "Số khách",
        "Phía",
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // Format header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4C6841"); // Màu xanh lá theo theme cưới
      headerRange.setFontColor("#FFFFFF");
      headerRange.setHorizontalAlignment("center");

      // Set độ rộng cột
      sheet.setColumnWidth(1, 180); // Thời gian
      sheet.setColumnWidth(2, 200); // Họ và tên
      sheet.setColumnWidth(3, 350); // Lời chúc
      sheet.setColumnWidth(4, 120); // Tham dự
      sheet.setColumnWidth(5, 100); // Số khách
      sheet.setColumnWidth(6, 120); // Phía

      // Freeze header row
      sheet.setFrozenRows(1);
    }

    // Tạo timestamp theo múi giờ Việt Nam
    const timestamp = Utilities.formatDate(
      new Date(),
      "Asia/Ho_Chi_Minh",
      "dd/MM/yyyy HH:mm:ss",
    );

    // Chuẩn bị dữ liệu để ghi vào sheet
    const rowData = [
      timestamp,
      data.name || "",
      data.message || "",
      data.attend || "",
      data.guests || "",
      data.side || "",
    ];

    // Ghi dữ liệu vào dòng tiếp theo
    sheet.appendRow(rowData);

    // Trả về kết quả thành công
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Đã lưu thông tin thành công!",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // Trả về lỗi nếu có
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Hàm test - chạy thử để kiểm tra
 * Chạy hàm này trong Apps Script Editor để test
 */
function testDoPost() {
  const testData = {
    postData: {
      contents: JSON.stringify({
        name: "Nguyễn Văn Test",
        message: "Chúc mừng hạnh phúc!",
        attend: "Có",
        guests: "2",
        side: "Nhà gái",
      }),
    },
  };

  const result = doPost(testData);
  Logger.log(result.getContent());
}
