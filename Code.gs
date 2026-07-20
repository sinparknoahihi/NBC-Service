/**
 * CODE.GS — Nhận dữ liệu Form từ Landing Page NBC/GS1 Vietnam về Google Sheets
 *
 * CÁCH DÙNG:
 * 1. Mở Google Sheet mới (hoặc sheet bạn muốn lưu dữ liệu).
 * 2. Vào menu: Tiện ích mở rộng (Extensions) > Apps Script.
 * 3. Xoá code mẫu, dán toàn bộ nội dung file này vào.
 * 4. Bấm "Triển khai" (Deploy) > "Triển khai mới" (New deployment).
 *    - Loại: "Ứng dụng web" (Web app)
 *    - Execute as: Me
 *    - Who has access: Anyone (bắt buộc để landing page gọi được)
 * 5. Copy URL "Web app" được cấp, dán vào biến CONFIG.GOOGLE_SCRIPT_URL
 *    trong file index.html.
 */

const SHEET_NAME = "Leads"; // Đổi tên tab sheet nếu cần

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),                     // Thời gian nhận
      data.companyName || "",
      data.phone || "",
      data.email || "",
      data.needTruyXuatNguonGoc ? "Có" : "",
      data.txngQty || "",
      data.txngProcess || "",
      data.needElabel ? "Có" : "",
      data.elabelQty || "",
      data.elabelInfo || "",
      data.needTemChongGia ? "Có" : "",
      data.temgiaQty || "",
      data.temgiaType || "",
      data.note || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Thời gian", "Tên doanh nghiệp", "SĐT", "Email",
      "Cần Truy xuất nguồn gốc", "SL SP (TXNG)", "Có quy trình SX/chuỗi CU?",
      "Cần Nhãn điện tử", "SL SP (E-label)", "Thông tin hiển thị (E-label)",
      "Cần Tem chống giả", "SL tem dự kiến", "Loại xác thực",
      "Ghi chú"
    ]);
    sheet.getRange(1, 1, 1, 14).setFontWeight("bold");
  }
  return sheet;
}
