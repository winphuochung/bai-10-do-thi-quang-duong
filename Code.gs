/**
 * Google Apps Script - Nhận kết quả học tập Bài 10: Đồ thị quãng đường - thời gian
 * Hướng dẫn: Dán toàn bộ mã này vào trình soạn thảo Apps Script của Google Sheets,
 * sau đó chọn Triển khai (Deploy) -> Tùy chọn triển khai mới (New deployment) -> Ứng dụng web (Web app).
 * - Thực thi dưới dạng (Execute as): Tôi (Me)
 * - Người có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 */

const SHEET_NAME = "KetQua";

/**
 * Xử lý yêu cầu GET - Dùng để kiểm tra link Web App có hoạt động không
 */
function doGet(e) {
  return ContentService
    .createTextOutput("Kết nối Google Sheets cho Bài 10 (Đồ thị quãng đường - thời gian) đang hoạt động bình thường!")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Xử lý yêu cầu POST - Nhận dữ liệu JSON từ web học tập và ghi vào tab "KetQua"
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // Tự động tạo tab "KetQua" nếu chưa có
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Tự tạo dòng tiêu đề nếu Sheet còn trống
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Thời gian",
        "Họ tên học sinh",
        "Lớp",
        "Điểm (/10)",
        "Số câu đúng",
        "Chi tiết đáp án",
        "Thời gian làm bài"
      ];
      sheet.appendRow(headers);

      // Định dạng nhẹ dòng tiêu đề cho dễ nhìn
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#0070C0");
      headerRange.setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }

    // Parse dữ liệu JSON gửi lên (gửi dạng text/plain để tránh lỗi CORS)
    var rawData = e.postData && e.postData.contents ? e.postData.contents : "{}";
    var data = JSON.parse(rawData);

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var hoTen = data.hoTen || "";
    var lop = data.lop || "";
    var diem = data.diem !== undefined ? data.diem : "";
    var soCauDung = data.soCauDung || "";
    var chiTiet = data.chiTiet || "";
    var thoiGianLamBai = data.thoiGianLamBai || "";

    sheet.appendRow([
      timestamp,
      hoTen,
      lop,
      diem,
      soCauDung,
      chiTiet,
      thoiGianLamBai
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", message: "Đã lưu kết quả thành công" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
