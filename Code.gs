/**
 * Web app tra cứu sản phẩm theo Ship / Code.
 * Toàn bộ xử lý Excel chạy ở trình duyệt (trong Index.html);
 * phía server chỉ phục vụ trang và (tuỳ chọn) lưu file kết quả vào Google Drive.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tra cứu sản phẩm theo Ship / Code')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lưu file Excel (base64) vào thư mục gốc của Google Drive người dùng. */
function saveToDrive(base64, fileName) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    fileName
  );
  const file = DriveApp.createFile(blob);
  return { url: file.getUrl(), name: file.getName() };
}
