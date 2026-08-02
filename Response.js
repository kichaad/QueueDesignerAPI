/**
 * Формирование HTTP-ответов
 */
class Response {

  static success(data) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        data: data
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  static error(message) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

}