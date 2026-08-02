/**
 * Сервис работы с Google Sheets
 */
class SheetService {

  /**
   * Возвращает лист по имени
   */
  static getSheet(sheetName) {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    return spreadsheet.getSheetByName(sheetName);
  }

  /**
   * Возвращает все строки листа
   */
  static getData(sheetName) {
    const sheet = this.getSheet(sheetName);
    return sheet.getDataRange().getValues();
  }

  /**
   * Возвращает данные в виде массива объектов
   */
  static getObjects(sheetName) {

    const rows = this.getData(sheetName);

    if (rows.length < 2) {
      return [];
    }

    const headers = rows[0];
    const result = [];

    for (let i = 1; i < rows.length; i++) {

      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = rows[i][index];
      });

      result.push(obj);
    }

    return result;
  }
/**
 * Сохраняет рассчитанные даты в таблицу
 */
static saveCalendar(orders) {

  const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return;

  // H = Начало
  // J = Окончание
  const startColumn = 8;
  const finishColumn = 10;

  for (let i = 0; i < orders.length; i++) {

    const row = i + 2;

    sheet.getRange(row, startColumn).setValue(
      orders[i].startDate || ""
    );

    sheet.getRange(row, finishColumn).setValue(
      orders[i].finishDate || ""
    );

  }

}
}