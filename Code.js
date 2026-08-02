function doGet(e) {
  // Если параметр resource существует
  if (e && e.parameter && e.parameter.resource) {
    return Api.handleGet(e);
  }

  // Иначе открываем приложение
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Очередь конструкторов')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return Api.handlePost(e);
}

function testOrders() {
  const orders = OrderService.getOrders();
  Logger.log(JSON.stringify(orders, null, 2));
}

/**
 * Возвращает список изделий для интерфейса
 */
function getOrders() {
  return OrderService.getOrders();
}

/**
 * Подключение HTML-фрагментов в шаблон Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Сохраняет новую стадию заказа по номеру договора
 */
function saveStage(contractNumber, newStage) {
  try {
    const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Изделия');
    
    if (!sheet) throw new Error('Лист "Изделия" не найден');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const contractIndex = headers.indexOf('Номер договора');
    const stageIndex = headers.indexOf('Стадия');

    if (contractIndex === -1 || stageIndex === -1) {
      throw new Error('Не найдены колонки "Номер договора" или "Стадия"');
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][contractIndex]) === String(contractNumber)) {
        sheet.getRange(i + 1, stageIndex + 1).setValue(newStage);
        return { success: true, message: 'Стадия обновлена' };
      }
    }

    throw new Error(`Заказ "${contractNumber}" не найден`);

  } catch (e) {
    Logger.log(`Ошибка saveStage: ${e.message}`);
    throw e;
  }
}