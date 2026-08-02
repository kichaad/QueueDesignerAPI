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