/**
 * Сервис работы с очередью изделий
 */
class OrderService {

  /**
   * Возвращает очередь изделий
   */
  static getOrders() {

    const orders = SheetService
      .getObjects(CONFIG.SHEETS.ORDERS)
      .map(row => OrderMapper.map(row))
      .sort((a, b) => a.queue - b.queue);

    // Расчет календаря
    const calendar = CalendarService.buildCalendar(orders);

    // Сохранение рассчитанных дат
    SheetService.saveCalendar(orders);

    // Возврат календаря клиенту
    return JSON.parse(JSON.stringify(calendar));

  }

}