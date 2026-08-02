/**
 * Построение календарной очереди
 */
class CalendarService {

  /**
   * Строит календарь автоматически
   * по очередности и длительности
   */
  static buildCalendar(orders) {

    const rows = [];

    // Уже отсортированы OrderService,
    // но на всякий случай сортируем еще раз.
    orders.sort((a, b) => a.queue - b.queue);

    // Первый рабочий день
    let current = this.firstWorkDay();

    orders.forEach(order => {

      const duration = Number(order.duration) || 1;

      // Начало изделия
      const start = new Date(current);

      // Для каждого рабочего дня создаем строку календаря
      for (let i = 0; i < duration; i++) {

        rows.push({

          ...order,

          weekday: this.getWeekday(current),

          startDate: this.formatDate(current)

        });

        current = this.nextWorkDay(current);

      }

      // Заполняем расчетные даты
      order.startDate = this.formatDate(start);
      order.finishDate = this.formatDate(this.prevWorkDay(current));

    });

    return rows;

  }

  /**
   * Первый рабочий день
   */
  static firstWorkDay() {

    const d = new Date();

    while (d.getDay() === 0 || d.getDay() === 6) {
      d.setDate(d.getDate() + 1);
    }

    return d;

  }

  /**
   * Следующий рабочий день
   */
  static nextWorkDay(date) {

    const d = new Date(date);

    d.setDate(d.getDate() + 1);

    while (d.getDay() === 0 || d.getDay() === 6) {
      d.setDate(d.getDate() + 1);
    }

    return d;

  }

  /**
   * Предыдущий рабочий день
   */
  static prevWorkDay(date) {

    const d = new Date(date);

    d.setDate(d.getDate() - 1);

    while (d.getDay() === 0 || d.getDay() === 6) {
      d.setDate(d.getDate() - 1);
    }

    return d;

  }

  static formatDate(date) {

    return Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "dd.MM"
    );

  }

  static getWeekday(date) {

    return ["ВС","ПН","ВТ","СР","ЧТ","ПТ","СБ"][date.getDay()];

  }

}