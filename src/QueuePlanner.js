/**
 * Планировщик очереди
 */
class QueuePlanner {

  /**
   * Строит календарь и рассчитывает даты
   */
  static build(orders) {

    let current = this.nextWorkDay(new Date());

    return orders.map(order => {

      const duration = Number(order.duration) || 1;

      order.startDate = new Date(current);

      let finish = new Date(current);

      for (let i = 1; i < duration; i++) {
        finish = this.nextWorkDay(finish);
      }

      order.finishDate = finish;

      current = this.nextWorkDay(finish);

      return order;

    });

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

}