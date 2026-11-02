/**
 * Построение календарной очереди
 */
class CalendarService {

  /**
   * Строит календарь автоматически
   * по очередности и длительности
   */
  static buildCalendar(orders, options) {

    const rows = [];
    const emitRows = !options || options.emitRows !== false;
    const firstWorkDay = this.firstWorkDay();
    const formattedDates = {};
    const formatDate = date => {
      const key = new Date(date).setHours(0, 0, 0, 0);
      if (!formattedDates[key]) formattedDates[key] = this.formatDate(date);
      return formattedDates[key];
    };

    // Уже отсортированы OrderService,
    // но на всякий случай сортируем еще раз.
    orders.sort((a, b) => a.queue - b.queue);

    // Первый рабочий день
    // Разделяем очередь по конструкторам
const queues = {};

orders.forEach(order => {

  const constructor = order.constructor || "Без конструктора";

  if (!queues[constructor]) {
    queues[constructor] = [];
  }

  queues[constructor].push(order);

});


Object.values(queues).forEach(queue => {
  let current = new Date(firstWorkDay);
  let halfDay = 0;
  const calendar = this;
  const byQueue = (left, right) => Number(left.queue || 0) - Number(right.queue || 0);
  const isActive = order => String(order.status || "").trim().toLowerCase() === "в работе";
  const isConstructorTask = order => {
    const stage = String(order.stage || "").trim().toLowerCase();
    return ["proforma", "kd", "kd-reserve"].includes(order.taskType) ||
      ["проформа", "кд", "согласование"].includes(stage);
  };
  const getHalfDays = order => {
    const duration = isConstructorTask(order) && Number(order.remainingDuration) > 0
      ? Number(order.remainingDuration)
      : Number(order.duration) || 1;
    return Math.max(1, Math.round(duration * 2));
  };

  const canFinishBefore = (order, targetDate) => {
    let probeDate = new Date(current);
    let probeHalfDay = halfDay;
    const delayedStart = calendar.parseStoredDate(order.plannedDate);
    if (!order.manualPriority && delayedStart && delayedStart > probeDate) {
      probeDate = new Date(delayedStart);
      probeHalfDay = 0;
    }
    for (let index = 0; index < getHalfDays(order); index++) {
      probeHalfDay += 1;
      if (probeHalfDay === 2) {
        probeDate = calendar.nextWorkDay(probeDate);
        probeHalfDay = 0;
      }
    }
    return probeDate <= targetDate;
  };

  const schedule = order => {
    // Согласование у конструктора — это ожидающее решения менеджера действие,
    // а не работа, занимающая производственную мощность.
    if (order.displayOnly) {
      const displayDate = new Date(firstWorkDay);
      const row = Object.assign({}, order, {
        weekday: calendar.getWeekday(displayDate),
        startDate: formatDate(displayDate),
        finishDate: formatDate(displayDate),
        dayPart: ""
      });
      if (emitRows) rows.push(row);
      order.startDate = row.startDate;
      order.finishDate = row.finishDate;
      order.scheduleStart = new Date(displayDate);
      order.scheduleFinish = new Date(displayDate);
      order.dayPart = "";
      return;
    }

    // Для конструктора фактический старт не должен зависеть от последующих
    // пересчётов поля «Начало»: он хранится отдельно вместе с половиной дня.
    const actualStart = isActive(order) && isConstructorTask(order) && calendar.parseStoredDate(
      order.actualStartDate || order.startDate
    );
    if (actualStart) {
      current = new Date(actualStart);
      halfDay = calendar.normalizeHalfDay(order.actualStartHalf);
    }

    const plannedStart = calendar.parseStoredDate(order.plannedDate);
    // Обычная плановая дата (например, пауза) сохраняет прежнее поведение.
    // Ручной приоритет назначается внешним алгоритмом точно на заданную дату.
    if (!order.manualPriority && plannedStart && plannedStart > current) {
      current = new Date(plannedStart);
      halfDay = 0;
    }

    const start = new Date(current);
    const startDayPart = halfDay === 0 ? "Ⅰ половина" : "Ⅱ половина";
    let finish = new Date(current);
    const usedDates = emitRows ? {} : null;
    for (let index = 0; index < getHalfDays(order); index++) {
      finish = new Date(current);
      const dateKey = emitRows ? formatDate(current) : "";
      if (emitRows && !usedDates[dateKey]) {
        const row = Object.assign({}, order, {
          weekday: calendar.getWeekday(current),
          startDate: dateKey,
          dayPart: startDayPart
        });
        rows.push(row);
        usedDates[dateKey] = true;
      }
      halfDay += 1;
      if (halfDay === 2) {
        current = calendar.nextWorkDay(current);
        halfDay = 0;
      }
    }
    order.startDate = formatDate(start);
    order.finishDate = formatDate(finish);
    order.scheduleStart = start;
    order.scheduleFinish = finish;
    order.dayPart = startDayPart;

    // Незавершённая работа не освобождает ресурс раньше текущей половины дня,
    // даже если её первоначальная оценка уже закончилась.
    if (isActive(order) && isConstructorTask(order)) {
      const minimumAvailable = calendar.getCurrentAvailablePosition();
      if (calendar.isPositionBefore(current, halfDay, minimumAvailable.date, minimumAvailable.halfDay)) {
        current = minimumAvailable.date;
        halfDay = minimumAvailable.halfDay;
      }
    }
  };

  const ordered = queue.slice().sort(byQueue);
  const displayOnly = ordered.filter(order => order.displayOnly);
  const active = ordered.filter(order => !order.displayOnly && isActive(order));
  const manual = ordered.filter(order => !order.displayOnly && !isActive(order) && order.manualPriority);
  const regular = ordered.filter(order => !order.displayOnly && !isActive(order) && !order.manualPriority);

  // Незавершённая фактическая работа никогда не прерывается вручную заданной датой.
  active.forEach(order => {
    const previousDate = new Date(current);
    const previousHalfDay = halfDay;
    schedule(order);
    // Несколько активных изделий могут иметь разные фактические даты старта.
    // Последующая очередь начинается после самого позднего окончания.
    if (calendar.isPositionBefore(current, halfDay, previousDate, previousHalfDay)) {
      current = previousDate;
      halfDay = previousHalfDay;
    }
  });

  // Ручная дата имеет приоритет над обычной очередью, но обычные задачи успевают
  // выполняться до неё, если полностью помещаются до начала выбранного дня.
  while (manual.length) {
    manual.sort((left, right) => {
      const leftDate = calendar.parseStoredDate(left.plannedDate) || current;
      const rightDate = calendar.parseStoredDate(right.plannedDate) || current;
      return leftDate - rightDate || byQueue(left, right);
    });
    const nextManual = manual[0];
    const targetDate = calendar.parseStoredDate(nextManual.plannedDate) || new Date(current);

    if (targetDate <= current) {
      manual.shift();
      schedule(nextManual);
      continue;
    }

    if (regular.length && canFinishBefore(regular[0], targetDate)) {
      schedule(regular.shift());
      continue;
    }

    current = new Date(targetDate);
    halfDay = 0;
    manual.shift();
    schedule(nextManual);
  }

  regular.forEach(schedule);
  displayOnly.forEach(schedule);
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
      "dd.MM.yyyy"
    );

  }

  static getWeekday(date) {

    return ["ВС","ПН","ВТ","СР","ЧТ","ПТ","СБ"][date.getDay()];

  }

  static normalizeHalfDay(value) {
    const normalized = String(value || "").trim();
    return Number(value) === 1 || normalized === "Ⅱ" || normalized === "2" || normalized.toLowerCase() === "ii" ? 1 : 0;
  }

  static getCurrentHalfDay() {
    const hour = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "H"));
    return hour >= 14 ? 1 : 0;
  }

  // Текущая половина дня ещё занята незавершённой работой. Возвращаем первую
  // точку, когда следующий заказ может начать работу.
  static getCurrentAvailablePosition() {
    const today = this.firstWorkDay();
    if (this.getCurrentHalfDay() === 0) {
      return { date: today, halfDay: 1 };
    }
    return { date: this.nextWorkDay(today), halfDay: 0 };
  }

  static isPositionBefore(date, halfDay, otherDate, otherHalfDay) {
    const left = new Date(date).setHours(0, 0, 0, 0);
    const right = new Date(otherDate).setHours(0, 0, 0, 0);
    return left < right || (left === right && Number(halfDay) < Number(otherHalfDay));
  }

  static countCompletedHalfDays(startDate, startHalfDay, endDate, endHalfDay) {
    let current = new Date(startDate);
    let halfDay = this.normalizeHalfDay(startHalfDay);
    const target = new Date(endDate);
    const targetHalfDay = this.normalizeHalfDay(endHalfDay);
    let completed = 0;

    while (this.isPositionBefore(current, halfDay, target, targetHalfDay)) {
      completed += 1;
      if (halfDay === 0) {
        halfDay = 1;
      } else {
        current = this.nextWorkDay(current);
        halfDay = 0;
      }
    }
    return completed;
  }

  static parseInputDate(value) {

    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || "")) {
      throw new Error("Укажите дату переноса.");
    }

    const date = new Date(value + "T00:00:00");

    if (Number.isNaN(date.getTime())) {
      throw new Error("Указана некорректная дата.");
    }

    while (date.getDay() === 0 || date.getDay() === 6) {
      date.setDate(date.getDate() + 1);
    }

    return date;

  }

  static parseStoredDate(value) {

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    const match = String(value || "").match(/^(\d{2})\.(\d{2})(?:\.(\d{4}))?$/);

    if (!match) return null;

    const currentYear = new Date().getFullYear();
    const date = new Date(
      Number(match[3] || currentYear),
      Number(match[2]) - 1,
      Number(match[1])
    );

    return Number.isNaN(date.getTime()) ? null : date;

  }

}
