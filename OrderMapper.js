class OrderMapper {

  static map(row) {

    return new Order({

      id: row["ID"],

      manager: row["Менеджер"],

      contract: row["Договор"],

      client: row["Клиент"],

      product: row["Изделие"],

      displayName:
        (row["Договор"] || "") +
        " • " +
        (row["Клиент"] || "") +
        " • " +
        (row["Изделие"] || ""),

      stage: row["Этап"],

      responsible: row["Ответственный"],

      constructor: row["Конструктор"],

      queue: Number(row["Очередность"]) || 0,

      startDate: row["Начало"],

      duration: Number(row["Длительность"]) || 1,

      finishDate: row["Окончание"],

      comment: row["Комментарий"]

    });

  }

}