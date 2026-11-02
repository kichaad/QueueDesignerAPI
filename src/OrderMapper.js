/**
 * Преобразование строки таблицы в объект Order
 */
class OrderMapper {

  static map(row) {

    return new Order({

      // Новый составной ключ
      key: this.buildKey(row),

      manager: row["Менеджер"] || "",

      contract: row["Договор"] || "",

      client: row["Клиент"] || "",

      product: row["Изделие"] || "",

      displayName: [
        row["Договор"],
        row["Клиент"],
        row["Изделие"]
      ]
      .filter(Boolean)
      .join(" • "),

      stage: row["Этап"] || "",

      responsible: row["Ответственный"] || "",

      constructor: row["Конструктор"] || "",

      queue: Number(row["Очередность"] ?? row["Очередь"]) || 0,

      startDate: row["Начало"] || "",

      plannedDate: row["Плановая дата"] || "",

      deadline: row["Дедлайн"] || "",

      deadlineMode: row["Способ определения дедлайна"] || "От согласования проформы",

      manufacturingDays: Number(row["Срок изготовления, дней"]) || 50,

      manualDeadlineDate: row["Дата дедлайна вручную"] || "",

      manualPriority: this.toBoolean(row["Ручной приоритет даты"]),

      actualStartDate: row["Факт. начало работы"] || "",

      actualStartHalf: row["Факт. половина начала"] || "",

      remainingDuration: Number(row["Остаток длительности"]) || 0,

      duration: Number(row["Длительность"]) || 1,

      tzDays: Number(row["Дни ТЗ"]) || 0,

      proformaDays: Number(row["Дни Проформы"]) || 0,

      kdDays: Number(row["Дни КД"]) || 0,

      kdReserveQueue: Number(row["Резерв КД — очередь"]) || 0,

      proformaApprovedDate: row["Проформа Согласована"] || "",

      finishDate: row["Окончание"] || "",

      comment: row["Комментарий"] || "",

      installationCorpusComment: row["Комментарий монтажа корпуса"] || "",
      installationCorpusMdfComment: row["Комментарий монтажа корпуса МДФ"] || "",
      installationPaintedLdspComment: row["Комментарий монтажа крашеного ЛДСП"] || "",
      installationFacadeComment: row["Комментарий монтажа фасада"] || "",
      installationMdfComment: row["Комментарий монтажа фасада МДФ"] || "",
      installationVeneerComment: row["Комментарий монтажа фасада Шпон"] || "",
      installationLegacyComment: row["Комментарий монтажа legacy"] || "",
      claimComment: row["Комментарий рекламации"] || "",
      manualClaim: row["Ручная рекламация"] === true || String(row["Ручная рекламация"] || "").toLowerCase() === "true",

      status: row["Статус"] || "",

      cloudFolderUrl: row["Папка Google Drive"] || "",

      productionDays: {
        preparation: Number(row["Дни Подготовка"]) || 0,
        primer1: Number(row["Дни Грунт 1 слой"]) || Number(row["Дни Грунт"]) || 0,
        sanding1: Number(row["Дни Шлифовка"]) || 0,
        primer2: Number(row["Дни Грунт 2 слой"]) || 0,
        sanding2: Number(row["Дни Шлифовка 2 слой"]) || Number(row["Дни Белая шлифовка"]) || 0,
        backPainting: Number(row["Дни Покраска тыльной стороны"]) || 0,
        masking: Number(row["Дни Обклейка"]) || 0,
        frontPainting: Number(row["Дни Покраска лицевой стороны"]) || Number(row["Дни Покраска"]) || 0,
        preLacquerSanding: Number(row["Дни Шлифовка перед лаком"]) || 0,
        lacquer: Number(row["Дни Лак"]) || 0,
        polishing: Number(row["Дни Полировка"]) || 0
      },

      lacquerOperations: this.toBoolean(row["Лаковые операции"]),

      routeExclusions: {
        primer2: this.toBoolean(row["Исключить Грунт 2 слой"]),
        sanding2: this.toBoolean(row["Исключить Шлифовка 2 слой"]),
        backPainting: this.toBoolean(row["Исключить Покраска тыльной стороны"]),
        masking: this.toBoolean(row["Исключить Обклейка"])
      },

      paintingArea: Number(row["Площадь покраски, м2"]) || 0,

      backPaintingArea: Number(row["Площадь тыльной стороны, м2"]) || 0,

      frontPaintingArea: Number(row["Площадь лица и торцов, м2"]) || 0,

      polishingArea: Number(row["Площадь полировки, м2"]) || 0,

      ldspArea: Number(row["ЛДСП м2"]) || 0,
      mdfArea: Number(row["МДФ м2"]) || 0,
      lhdfArea: Number(row["ЛХДФ м2"]) || 0,
      veneerArea: Number(row["Шпон м2"]) || 0,

      facadeType: row["Тип фасада"] || "",

      paintingOperationCode: row["Операция Малярки"] || "",

      productionQueue: Number(row["Очередь производства"]) || 0,

      paintingManualPriority: this.toBoolean(row["Малярка — ручной приоритет"]),

      ldspStage: row["Этап ЛДСП"] || "",
      ldspStatus: row["Статус ЛДСП"] || "",
      // Новый флажок корпуса читается отдельно. Для исторических строк без
      // колонки обычный корпус остаётся включённым по умолчанию.
      hasCorpusLdsp: String(row["Корпус ЛДСП"] == null ? "" : row["Корпус ЛДСП"]).trim() === ""
        ? true
        : this.toBoolean(row["Корпус ЛДСП"]),
      hasCorpusMdf: this.toBoolean(row["Корпус МДФ"]),
      corpusMdfStage: row["Этап Корпус МДФ"] || "",
      corpusMdfStatus: row["Статус Корпус МДФ"] || "",
      corpusMdfPaintingStage: row["Этап Малярка Корпус МДФ"] || "",
      corpusMdfPaintingStatus: row["Статус Малярка Корпус МДФ"] || "",
      corpusMdfPaintingOperationCode: row["Операция Малярка Корпус МДФ"] || "",
      hasPaintedLdsp: this.toBoolean(row["Крашеный ЛДСП"]),
      paintedLdspStage: row["Этап Крашеный ЛДСП"] || "",
      paintedLdspStatus: row["Статус Крашеный ЛДСП"] || "",
      paintedLdspPaintingStage: row["Этап Малярка Крашеный ЛДСП"] || "",
      paintedLdspPaintingStatus: row["Статус Малярка Крашеный ЛДСП"] || "",
      paintedLdspPaintingOperationCode: row["Операция Малярка Крашеный ЛДСП"] || "",
      mdfStage: row["Этап МДФ"] || "",
      mdfStatus: row["Статус МДФ"] || "",
      hasMdfFacade: this.toBoolean(row["Фасад МДФ"]),
      hasLdspFacade: this.toBoolean(row["Фасад ЛДСП"]),
      veneerStage: row["Этап Шпон"] || "",
      veneerStatus: row["Статус Шпон"] || "",
      hasVeneerFacade: this.toBoolean(row["Фасад Шпон"]),
      hasNoFacade: this.toBoolean(row["Без фасада"]),
      facadePackagingStage: row["Этап упаковки фасада"] || "",
      facadePackagingStatus: row["Статус упаковки фасада"] || "",

      // Новые независимые фасадные ветки. Старые общие поля выше остаются
      // только для совместимости с уже начатыми маршрутами.
      mdfPaintingStage: row["Этап Малярка МДФ"] || "",
      mdfPaintingStatus: row["Статус Малярка МДФ"] || "",
      mdfPaintingOperationCode: row["Операция Малярка МДФ"] || "",
      veneerPaintingStage: row["Этап Малярка Шпон"] || "",
      veneerPaintingStatus: row["Статус Малярка Шпон"] || "",
      veneerPaintingOperationCode: row["Операция Малярка Шпон"] || "",
      mdfFacadePackagingStage: row["Этап упаковки фасада МДФ"] || "",
      mdfFacadePackagingStatus: row["Статус упаковки фасада МДФ"] || "",
      veneerFacadePackagingStage: row["Этап упаковки фасада Шпон"] || "",
      veneerFacadePackagingStatus: row["Статус упаковки фасада Шпон"] || "",

      sandingWorker: row["Шлифовщик"] || "",

      sandingQueue: Number(row["Очередь шлифовки"]) || 0,

      whiteSandingReserveQueue: Number(row["Резерв Белая шлифовка — очередь"]) || 0,

      sandingPlannedDate: row["Плановая дата шлифовки"] || "",

      painter: row["Маляр"] || "",

      primerQueue: Number(row["Очередь грунт"]) || 0,

      paintingReserveQueue: Number(row["Резерв Покраска — очередь"]) || 0,

      paintingPlannedDate: row["Плановая дата покраски"] || ""

    });

  }

  /**
   * Составной ключ изделия
   */
  static buildKey(row) {

    return [
      row["Договор"] || "",
      row["Клиент"] || "",
      row["Изделие"] || ""
    ].join("|");

  }

  static toBoolean(value) {
    return value === true || ["true", "да", "1"].includes(String(value || "").trim().toLowerCase());
  }

}
