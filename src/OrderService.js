/**
 * Сервис работы с очередью изделий
 */
class OrderService {

  // Явная пользовательская миграция должна быть выполнена до первого нового
  // snapshot: иначе старые значения не попадут в новую вкладку СГП. Маркер
  // делает операцию одноразовой, а сама миграция остаётся идемпотентной.
  static ensureSgpStageMigration() {
    const properties = PropertiesService.getScriptProperties();
    // V2 охватывает также отдельный лист «Этапы». V1 мог быть отмечен до
    // фактической замены исторических значений в рабочей таблице.
    const marker = "MIGRATION_SGP_STAGE_NAMES_V2";
    if (properties.getProperty(marker) === "done") return { success: true, skipped: true };
    const result = SheetService.migrateDeliveryStagesToSgp();
    properties.setProperty(marker, "done");
    return result;
  }

  /**
   * Возвращает очередь изделий
   */
  static getOrders(section, preloadedRows) {

    const activeSection = ["constructors", "malarka", "production", "production_master", "installation", "claims", "order_queue", "supply", "plate", "supplier_orders", "equipment", "equipment_invoices", "tabletops"].includes(section)
      ? section
      : "constructors";

    if (activeSection === "supply") return SupplyService.getSnapshot();
    if (activeSection === "plate") return PlateMaterialService.getSnapshot();
    if (activeSection === "supplier_orders") return SupplierOrderService.getSnapshot();
    if (activeSection === "equipment") return EquipmentService.getSnapshot();
    if (activeSection === "equipment_invoices") return EquipmentInvoiceService.getSnapshot();
    if (activeSection === "tabletops") return TabletopService.getSnapshot();

    if (!preloadedRows) this.ensureSgpStageMigration();

    let orderRows = preloadedRows || SheetService.getData(CONFIG.SHEETS.ORDERS);
    const orders = SheetService
      .objectsFromRows(orderRows)

      .map(row => OrderMapper.map(row))

      .sort((a, b) => a.queue - b.queue);

    let overallMismatches = 0;
    orders.forEach(order => {
      const overall = this.resolveOverallStage(order);
      if (!overall) return;
      if (order.stage !== overall.stage || order.status !== overall.status) overallMismatches += 1;
      order.stage = overall.stage;
      order.status = overall.status;
    });
    if (overallMismatches) console.log(JSON.stringify({ event: "OVERALL_STAGE_DIFFERENCES", count: overallMismatches }));

    let sectionOrders = [];

    if (activeSection === "order_queue") {
      const plannedReadinessDates = this.getPaintingPlannedReadinessDates(orders);
      sectionOrders = orders.map(order => Object.assign({}, order.toJSON(), {
        weekday: "",
        plannedReadinessDate: plannedReadinessDates[order.key] || ""
      }));
    } else if (activeSection === "constructors") {
      const queueOrders = this.getConstructorTasks(orders);
      const calendar = CalendarService.buildCalendar(queueOrders);
      const readyDates = queueOrders.reduce((result, task) => {
        if (task.taskType === "kd" || task.taskType === "kd-reserve") {
          result[task.key] = task.finishDate || result[task.key] || "";
        }
        return result;
      }, {});
      calendar.forEach(task => {
        task.readyDate = readyDates[task.key] || task.finishDate || "";
      });
      const designStages = new Set(["ТЗ", "Технолог", "Проформа", "Согласование", "Правки", "КД"]);
      const inactiveOrders = orders
        .filter(order => designStages.has(String(order.stage || "").trim()) && !this.isConstructorQueueOrder(order))
        .map(order => Object.assign({}, order.toJSON(), { weekday: "" }));
      sectionOrders = calendar.concat(inactiveOrders);
    } else if (activeSection === "malarka") {
      const paintingNorms = PaintingNormService.getClientConfig();
      const calendar = this.getPaintingRouteSchedule(orders);
      const bufferOrders = this.getPaintingBranchOrders(orders)
        .filter(order => String(order.stage || "").trim() === "В покраску")
        .map(order => {
          const preview = PaintingNormService.previewOrder(order);
          const totalNormHours = preview && preview.valid ? Math.round(preview.totalRawMinutes / 6) / 10 : null;
          return Object.assign({}, order.toJSON(), {
            weekday: "",
            normProfile: preview && preview.profile ? preview.profile.name : "",
            normHours: totalNormHours,
            totalNormHours,
            normError: preview && !preview.valid
              ? ((preview.operations.find(item => !item.valid) || {}).message || "Требуется настройка норм.")
              : ""
          });
        });
      sectionOrders = calendar.concat(bufferOrders);
    } else if (activeSection === "production_master") {
      sectionOrders = this.getMasterWorkTasks(orders);
    } else if (activeSection === "claims") {
      sectionOrders = orders
        .filter(order => String(order.stage || "").trim() === "Рекламация")
        .map(order => Object.assign({}, order.toJSON(), {
          weekday: "",
          claimOverview: this.getClaimOverview(order)
        }));
    } else {
      const manufacturingTasks = this.getManufacturingTasks(orders);
      const manufacturingTaskKeys = new Set(manufacturingTasks.map(task => String(task.orderKey || task.key || "").trim()));
      const productionStages = new Set(["Очередь", "Запуск", "Распил", "Кромка", "ЧПУ", "Присадка", "Фрезеровка", "Подготовка", "Сборка", "Упаковка", "СГП", "Монтаж", "Завершен"]);
      const baseOrders = orders
        .filter(order => productionStages.has(String(order.stage || "").trim()) &&
          !this.hasManufacturingTasks(order) && !manufacturingTaskKeys.has(String(order.key || "").trim()))
        .map(order => Object.assign({}, order.toJSON(), { weekday: "" }));
      if (activeSection === "installation") {
        // Вкладки Монтажей строятся только из явных потоков, а не из общего
        // этапа изделия. Это исключает строку «Запуск» рядом с фасадом.
        sectionOrders = this.getInstallationTasks(orders).concat(this.getInstallationOverviewOrders(orders));
      } else {
        sectionOrders = manufacturingTasks.concat(baseOrders);
      }
    }



    // Возврат данных клиенту
    const dictionaries = Object.assign({}, DictionaryService.getAll());
    dictionaries["Доступные шлифовщики"] = this.getAvailableSandingWorkers(dictionaries);

    return JSON.parse(JSON.stringify({

      section: activeSection,

      orders: sectionOrders,

      dictionaries: dictionaries,

      paintingNorms: activeSection === "malarka" ? PaintingNormService.getClientConfig() : null

    }));


  }

  /**
   * Формирует XLSX из уже отфильтрованных в интерфейсе строк общей очереди.
   * Временная Google-таблица удаляется в любом исходе операции.
   */
  static exportAllOrdersExcel(rows, filters) {
    const source = Array.isArray(rows) ? rows : [];
    const values = source.filter(item => item && typeof item === "object").map((item, index) => [
      Number(item.number) || index + 1,
      String(item.deadline || "—"),
      String(item.proformaApproved || "—"),
      String(item.plannedReadiness || "—"),
      String(item.contract || ""),
      String(item.client || ""),
      String(item.manager || "—"),
      String(item.product || ""),
      String(item.stage || "")
    ]);
    if (!values.length) throw new Error("По текущим фильтрам нет строк для выгрузки в Excel.");

    const columns = ["#", "Дедлайн", "Проформа согласована", "Плановая готовность", "Договор", "Клиент", "Менеджер", "Изделие", "Этап"];
    const xlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    const spreadsheet = SpreadsheetApp.create("Временный экспорт очереди заказов " + stamp);
    const spreadsheetId = spreadsheet.getId();
    const sheet = spreadsheet.getSheets()[0];

    try {
      sheet.setName("Очередь заказов");
      sheet.getRange(1, 1, 1, columns.length).merge()
        .setValue("Очередь заказов")
        .setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14)
        .setHorizontalAlignment("center");
      sheet.getRange(2, 1, 1, columns.length).merge()
        .setValue("Фильтры: " + String(filters || "Все строки"))
        .setBackground("#f8fafc").setFontColor("#475569").setWrap(true);
      sheet.getRange(4, 1, 1, columns.length).setValues([columns])
        .setBackground("#0f766e").setFontColor("#ffffff").setFontWeight("bold").setWrap(true);
      sheet.getRange(5, 1, values.length, columns.length).setValues(values)
        .setVerticalAlignment("top").setWrap(true)
        .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
      sheet.setFrozenRows(4);
      sheet.autoResizeColumns(1, columns.length);
      sheet.setColumnWidth(8, 260);
      sheet.getRange(5, 1, values.length, 1).setHorizontalAlignment("center");
      SpreadsheetApp.flush();

      const response = UrlFetchApp.fetch(
        "https://www.googleapis.com/drive/v3/files/" + spreadsheetId + "/export?mimeType=" + encodeURIComponent(xlsxMimeType),
        { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
      );
      if (response.getResponseCode() !== 200) {
        throw new Error("Не удалось сформировать Excel-файл: " + response.getContentText().slice(0, 300));
      }
      return {
        filename: "Очередь_заказов_" + stamp + ".xlsx",
        mimeType: xlsxMimeType,
        base64: Utilities.base64Encode(response.getBlob().getBytes())
      };
    } finally {
      try {
        Drive.Files.trash(spreadsheetId);
      } catch (cleanupError) {
        console.warn("Не удалось удалить временную таблицу экспорта очереди: " + cleanupError.message);
      }
    }
  }

  /**
   * Формирует XLSX из текущего списка производственных потоков.
   * Набор строк и порядок уже рассчитаны интерфейсом для активной вкладки.
   */
  static exportProductionExcel(rows, filters, layout) {
    const source = Array.isArray(rows) ? rows : [];
    const isMasterMatrix = layout === "masterMatrix";
    const isCompactProduction = layout === "productionFlow" || layout === "productionLaunch";
    const isLaunch = layout === "productionLaunch";
    const masterSupplyColumns = SupplyService.MASTER_SIGNAL_GROUPS;
    const masterSignalValue = (item, key) => String(((item.supplySignals || {})[key] || {}).label || "Нет позиций");
    const values = source.filter(item => item && typeof item === "object").map((item, index) => isMasterMatrix ? [
      Number(item.number) || index + 1,
      ...masterSupplyColumns.map(group => masterSignalValue(item, group.key)),
      String(item.deadline || "—"),
      String(item.orderMeta || ""),
      String(item.product || ""),
      String(item.corpus || "—"),
      String(item.corpusMdf || "—"),
      String(item.paintedLdsp || "—"),
      String(item.mdf || "—"),
      String(item.veneer || "—")
    ] : isCompactProduction ? [
      Number(item.number) || index + 1,
      String(item.deadline || "—"),
      String(item.orderMeta || ""),
      String(item.product || ""),
      String(item.productType || "—"),
      String(item.area || "—"),
      String(item.stage || ""),
      String(isLaunch ? item.constructor || "—" : item.status || "—")
    ] : [
      Number(item.number) || index + 1,
      String(item.deadline || "—"),
      String(item.contract || ""),
      String(item.client || ""),
      String(item.product || ""),
      String(item.productType || "—"),
      String(item.area || "—"),
      String(item.stage || ""),
      String(item.status || "—")
    ]);
    if (!values.length) throw new Error("В текущем производственном списке нет строк для выгрузки в Excel.");

    const columns = isMasterMatrix
      ? ["#"].concat(masterSupplyColumns.map(group => group.title), ["Дедлайн", "Договор / клиент", "Изделие", "Корпус", "Корпус МДФ", "Крашеный ЛДСП", "Фасад МДФ", "Фасад Шпон"])
      : isCompactProduction ? ["#", "Дедлайн", "Договор / клиент", "Изделие", "Тип изделия", "Площадь, м²", "Этап", isLaunch ? "Конструктор" : "Статус"]
      : ["#", "Дедлайн", "Договор", "Клиент", "Изделие", "Тип изделия", "Площадь, м²", "Этап", "Статус"];
    const xlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    const spreadsheet = SpreadsheetApp.create("Временный экспорт производства " + stamp);
    const spreadsheetId = spreadsheet.getId();
    const sheet = spreadsheet.getSheets()[0];

    try {
      sheet.setName(isMasterMatrix ? "АРМ мастера" : "Производство");
      sheet.getRange(1, 1, 1, columns.length).merge()
        .setValue(isMasterMatrix ? "АРМ мастера" : "Производство")
        .setBackground("#7c2d12").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14)
        .setHorizontalAlignment("center");
      sheet.getRange(2, 1, 1, columns.length).merge()
        .setValue("Фильтры: " + String(filters || "Текущий список"))
        .setBackground("#fff7ed").setFontColor("#9a3412").setWrap(true);
      sheet.getRange(4, 1, 1, columns.length).setValues([columns])
        .setBackground("#c2410c").setFontColor("#ffffff").setFontWeight("bold").setWrap(true);
      sheet.getRange(5, 1, values.length, columns.length).setValues(values)
        .setVerticalAlignment("top").setWrap(true)
        .setBorder(true, true, true, true, true, true, "#fed7aa", SpreadsheetApp.BorderStyle.SOLID);
      sheet.setFrozenRows(4);
      sheet.autoResizeColumns(1, columns.length);
      if (isMasterMatrix) {
        masterSupplyColumns.forEach((group, index) => {
          sheet.setColumnWidth(index + 2, 88);
          const colorByState = { none: "#ffffff", red: "#ff3030", lilac: "#d43dff", yellow: "#ffe600", green: "#00d45a" };
          const backgrounds = source.map(item => [colorByState[String(((item.supplySignals || {})[group.key] || {}).state || "none")] || colorByState.none]);
          sheet.getRange(5, index + 2, backgrounds.length, 1).setBackgrounds(backgrounds)
            .setFontColor("#000000").setFontWeight("bold")
            .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        });
        sheet.setColumnWidth(11, 280);
        [12, 13, 14, 15, 16].forEach(column => sheet.setColumnWidth(column, 180));
      } else {
        sheet.setColumnWidth(isCompactProduction ? 4 : 5, 280);
      }
      sheet.getRange(5, 1, values.length, 1).setHorizontalAlignment("center");
      SpreadsheetApp.flush();

      const response = UrlFetchApp.fetch(
        "https://www.googleapis.com/drive/v3/files/" + spreadsheetId + "/export?mimeType=" + encodeURIComponent(xlsxMimeType),
        { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
      );
      if (response.getResponseCode() !== 200) {
        throw new Error("Не удалось сформировать Excel-файл: " + response.getContentText().slice(0, 300));
      }
      return {
        filename: (isMasterMatrix ? "АРМ_мастера_" : "Производство_") + stamp + ".xlsx",
        mimeType: xlsxMimeType,
        base64: Utilities.base64Encode(response.getBlob().getBytes())
      };
    } finally {
      try {
        Drive.Files.trash(spreadsheetId);
      } catch (cleanupError) {
        console.warn("Не удалось удалить временную таблицу экспорта производства: " + cleanupError.message);
      }
    }
  }

  // Универсальный экспорт экранных списков. Клиент передаёт уже видимые строки,
  // а схема фиксирована по типу списка, чтобы в Excel не попадали служебные поля.
  static exportListExcel(exportType, rows, filters) {
    const schemas = {
      constructor: { title: "Конструкторы", columns: ["#", "Дедлайн", "Договор", "Клиент", "Изделие", "Этап", "Статус", "Конструктор"], keys: ["number", "deadline", "contract", "client", "product", "stage", "status", "responsible"] },
      painting: { title: "Малярка", columns: ["#", "Дедлайн", "Договор", "Клиент", "Изделие", "Тип изделия", "Этап", "Статус", "Исполнитель"], keys: ["number", "deadline", "contract", "client", "product", "productType", "stage", "status", "responsible"] },
      installation: { title: "Монтажи", columns: ["#", "Дедлайн", "Договор", "Клиент", "Изделие", "Тип изделия", "Этап", "Статус"], keys: ["number", "deadline", "contract", "client", "product", "productType", "stage", "status"] },
      installationAll: { title: "Все монтажи", columns: ["#", "Дедлайн", "Договор", "Клиент", "Изделие", "Корпус", "Корпус МДФ", "Крашеный ЛДСП", "Фасад МДФ", "Фасад Шпон"], keys: ["number", "deadline", "contract", "client", "product", "corpus", "corpusMdf", "paintedLdsp", "mdf", "veneer"] },
      claims: { title: "Рекламации", columns: ["#", "Дедлайн", "Договор", "Клиент", "Изделие", "Корпус", "Корпус МДФ", "Крашеный ЛДСП", "Фасад МДФ", "Фасад Шпон", "Комментарий"], keys: ["number", "deadline", "contract", "client", "product", "corpus", "corpusMdf", "paintedLdsp", "mdf", "veneer", "comment"] },
      supply: { title: "Снабжение", columns: ["#", "Договор", "Клиент", "Изделие", "Дедлайн", "Артикул", "Наименование", "Количество", "Ед.", "К закупке", "Поставщик", "Цена", "Статус"], keys: ["number", "contract", "client", "product", "deadline", "article", "name", "quantity", "unit", "purchase", "supplier", "price", "status"] },
      supplyRequests: { title: "Снабжение", columns: ["#", "Договор", "Клиент", "Изделие", "Дедлайн", "Артикул", "Наименование", "Комментарий", "Количество", "Ед.", "Статус"], keys: ["number", "contract", "client", "product", "deadline", "article", "name", "comment", "quantity", "unit", "status"] },
      plate: { title: "Плитный материал", columns: ["#", "Договор", "Клиент", "Изделие", "Дедлайн", "Категория", "Артикул", "Наименование", "Декор", "Толщина, мм", "Размер листа", "К заказу, листов", "Статус"], keys: ["number", "contract", "client", "product", "deadline", "category", "article", "name", "decor", "thickness", "sheetSize", "purchase", "status"] },
      supplierOrders: { title: "Заказы поставщикам", columns: ["#", "Номер", "Создан", "Поставщик", "Статус", "Номера счетов", "Дата поставки", "Изделия", "Позиций", "Сумма"], keys: ["number", "orderNumber", "createdAt", "supplier", "status", "invoices", "deliveryDate", "products", "positions", "total"] },
      equipment: { title: "Техника", columns: ["Договор", "Клиент", "Наименование", "Модель", "Менеджер", "Поставщик", "Статус"], keys: ["contract", "client", "name", "model", "manager", "supplier", "status"] },
      equipmentInvoices: { title: "Счета техники", columns: ["Номер", "Дата", "Поставщик", "Сумма", "Статус", "Позиции"], keys: ["number", "date", "supplier", "amount", "status", "lines"] },
      tabletops: { title: "Столешницы", columns: ["Договор", "Клиент", "Наименование", "Менеджер", "Поставщик", "Статус"], keys: ["contract", "client", "name", "manager", "supplier", "status"] }
    };
    const schema = schemas[String(exportType || "").trim()];
    if (!schema) throw new Error("Неизвестный список для выгрузки.");
    const source = Array.isArray(rows) ? rows.filter(item => item && typeof item === "object") : [];
    if (!source.length) throw new Error("По текущим фильтрам нет строк для выгрузки в Excel.");
    const values = source.map((item, index) => schema.keys.map(key =>
      String(key === "number" && item[key] == null ? index + 1 : (item[key] == null ? "—" : item[key]))
    ));
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    const xlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    const spreadsheet = SpreadsheetApp.create("Временный экспорт " + schema.title + " " + stamp);
    const spreadsheetId = spreadsheet.getId();
    const sheet = spreadsheet.getSheets()[0];
    try {
      sheet.setName(schema.title.slice(0, 99));
      sheet.getRange(1, 1, 1, schema.columns.length).merge().setValue(schema.title)
        .setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
      sheet.getRange(2, 1, 1, schema.columns.length).merge().setValue("Фильтры: " + String(filters || "Текущий список"))
        .setBackground("#f8fafc").setFontColor("#475569").setWrap(true);
      sheet.getRange(4, 1, 1, schema.columns.length).setValues([schema.columns])
        .setBackground("#0f766e").setFontColor("#ffffff").setFontWeight("bold").setWrap(true);
      sheet.getRange(5, 1, values.length, schema.columns.length).setValues(values)
        .setVerticalAlignment("top").setWrap(true)
        .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
      sheet.setFrozenRows(4);
      sheet.autoResizeColumns(1, schema.columns.length);
      SpreadsheetApp.flush();
      const response = UrlFetchApp.fetch(
        "https://www.googleapis.com/drive/v3/files/" + spreadsheetId + "/export?mimeType=" + encodeURIComponent(xlsxMimeType),
        { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
      );
      if (response.getResponseCode() !== 200) throw new Error("Не удалось сформировать Excel-файл: " + response.getContentText().slice(0, 300));
      return { filename: schema.title.replace(/[^а-яёa-z0-9]+/gi, "_") + "_" + stamp + ".xlsx", mimeType: xlsxMimeType, base64: Utilities.base64Encode(response.getBlob().getBytes()) };
    } finally {
      try { Drive.Files.trash(spreadsheetId); } catch (error) { console.warn("Не удалось удалить временную таблицу экспорта: " + error.message); }
    }
  }

  static getConstructorTasks(orders) {
    const tasks = [];

    orders.forEach(order => {
      if (order.stage === "Проформа" && String(order.status).toLowerCase() !== "готово") {
        const proformaDuration = order.remainingDuration || order.proformaDays || order.duration;
        tasks.push(Object.assign({}, order.toJSON(), {
          taskType: "proforma",
          duration: proformaDuration,
          queue: order.queue,
          canStart: true
        }));
        tasks.push(Object.assign({}, order.toJSON(), {
          taskType: "kd-reserve",
          stage: "КД (резерв)",
          status: "Резерв",
          duration: order.kdDays || 0.5,
          queue: order.kdReserveQueue || (Number(order.queue) + 1),
          manualPriority: false,
          remainingDuration: 0,
          canStart: false
        }));
      } else if (order.stage === "Согласование") {
        // После отправки на согласование Проформа видна одновременно
        // менеджеру и конструктору. Завершённые ранее Проформы со статусом
        // «Готово» в очередь конструктора не возвращаются.
        if (String(order.status).toLowerCase() !== "готово") {
          const proformaDuration = order.remainingDuration || order.proformaDays || order.duration;
          tasks.push(Object.assign({}, order.toJSON(), {
            taskType: "proforma",
            duration: proformaDuration,
            queue: order.queue,
            displayOnly: true,
            canStart: true
          }));
        }
        // Резерв мощности КД сохраняется в календаре до решения менеджера.
        tasks.push(Object.assign({}, order.toJSON(), {
          taskType: "kd-reserve",
          stage: "КД (резерв)",
          status: "Резерв",
          duration: order.kdDays || 0.5,
          queue: order.kdReserveQueue || (Number(order.queue) + 1),
          manualPriority: false,
          remainingDuration: 0,
          canStart: false
        }));
      } else if (order.stage === "КД" && String(order.status).toLowerCase() !== "готово") {
        const kdDuration = order.remainingDuration || order.kdDays || order.duration;
        tasks.push(Object.assign({}, order.toJSON(), {
          taskType: "kd",
          duration: kdDuration,
          queue: order.queue,
          canStart: true
        }));
      }
    });

    return tasks;
  }

  static hasManufacturingTasks(order) {
    return Boolean(String(order.ldspStage || "").trim() || String(order.mdfStage || "").trim() ||
      String(order.veneerStage || "").trim() || String(order.corpusMdfStage || "").trim() ||
      String(order.paintedLdspStage || "").trim() || String(order.facadePackagingStage || "").trim() ||
      String(order.mdfFacadePackagingStage || "").trim() || String(order.veneerFacadePackagingStage || "").trim() ||
      String(order.corpusMdfPaintingStage || "").trim() || String(order.paintedLdspPaintingStage || "").trim() ||
      String(order.mdfPaintingStage || "").trim() === "Подготовка" ||
      String(order.veneerPaintingStage || "").trim() === "Подготовка");
  }

  // Дополнительные корпусные маршруты независимы от обычного ЛДСП-корпуса.
  // Их поля не создаются и не заполняются для старых изделий без флажка.
  static getAdditionalCorpusBranchConfig(branch) {
    return {
      "corpus-mdf": {
        selectedProperty: "hasCorpusMdf", selectedField: "Корпус МДФ",
        stageProperty: "corpusMdfStage", statusProperty: "corpusMdfStatus",
        stageField: "Этап Корпус МДФ", statusField: "Статус Корпус МДФ",
        paintingStageProperty: "corpusMdfPaintingStage", paintingStatusProperty: "corpusMdfPaintingStatus",
        paintingOperationProperty: "corpusMdfPaintingOperationCode",
        paintingStageField: "Этап Малярка Корпус МДФ", paintingStatusField: "Статус Малярка Корпус МДФ",
        paintingOperationField: "Операция Малярка Корпус МДФ",
        initialStage: "Распил МДФ", productFlow: "Корпус МДФ", commentField: "Комментарий монтажа корпуса МДФ",
        commentProperty: "installationCorpusMdfComment", needsPreparation: true
      },
      "painted-ldsp": {
        selectedProperty: "hasPaintedLdsp", selectedField: "Крашеный ЛДСП",
        stageProperty: "paintedLdspStage", statusProperty: "paintedLdspStatus",
        stageField: "Этап Крашеный ЛДСП", statusField: "Статус Крашеный ЛДСП",
        paintingStageProperty: "paintedLdspPaintingStage", paintingStatusProperty: "paintedLdspPaintingStatus",
        paintingOperationProperty: "paintedLdspPaintingOperationCode",
        paintingStageField: "Этап Малярка Крашеный ЛДСП", paintingStatusField: "Статус Малярка Крашеный ЛДСП",
        paintingOperationField: "Операция Малярка Крашеный ЛДСП",
        initialStage: "Распил ЛДСП", productFlow: "Крашеный ЛДСП", commentField: "Комментарий монтажа крашеного ЛДСП",
        commentProperty: "installationPaintedLdspComment", needsPreparation: false
      }
    }[String(branch || "").trim().toLowerCase()] || null;
  }

  static overallStageForOperation(value) {
    const stage = String(value || "").trim();
    if (["Распил", "Распил ЛДСП", "Распил МДФ", "Распил Шпон"].includes(stage)) return "Распил";
    if (["Кромка", "Кромка Шпон"].includes(stage)) return "Кромка";
    if (["ЧПУ", "Присадка", "Присадка Шпон", "Фрезеровка"].includes(stage)) return "ЧПУ";
    if (stage === "Подготовка") return "Подготовка";
    if (stage === "Сборка") return "Сборка";
    if (["В покраску", "Грунт 1 слой", "Шлифовка 1 слой", "Грунт 2 слой",
      "Шлифовка 2 слой", "Покраска тыльной стороны", "Обклейка", "Покраска лицевой стороны",
      "Шлифовка перед лаком", "Лак", "Полировка"].includes(stage)) return "В покраску";
    if (["Упаковка", "Упаковка Корпуса", "Упаковка Фасада"].includes(stage)) return "Упаковка";
    if (["СГП", "СГП Корпуса", "СГП Фасада"].includes(stage)) return "СГП";
    if (["Монтаж", "Рекламация", "Завершен"].includes(stage)) return stage;
    return "";
  }

  // Возвращает null для старой или неполной записи: отсутствие ветки нельзя
  // подменять догадкой о её состоянии. Чтение snapshot ничего не записывает.
  static resolveOverallStage(order) {
    // Рекламация — отдельная послемонтажная работа изделия. Ветки уже могут
    // быть завершены, поэтому общий этап не выводится из их раннего состояния.
    if (String(order.stage || "").trim() === "Рекламация") {
      return { stage: "Рекламация", status: "" };
    }
    const stages = [];
    const addSelectedBranch = (selected, manufacturingStage, paintingStage) => {
      if (!selected) return true;
      const stage = String(paintingStage || manufacturingStage || "").trim();
      if (!stage) return false;
      stages.push(stage);
      return true;
    };
    const body = String(order.ldspStage || "").trim();
    if (body) stages.push(body);
    if (!addSelectedBranch(order.hasCorpusMdf, order.corpusMdfStage, order.corpusMdfPaintingStage)) return null;
    if (!addSelectedBranch(order.hasPaintedLdsp, order.paintedLdspStage, order.paintedLdspPaintingStage)) return null;
    const mdf = order.mdfFacadePackagingStage || order.mdfPaintingStage || order.mdfStage;
    const veneer = order.veneerFacadePackagingStage || order.veneerPaintingStage || order.veneerStage;
    const independentFacade = Boolean(String(order.mdfPaintingStage || "").trim() ||
      String(order.veneerPaintingStage || "").trim() || String(order.mdfFacadePackagingStage || "").trim() ||
      String(order.veneerFacadePackagingStage || "").trim());
    const legacyFacade = !independentFacade ? String(order.facadePackagingStage || "").trim() : "";
    if (order.hasMdfFacade) {
      const stage = mdf || (!order.hasVeneerFacade ? legacyFacade : "");
      if (!stage) return null;
      stages.push(stage);
    }
    if (order.hasVeneerFacade) {
      const stage = veneer || (!order.hasMdfFacade ? legacyFacade : "");
      if (!stage) return null;
      stages.push(stage);
    }
    if (!order.hasMdfFacade && !order.hasVeneerFacade && !order.hasLdspFacade && !order.hasNoFacade && legacyFacade) {
      stages.push(legacyFacade);
    }
    if (!stages.length) return null;
    const sequence = ["Распил", "Кромка", "ЧПУ", "Подготовка", "Сборка", "В покраску", "Упаковка", "СГП", "Монтаж", "Рекламация", "Завершен"];
    const phases = stages.map(stage => this.overallStageForOperation(stage));
    if (phases.some(stage => !stage)) return null;
    const stage = phases.reduce((earliest, current) =>
      sequence.indexOf(current) < sequence.indexOf(earliest) ? current : earliest);
    return { stage, status: stage === "Завершен" ? "Готово" : "" };
  }

  static projectRouteOrder(order, fields) {
    const names = {
      "Этап": "stage", "Статус": "status", "Корпус ЛДСП": "hasCorpusLdsp", "Этап ЛДСП": "ldspStage", "Статус ЛДСП": "ldspStatus",
      "Корпус МДФ": "hasCorpusMdf", "Этап Корпус МДФ": "corpusMdfStage", "Статус Корпус МДФ": "corpusMdfStatus",
      "Этап Малярка Корпус МДФ": "corpusMdfPaintingStage", "Статус Малярка Корпус МДФ": "corpusMdfPaintingStatus",
      "Крашеный ЛДСП": "hasPaintedLdsp", "Этап Крашеный ЛДСП": "paintedLdspStage", "Статус Крашеный ЛДСП": "paintedLdspStatus",
      "Этап Малярка Крашеный ЛДСП": "paintedLdspPaintingStage", "Статус Малярка Крашеный ЛДСП": "paintedLdspPaintingStatus",
      "Этап МДФ": "mdfStage", "Статус МДФ": "mdfStatus", "Фасад МДФ": "hasMdfFacade",
      "Этап Шпон": "veneerStage", "Статус Шпон": "veneerStatus", "Фасад Шпон": "hasVeneerFacade",
      "Фасад ЛДСП": "hasLdspFacade", "Без фасада": "hasNoFacade",
      "Этап упаковки фасада": "facadePackagingStage", "Статус упаковки фасада": "facadePackagingStatus",
      "Этап Малярка МДФ": "mdfPaintingStage", "Статус Малярка МДФ": "mdfPaintingStatus",
      "Этап Малярка Шпон": "veneerPaintingStage", "Статус Малярка Шпон": "veneerPaintingStatus",
      "Этап упаковки фасада МДФ": "mdfFacadePackagingStage",
      "Статус упаковки фасада МДФ": "mdfFacadePackagingStatus",
      "Этап упаковки фасада Шпон": "veneerFacadePackagingStage",
      "Статус упаковки фасада Шпон": "veneerFacadePackagingStatus"
    };
    const projected = Object.assign({}, order);
    Object.entries(fields || {}).forEach(([name, value]) => {
      if (names[name]) projected[names[name]] = value;
    });
    return projected;
  }

  static updateRouteFields(key, order, fields) {
    const result = Object.assign({}, fields);
    const overall = this.resolveOverallStage(this.projectRouteOrder(order, result));
    if (overall) {
      result["Этап"] = overall.stage;
      result["Статус"] = overall.status;
      result["Ответственный"] = "Производство";
    }
    SheetService.updateFields(key, result);
    return overall;
  }

  static getOverallManufacturingStages() {
    return ["Запуск", "Распил", "Кромка", "ЧПУ", "Подготовка", "Сборка", "В покраску", "Упаковка", "СГП", "Монтаж", "Рекламация", "Завершен"];
  }

  static buildOverallStageMoveFields(order, targetStage) {
    const stage = String(targetStage || "").trim();
    if (!this.getOverallManufacturingStages().includes(stage)) {
      throw new Error("Выберите производственный этап из списка.");
    }

    const hasFacadeSelection = order.hasLdspFacade || order.hasMdfFacade || order.hasVeneerFacade || order.hasNoFacade;
    // Рекламацию разрешаем выбрать и у исторической строки без заполненной
    // классификации фасада: это ручной послемонтажный перевод всего изделия.
    if (stage !== "Запуск" && stage !== "Рекламация" && !hasFacadeSelection) {
      throw new Error("Укажите материал фасада или выберите «Без фасада».");
    }
    if (stage === "В покраску" && !order.hasMdfFacade && !order.hasVeneerFacade && !order.hasCorpusMdf && !order.hasPaintedLdsp) {
      throw new Error("В этап «В покраску» можно перевести только изделие с окрашиваемой веткой.");
    }

    // Ручной перевод исторического изделия в рекламацию не должен менять
    // состояния его независимых веток: они могут ещё находиться в маршруте.
    if (stage === "Рекламация") {
      return {
        "Этап": "Рекламация",
        "Статус": "",
        "Ответственный": "Производство"
      };
    }

    const fields = {
      "Этап": stage,
      "Статус": stage === "Завершен" ? "Готово" : "",
      "Ответственный": "Производство",
      "Этап ЛДСП": "", "Статус ЛДСП": "",
      "Этап Корпус МДФ": "", "Статус Корпус МДФ": "",
      "Этап Малярка Корпус МДФ": "", "Статус Малярка Корпус МДФ": "", "Операция Малярка Корпус МДФ": "",
      "Этап Крашеный ЛДСП": "", "Статус Крашеный ЛДСП": "",
      "Этап Малярка Крашеный ЛДСП": "", "Статус Малярка Крашеный ЛДСП": "", "Операция Малярка Крашеный ЛДСП": "",
      "Этап МДФ": "", "Статус МДФ": "",
      "Этап Шпон": "", "Статус Шпон": "",
      "Этап Малярка МДФ": "", "Статус Малярка МДФ": "", "Операция Малярка МДФ": "",
      "Этап Малярка Шпон": "", "Статус Малярка Шпон": "", "Операция Малярка Шпон": "",
      "Этап упаковки фасада МДФ": "", "Статус упаковки фасада МДФ": "",
      "Этап упаковки фасада Шпон": "", "Статус упаковки фасада Шпон": "",
      "Этап упаковки фасада": "", "Статус упаковки фасада": "",
      "Операция Малярки": ""
    };
    if (stage === "Запуск") return fields;
    const route = {
      "Распил": { body: "Распил ЛДСП", mdf: "Распил МДФ", veneer: "Распил Шпон" },
      "Кромка": { body: "Кромка", mdf: "Фрезеровка", veneer: "Кромка Шпон" },
      "ЧПУ": { body: "Присадка", mdf: "Фрезеровка", veneer: "Присадка Шпон" },
      "Подготовка": { body: "Сборка", mdfPainting: "Подготовка", veneerPainting: "Подготовка" },
      "Сборка": { body: "Сборка", mdfPainting: "В покраску", veneerPainting: "В покраску" },
      "В покраску": { body: "Упаковка Корпуса", mdfPainting: "В покраску", veneerPainting: "В покраску" },
      "Упаковка": { body: "Упаковка Корпуса", mdfPackaging: "Упаковка Фасада", veneerPackaging: "Упаковка Фасада" },
      "СГП": { body: "СГП Корпуса", mdfPackaging: "СГП Фасада", veneerPackaging: "СГП Фасада" },
      "Монтаж": { body: "Монтаж", mdfPackaging: "Монтаж", veneerPackaging: "Монтаж" },
      "Завершен": { body: "Завершен", mdfPackaging: "Завершен", veneerPackaging: "Завершен" }
    }[stage];

    if (order.hasCorpusLdsp !== false) {
      fields["Этап ЛДСП"] = route.body;
      fields["Статус ЛДСП"] = stage === "Завершен" ? "Готово" : "";
    }
    if (order.hasCorpusMdf) {
      if (["Распил", "Кромка", "ЧПУ"].includes(stage)) fields["Этап Корпус МДФ"] = stage === "Распил" ? "Распил МДФ" : "Фрезеровка";
      else if (stage === "Подготовка") fields["Этап Малярка Корпус МДФ"] = "Подготовка";
      else if (stage === "В покраску") fields["Этап Малярка Корпус МДФ"] = "В покраску";
      else fields["Этап Корпус МДФ"] = stage === "Упаковка" ? "Упаковка Корпуса" : stage === "СГП" ? "СГП Корпуса" : stage;
      if (stage === "Завершен") fields["Статус Корпус МДФ"] = "Готово";
    }
    if (order.hasPaintedLdsp) {
      if (stage === "Распил") fields["Этап Крашеный ЛДСП"] = "Распил ЛДСП";
      else if (stage === "Кромка") fields["Этап Крашеный ЛДСП"] = "Кромка";
      else if (stage === "ЧПУ") fields["Этап Крашеный ЛДСП"] = "Присадка";
      else if (stage === "В покраску") fields["Этап Малярка Крашеный ЛДСП"] = "В покраску";
      else fields["Этап Крашеный ЛДСП"] = stage === "Подготовка" ? "Сборка" : stage === "Упаковка" ? "Упаковка Корпуса" : stage === "СГП" ? "СГП Корпуса" : stage;
      if (stage === "Завершен") fields["Статус Крашеный ЛДСП"] = "Готово";
    }
    if (order.hasMdfFacade) {
      if (route.mdf) fields["Этап МДФ"] = route.mdf;
      if (route.mdfPainting) fields["Этап Малярка МДФ"] = route.mdfPainting;
      if (route.mdfPackaging) fields["Этап упаковки фасада МДФ"] = route.mdfPackaging;
      if (stage === "Завершен") fields["Статус упаковки фасада МДФ"] = "Готово";
    }
    if (order.hasVeneerFacade) {
      if (route.veneer) fields["Этап Шпон"] = route.veneer;
      if (route.veneerPainting) fields["Этап Малярка Шпон"] = route.veneerPainting;
      if (route.veneerPackaging) fields["Этап упаковки фасада Шпон"] = route.veneerPackaging;
      if (stage === "Завершен") fields["Статус упаковки фасада Шпон"] = "Готово";
    }

    const projected = this.projectRouteOrder(order, fields);
    const overall = this.resolveOverallStage(projected);
    if (!overall || overall.stage !== stage) {
      throw new Error("Не удалось согласовать производственные ветки с выбранным общим этапом.");
    }
    return fields;
  }

  // Адресное восстановление подтверждённых исторических строк. dryRun не
  // записывает таблицу; повторный запуск пропускает уже исправленные строки.
  static migrateOverallStages(dryRun) {
    const specs = [];
    const add = (row, key, oldStage, expected, fields, newStage) =>
      specs.push({ row, key, oldStage, expected, fields, newStage });
    add(14, "36RM|Шагал|С/У гост", "СГП",
      { ldspStage: "СГП Корпуса", mdfPaintingStage: "В покраску" }, {}, "В покраску");
    add(23, "35RM|Пядухов|С/У", "Монтаж",
      { ldspStage: "Монтаж", mdfPaintingStage: "В покраску" }, {}, "В покраску");
    add(34, "38RM|Малыш|Кухня", "СГП",
      { ldspStage: "СГП Корпуса", mdfPaintingStage: "Подготовка", veneerStage: "Распил Шпон" }, {}, "Распил");
    add(35, "39RM|Малыш|Прихожая", "СГП",
      { ldspStage: "СГП Корпуса", mdfPaintingStage: "Подготовка" }, {}, "В покраску");
    add(36, "40RM|Малыш|Шкаф детская", "СГП",
      { ldspStage: "СГП Корпуса", mdfPaintingStage: "В покраску" }, {}, "В покраску");
    add(37, "41RM|Малыш|Шкаф спальня", "СГП",
      { ldspStage: "СГП Корпуса", mdfPaintingStage: "Подготовка" }, {}, "В покраску");
    add(40, "44RM|Лебедева Светлана|кухня", "В покраску",
      { ldspStage: "Сборка", mdfPaintingStage: "Подготовка" }, {}, "Сборка");
    add(41, "46RM|Коробков Александр|кухня", "СГП",
      { ldspStage: "СГП Корпуса", mdfStage: "В покраску" }, {}, "В покраску");
    add(48, "34RM26|Татанов|ГАРДЕРОБНАЯ", "СГП",
      { ldspStage: "СГП Корпуса", mdfPaintingStage: "В покраску" }, {}, "В покраску");
    add(51, "206RM25|Тимофеева|Кабинет", "СГП",
      { ldspStage: "СГП Корпуса", mdfPaintingStage: "В покраску", veneerStage: "Распил Шпон" }, {}, "Распил");
    add(54, "52RM26|Цыганова Юлия Сергеевна|Кабинет", "Распил",
      { ldspStage: "Присадка" }, {}, "ЧПУ");
    [
      [4, "30RM|Дубцова|Кухня", "Монтаж", "Монтаж"],
      [9, "225RM25|Новосельцева|Спальня", "СГП", "СГП Фасада"],
      [10, "206RM|Тимофеева|Рабочая зона", "Монтаж", "Монтаж"],
      [22, "35RM|Пядухов|Экран для ванной", "Монтаж", "Монтаж"],
      [24, "35RM|Пядухов|С/У мастер", "Упаковка", "Упаковка Фасада"],
      [25, "35RM|Пядухов|Полки на пилоне", "Упаковка", "Упаковка Фасада"],
      [26, "35RM|Пядухов|Полки в спальню", "Упаковка", "Упаковка Фасада"],
      [43, "25RM26|Пядухов|Красная тумба", "Монтаж", "Монтаж"]
    ].forEach(([row, key, oldStage, facade]) =>
      add(row, key, oldStage, { ldspStage: "", facadePackagingStage: facade },
        { "Этап ЛДСП": "Сборка", "Статус ЛДСП": "" }, "Сборка"));
    add(7, "27RM|Карташова|Кухня", "Монтаж",
      { ldspStage: "Завершен", facadePackagingStage: "Монтаж", mdfFacadePackagingStage: "Монтаж" },
      { "Этап упаковки фасада": "", "Статус упаковки фасада": "" }, "Монтаж");
    add(59, "21RM26|Каспарьян|Кухня", "Монтаж",
      { ldspStage: "Монтаж", facadePackagingStage: "Монтаж", mdfFacadePackagingStage: "Монтаж", veneerStage: "Монтаж" },
      { "Этап упаковки фасада": "", "Статус упаковки фасада": "" }, "Монтаж");
    add(8, "225RM25|Новосельцева|ТВ зона", "В покраску",
      { ldspStage: "Сборка", hasMdfFacade: false, mdfPaintingStage: "Подготовка" },
      { "Этап Малярка МДФ": "", "Статус Малярка МДФ": "", "Операция Малярка МДФ": "",
        "Этап МДФ": "", "Статус МДФ": "" }, "Сборка");
    add(27, "35RM|Пядухов|Входная группа", "Монтаж",
      { ldspStage: "В покраску", facadePackagingStage: "Монтаж" },
      { "Этап ЛДСП": "Монтаж", "Статус ЛДСП": "" }, "Монтаж");
    add(31, "45RM|Ковалева|Прихожая", "Распил",
      { ldspStage: "Завершен", hasMdfFacade: true, facadePackagingStage: "Упаковка Фасада", mdfPaintingStage: "" },
      { "Этап упаковки фасада": "", "Статус упаковки фасада": "",
        "Этап МДФ": "", "Статус МДФ": "Готово",
        "Этап Малярка МДФ": "В покраску", "Статус Малярка МДФ": "", "Операция Малярка МДФ": "" }, "В покраску");
    add(32, "45RM|Ковалева|Гардероб", "Распил",
      { ldspStage: "Завершен", hasLdspFacade: true, hasMdfFacade: false, hasVeneerFacade: false,
        facadePackagingStage: "Упаковка Фасада" },
      { "Этап упаковки фасада": "", "Статус упаковки фасада": "" }, "Завершен");
    add(42, "206RM25|Тимофеева|Гардероб", "Монтаж",
      { ldspStage: "Монтаж", hasMdfFacade: false, mdfStage: "В покраску" },
      { "Этап МДФ": "", "Статус МДФ": "", "Этап Малярка МДФ": "",
        "Статус Малярка МДФ": "", "Операция Малярка МДФ": "" }, "Монтаж");
    add(52, "206RM|Тимофеева|Санузлы", "Покраска тыльной стороны",
      { ldspStage: "", hasMdfFacade: true, paintingOperationCode: "backPainting", mdfPaintingStage: "" },
      { "Этап ЛДСП": "Упаковка Корпуса", "Статус ЛДСП": "", "Этап МДФ": "",
        "Статус МДФ": "Готово", "Этап Малярка МДФ": "Покраска тыльной стороны",
        "Статус Малярка МДФ": "", "Операция Малярка МДФ": "backPainting",
        "Операция Малярки": "" }, "В покраску");

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const rows = SheetService.getData(CONFIG.SHEETS.ORDERS);
      const headers = rows[0].map(value => String(value || "").trim());
      const report = { dryRun: Boolean(dryRun), changes: [], alreadyApplied: [], conflicts: [] };
      const queued = [];
      specs.forEach(spec => {
        const row = rows[spec.row - 1];
        if (!row) {
          report.conflicts.push({ row: spec.row, key: spec.key, reason: "Строка отсутствует." });
          return;
        }
        const raw = {};
        headers.forEach((header, index) => { raw[header] = row[index]; });
        const order = OrderMapper.map(raw);
        const key = String(order.key || "").split("|").map(value => value.trim()).join("|");
        if (key !== spec.key) {
          report.conflicts.push({ row: spec.row, key: spec.key, reason: "Ключ строки изменился." });
          return;
        }
        const projected = this.projectRouteOrder(order, spec.fields);
        const overall = this.resolveOverallStage(projected);
        if (!overall || overall.stage !== spec.newStage) {
          report.conflicts.push({ row: spec.row, key, reason: "Расчёт общего этапа не совпал с проверенным." });
          return;
        }
        const fields = Object.assign({}, spec.fields, {
          "Этап": overall.stage, "Статус": overall.status, "Ответственный": "Производство"
        });
        const already = Object.entries(fields).every(([name, value]) =>
          String(raw[name] == null ? "" : raw[name]).trim() === String(value).trim());
        if (already) {
          report.alreadyApplied.push({ row: spec.row, key });
          return;
        }
        const matches = String(order.stage || "").trim() === spec.oldStage &&
          Object.entries(spec.expected).every(([name, value]) =>
            String(order[name] == null ? "" : order[name]).trim() === String(value).trim());
        if (!matches) {
          report.conflicts.push({ row: spec.row, key, reason: "Исходное состояние веток изменилось." });
          return;
        }
        report.changes.push({ row: spec.row, key, from: order.stage, to: overall.stage, fields });
        queued.push({ key: order.key, fields });
      });
      if (!dryRun) {
        if (report.conflicts.length) throw new Error("Миграция остановлена: " + JSON.stringify(report.conflicts));
        queued.forEach(item => SheetService.updateFields(item.key, item.fields));
      }
      return report;
    } finally {
      lock.releaseLock();
    }
  }

  // Одноразовое исправление рекламаций, созданных до правила, по которому
  // передача в рекламацию завершает монтаж выбранной ветки.
  static migrateClaimedInstallationBranchesToFinished() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row));
      const migrated = [];
      orders.forEach(order => {
        if (String(order.stage || "").trim() !== "Рекламация") return;
        const fields = {};
        const branches = [];
        this.getInstallationStreams(order).forEach(stream => {
          const config = this.getInstallationBranchConfig(stream.taskType);
          if (!config || String(stream.stage || "").trim() !== "Монтаж" ||
            String(stream.status || "").trim().toLowerCase() !== "рекламация") return;
          fields[config.stageField] = "Завершен";
          fields[config.statusField] = "Готово";
          branches.push(stream.taskType);
        });
        if (!branches.length) return;
        SheetService.updateFields(order.key, fields);
        migrated.push({ key: order.key, branches: branches });
      });
      return { migrated: migrated };
    } finally {
      lock.releaseLock();
    }
  }

  // Для старого изделия без корпусного поля единый расчёт недоступен.
  // Завершение его исторического потока по-прежнему обновляет общий этап.
  static getLegacyCompletionFields(order, completedBranch) {
    const streams = [
      ["ldsp", order.ldspStage], ["corpus-mdf", order.corpusMdfStage], ["painted-ldsp", order.paintedLdspStage],
      ["painting-corpus-mdf", order.corpusMdfPaintingStage], ["painting-painted-ldsp", order.paintedLdspPaintingStage],
      ["mdf", order.mdfStage], ["veneer", order.veneerStage],
      ["facade-packaging", order.facadePackagingStage],
      ["facade-packaging-mdf", order.mdfFacadePackagingStage],
      ["facade-packaging-veneer", order.veneerFacadePackagingStage],
      ["painting-mdf", order.mdfPaintingStage], ["painting-veneer", order.veneerPaintingStage]
    ];
    const active = streams.some(([branch, stage]) => branch !== completedBranch &&
      String(stage || "").trim() && String(stage || "").trim() !== "Завершен");
    return { "Этап": active ? "Монтаж" : "Завершен", "Статус": active ? "" : "Готово",
      "Ответственный": "Производство" };
  }

  // Единственная проекция монтажных потоков. Общий этап изделия здесь не является
  // операцией, пока существуют отдельные ветки корпуса или фасада.
  static getInstallationStreams(order) {
    const normalize = value => String(value || "").trim();
    const streams = [];
    const add = (taskType, stage, status, flow) => {
      if (normalize(stage)) streams.push({
        taskType,
        stage: normalize(stage),
        status: normalize(status),
        flow,
        // До появления отдельных колонок старый общий комментарий остаётся
        // доступен только как совместимое значение для уже начатых маршрутов.
        comment: this.getInstallationComment(order, taskType)
      });
    };
    const hasIndependentFacadeState = Boolean(
      normalize(order.mdfPaintingStage) || normalize(order.veneerPaintingStage) ||
      normalize(order.mdfFacadePackagingStage) || normalize(order.veneerFacadePackagingStage)
    );

    add("ldsp", order.ldspStage, order.ldspStatus, "corpus");
    add("corpus-mdf", order.corpusMdfStage || order.corpusMdfPaintingStage,
      order.corpusMdfStage ? order.corpusMdfStatus : order.corpusMdfPaintingStatus, "corpus-mdf");
    add("painted-ldsp", order.paintedLdspStage || order.paintedLdspPaintingStage,
      order.paintedLdspStage ? order.paintedLdspStatus : order.paintedLdspPaintingStatus, "painted-ldsp");
    if (normalize(order.facadePackagingStage) && !hasIndependentFacadeState) {
      add("facade-packaging", order.facadePackagingStage, order.facadePackagingStatus, "facade");
    } else {
      // Устаревший этап МДФ/Шпон не создаёт фасад сам по себе: нужен материал
      // либо новое независимое состояние ветки.
      if (order.hasMdfFacade || normalize(order.mdfPaintingStage) || normalize(order.mdfFacadePackagingStage)) {
        const mdfTaskType = normalize(order.mdfFacadePackagingStage) ? "facade-packaging-mdf" : "mdf";
        add(mdfTaskType,
          order.mdfFacadePackagingStage || order.mdfPaintingStage || order.mdfStage,
          order.mdfFacadePackagingStatus || order.mdfPaintingStatus || order.mdfStatus,
          "facade");
      }
      if (order.hasVeneerFacade || normalize(order.veneerPaintingStage) || normalize(order.veneerFacadePackagingStage)) {
        const veneerTaskType = normalize(order.veneerFacadePackagingStage) ? "facade-packaging-veneer" : "veneer";
        add(veneerTaskType,
          order.veneerFacadePackagingStage || order.veneerPaintingStage || order.veneerStage,
          order.veneerFacadePackagingStatus || order.veneerPaintingStatus || order.veneerStatus,
          "facade");
      }
    }
    if (!this.hasManufacturingTasks(order) && ["Монтаж", "Завершен"].includes(normalize(order.stage))) {
      add("legacy-installation", order.stage, order.status, "legacy");
    }
    return streams;
  }

  static getInstallationCommentField(branch) {
    return {
      ldsp: "Комментарий монтажа корпуса",
      "corpus-mdf": "Комментарий монтажа корпуса МДФ",
      "painted-ldsp": "Комментарий монтажа крашеного ЛДСП",
      "facade-packaging": "Комментарий монтажа фасада",
      "facade-packaging-mdf": "Комментарий монтажа фасада МДФ",
      mdf: "Комментарий монтажа фасада МДФ",
      "facade-packaging-veneer": "Комментарий монтажа фасада Шпон",
      veneer: "Комментарий монтажа фасада Шпон",
      "legacy-installation": "Комментарий монтажа legacy"
    }[String(branch || "").trim().toLowerCase()] || "";
  }

  static getInstallationComment(order, branch) {
    const property = {
      ldsp: "installationCorpusComment",
      "corpus-mdf": "installationCorpusMdfComment",
      "painted-ldsp": "installationPaintedLdspComment",
      "facade-packaging": "installationFacadeComment",
      "facade-packaging-mdf": "installationMdfComment",
      mdf: "installationMdfComment",
      "facade-packaging-veneer": "installationVeneerComment",
      veneer: "installationVeneerComment",
      "legacy-installation": "installationLegacyComment"
    }[String(branch || "").trim().toLowerCase()];
    return String((property && order[property]) || order.comment || "").trim();
  }

  static getInstallationBranchConfig(branch) {
    return {
      ldsp: { stageField: "Этап ЛДСП", statusField: "Статус ЛДСП" },
      "corpus-mdf": { stageField: "Этап Корпус МДФ", statusField: "Статус Корпус МДФ" },
      "painted-ldsp": { stageField: "Этап Крашеный ЛДСП", statusField: "Статус Крашеный ЛДСП" },
      "facade-packaging": { stageField: "Этап упаковки фасада", statusField: "Статус упаковки фасада" },
      "facade-packaging-mdf": { stageField: "Этап упаковки фасада МДФ", statusField: "Статус упаковки фасада МДФ" },
      mdf: { stageField: "Этап МДФ", statusField: "Статус МДФ" },
      "facade-packaging-veneer": { stageField: "Этап упаковки фасада Шпон", statusField: "Статус упаковки фасада Шпон" },
      veneer: { stageField: "Этап Шпон", statusField: "Статус Шпон" },
      "legacy-installation": { stageField: "Этап", statusField: "Статус" }
    }[String(branch || "").trim().toLowerCase()] || null;
  }

  // Старые записи могли хранить статус одной фасадной ветки в одном из
  // совместимых полей. При возврате из «Завершен» очищаем все такие статусы,
  // чтобы «Монтаж» не попал ошибочно в роль «Сданы».
  static getRouteStatusResetFields(branch) {
    const fields = {
      ldsp: ["Статус ЛДСП"],
      "corpus-mdf": ["Статус Корпус МДФ", "Статус Малярка Корпус МДФ"],
      "painted-ldsp": ["Статус Крашеный ЛДСП", "Статус Малярка Крашеный ЛДСП"],
      "facade-packaging": ["Статус упаковки фасада", "Статус МДФ", "Статус Шпон"],
      "facade-packaging-mdf": ["Статус упаковки фасада МДФ", "Статус МДФ", "Статус Малярка МДФ"],
      mdf: ["Статус МДФ", "Статус упаковки фасада МДФ", "Статус Малярка МДФ"],
      "facade-packaging-veneer": ["Статус упаковки фасада Шпон", "Статус Шпон", "Статус Малярка Шпон"],
      veneer: ["Статус Шпон", "Статус упаковки фасада Шпон", "Статус Малярка Шпон"],
      "legacy-installation": ["Статус"]
    }[String(branch || "").trim().toLowerCase()] || [];
    return fields.reduce((result, field) => {
      result[field] = "";
      return result;
    }, {});
  }

  static getClaimOverview(order) {
    const normalize = value => String(value || "").trim();
    const streams = this.getInstallationStreams(order);
    const isClaim = stream => normalize(stream && stream.stage) === "Монтаж" &&
      normalize(stream && stream.status).toLowerCase() === "рекламация";
    const isFinished = stream => normalize(stream && stream.stage) === "Завершен" &&
      normalize(stream && stream.status).toLowerCase() === "готово";
    const describe = stream => {
      if (!stream) return "Не поступил в монтаж";
      if (isClaim(stream)) return "Монтаж";
      if (isFinished(stream)) return "Завершен";
      return normalize(stream.stage) || "Не поступил в монтаж";
    };
    const take = taskTypes => streams.find(stream => taskTypes.includes(stream.taskType));
    const corpus = take(["ldsp", "legacy-installation"]);
    const corpusMdf = take(["corpus-mdf"]);
    const paintedLdsp = take(["painted-ldsp"]);
    const legacyFacade = take(["facade-packaging"]);
    // В исторической записи общий фасадный поток был единственным. Если
    // выбран только МДФ и отдельная ветка ещё не заведена, это и есть МДФ.
    const mdf = take(["facade-packaging-mdf", "mdf"]) ||
      (order.hasMdfFacade && !order.hasVeneerFacade ? legacyFacade : null);
    const veneer = take(["facade-packaging-veneer", "veneer"]);
    const required = [];
    if (order.hasCorpusLdsp !== false) required.push(corpus || null);
    if (order.hasCorpusMdf) required.push(corpusMdf || null);
    if (order.hasPaintedLdsp) required.push(paintedLdsp || null);
    if (order.hasMdfFacade) required.push(mdf || null);
    if (order.hasVeneerFacade) required.push(veneer || null);
    if (!order.hasMdfFacade && !order.hasVeneerFacade && legacyFacade) required.push(legacyFacade);
    const claimReady = order.manualClaim || (required.length > 0 && required.every(stream => isClaim(stream) || isFinished(stream)));
    return {
      corpus: order.hasCorpusLdsp === false ? "—" : describe(corpus),
      corpusMdf: order.hasCorpusMdf ? describe(corpusMdf) : "—",
      paintedLdsp: order.hasPaintedLdsp ? describe(paintedLdsp) : "—",
      mdf: order.hasMdfFacade ? describe(mdf) : "—",
      veneer: order.hasVeneerFacade ? describe(veneer) : "—",
      legacyFacade: !order.hasMdfFacade && !order.hasVeneerFacade && legacyFacade ? describe(legacyFacade) : "—",
      claimReady,
      pendingCount: order.manualClaim ? 0 : required.filter(stream => !(isClaim(stream) || isFinished(stream))).length
    };
  }

  static isEligibleForManualClaim(order) {
    if (String(order.stage || "").trim() === "Рекламация") return false;
    const overview = this.getClaimOverview(order);
    const statuses = [overview.corpus, overview.corpusMdf, overview.paintedLdsp, overview.mdf, overview.veneer, overview.legacyFacade].filter(value => value !== "—");
    return !order.manualClaim && overview.claimReady && statuses.length > 0 && statuses.every(value => value === "Завершен");
  }

  static getClaimCandidates() {
    return SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row))
      .filter(order => this.isEligibleForManualClaim(order))
      .map(order => ({ key: order.key, contract: order.contract, client: order.client, product: order.product }));
  }

  static addClaim(data) {
    const comment = String(data.comment || "").trim();
    if (data.mode === "new") {
      const order = { contract: String(data.contract || "").trim(), client: String(data.client || "").trim(), product: String(data.product || "").trim(), manager: String(data.manager || "").trim(), constructor: "", duration: 0, manualClaim: true, claimComment: comment };
      if (!order.client || !order.product || !order.manager) throw new Error("Заполните клиента, изделие и менеджера.");
      return SheetService.appendOrder(order);
    }
    if (data.mode !== "existing") throw new Error("Выберите способ добавления рекламации.");
    if (!comment) throw new Error("Укажите суть рекламации.");
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(data.key);
      if (!this.isEligibleForManualClaim(order)) throw new Error("Изделие уже в рекламации или монтаж не завершён.");
      SheetService.ensureOrderColumn("Комментарий рекламации");
      const fields = { "Этап": "Рекламация", "Статус": "", "Ответственный": "Производство", "Комментарий рекламации": comment };
      // Историческая запись не имеет отдельных веток: после смены общего
      // этапа восстановить её завершённый монтаж из потоков невозможно.
      if (!this.hasManufacturingTasks(order)) {
        SheetService.ensureOrderColumn("Ручная рекламация");
        fields["Ручная рекламация"] = true;
      }
      SheetService.updateFields(data.key, fields);
      return { success: true };
    } finally { lock.releaseLock(); }
  }

  static getInstallationTasks(orders) {
    const installationStages = new Set(["Монтаж", "Завершен"]);
    const seen = new Set();
    return orders.reduce((tasks, order) => {
      const isClaimOrder = String(order.stage || "").trim() === "Рекламация";
      this.getInstallationStreams(order).forEach(stream => {
        if (!installationStages.has(stream.stage)) return;
        // После передачи одной ветки в рекламацию она уже завершена и не
        // должна появляться в «Сданных». Незавершённые соседние ветки того же
        // изделия остаются в текущем монтаже до собственного завершения.
        if (isClaimOrder && String(stream.stage || "").trim() === "Завершен" &&
          String(stream.status || "").trim().toLowerCase() === "готово") return;
        const taskKey = order.key + "|" + stream.taskType;
        if (seen.has(taskKey)) return;
        seen.add(taskKey);
        tasks.push(Object.assign({}, order.toJSON(), stream, {
          installationStream: true,
          queue: Number(order.productionQueue || order.queue || 0)
        }));
      });
      return tasks;
    }, []);
  }

  static getInstallationOverviewOrders(orders) {
    const installationStages = new Set(["Монтаж", "Завершен"]);
    const normalize = value => String(value || "").trim();
    const isComplete = stream => normalize(stream.stage) === "Завершен" ||
      (normalize(stream.stage) === "Монтаж" && normalize(stream.status).toLowerCase() === "готово");
    const isAtInstallation = stream => installationStages.has(normalize(stream.stage));

    return orders.reduce((items, order) => {
      if (normalize(order.stage) === "Рекламация") return items;
      const streams = this.getInstallationStreams(order);
      const body = streams.find(stream => stream.taskType === "ldsp");
      const corpusMdf = streams.find(stream => stream.taskType === "corpus-mdf");
      const paintedLdsp = streams.find(stream => stream.taskType === "painted-ldsp");
      const mdfFacade = streams.find(stream => ["mdf", "facade-packaging-mdf"].includes(stream.taskType));
      const veneerFacade = streams.find(stream => ["veneer", "facade-packaging-veneer"].includes(stream.taskType));
      const facades = streams.filter(stream => stream.flow === "facade");
      const legacy = streams.find(stream => stream.flow === "legacy");
      const reachedInstallation = streams.some(isAtInstallation);
      if (!reachedInstallation) return items;

      const statusOf = stream => !stream ? "—" : (isComplete(stream) ? "Готово" : "В работе");
      const bodyComplete = legacy ? isComplete(legacy) : (order.hasCorpusLdsp === false ? true : isComplete(body || {}));
      const bodyStatus = legacy
        ? (bodyComplete ? "Готово" : "В работе")
        : order.hasCorpusLdsp === false ? "—" : statusOf(body);
      const facadesAtInstallation = facades.filter(isAtInstallation);
      const facadeStatus = facadesAtInstallation.length === 0
        ? "—"
        : facades.every(isComplete) ? "Готово" : "В работе";
      // Прочерк лишь означает, что поток не поступал в монтаж. Он не делает
      // изделие готовым: завершение по-прежнему определяется всеми ветками.
      const complete = bodyComplete &&
        (!order.hasCorpusMdf || isComplete(corpusMdf || {})) &&
        (!order.hasPaintedLdsp || isComplete(paintedLdsp || {})) &&
        (!order.hasMdfFacade || isComplete(mdfFacade || {})) &&
        (!order.hasVeneerFacade || isComplete(veneerFacade || {})) &&
        (!order.hasMdfFacade && !order.hasVeneerFacade ? (facades.length === 0 || facades.every(isComplete)) : true);

      items.push(Object.assign({}, order.toJSON(), {
        key: order.key + "::installation-overview",
        orderKey: order.key,
        installationOverview: true,
        corpusInstallationStatus: bodyStatus,
        corpusMdfInstallationStatus: order.hasCorpusMdf ? statusOf(corpusMdf) : "—",
        paintedLdspInstallationStatus: order.hasPaintedLdsp ? statusOf(paintedLdsp) : "—",
        mdfInstallationStatus: order.hasMdfFacade ? statusOf(mdfFacade) : "—",
        veneerInstallationStatus: order.hasVeneerFacade ? statusOf(veneerFacade) : "—",
        facadeInstallationStatus: facadeStatus,
        installationComplete: complete,
        weekday: ""
      }));
      return items;
    }, []);
  }

  // Ручное, адресное исправление старой ошибочной записи. Метод намеренно не
  // вызывается при чтении snapshot: этап меняется только по явной команде.
  static repairCorpusPaintingLeak(key) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = SheetService.getObjects(CONFIG.SHEETS.ORDERS)
        .map(row => OrderMapper.map(row))
        .find(item => item.key === key);
      if (!order) throw new Error("Изделие не найдено.");
      const isCorpusLeak = String(order.stage || "").trim() === "В покраску" &&
        !order.hasMdfFacade && !order.hasVeneerFacade &&
        (order.hasLdspFacade || order.hasNoFacade) &&
        !String(order.mdfPaintingStage || "").trim() &&
        !String(order.veneerPaintingStage || "").trim();
      if (!isCorpusLeak) {
        throw new Error("Запись не является корпусом, ошибочно направленным в Малярку.");
      }
      SheetService.updateFields(order.key, {
        "Этап": "Упаковка",
        "Статус": "",
        "Ответственный": "Производство",
        "Этап ЛДСП": "Упаковка Корпуса",
        "Статус ЛДСП": "",
        "Операция Малярки": "",
        "Маляр": "",
        "Шлифовщик": ""
      });
      return this.getOrders("production");
    } finally {
      lock.releaseLock();
    }
  }

  // Одноразовый перенос фасадов, которые были переданы в Малярку до появления
  // отдельных полей ветки. Список намеренно адресный, чтобы не затронуть
  // завершённые legacy-маршруты и независимый корпус.
  static migratePendingMdfPaintingBranches() {
    const targets = new Set([
      "27RM|Карташова|Кухня",
      "39RM|Малыш|Прихожая",
      "40RM|Малыш|Шкаф детская",
      "41RM|Малыш|Шкаф спальня",
      "206RM25|Тимофеева|Кабинет",
      "21RM26|Каспарьян|Кухня"
    ]);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row));
      const report = { migrated: [], skipped: [] };
      orders.filter(order => targets.has(order.key)).forEach(order => {
        const ready = order.hasMdfFacade &&
          String(order.mdfStage || "").trim() === "В покраску" &&
          !String(order.mdfPaintingStage || "").trim() &&
          !String(order.mdfPaintingStatus || "").trim() &&
          !String(order.mdfPaintingOperationCode || "").trim();
        if (!ready) {
          report.skipped.push({ key: order.key, reason: "Состояние уже изменено или не соответствует условиям миграции." });
          return;
        }
        SheetService.updateFields(order.key, {
          "Этап МДФ": "",
          "Статус МДФ": "Готово",
          "Этап Малярка МДФ": "В покраску",
          "Статус Малярка МДФ": "",
          "Операция Малярка МДФ": "",
          "Этап": "В покраску",
          "Статус": ""
        });
        report.migrated.push(order.key);
      });
      targets.forEach(key => {
        if (!orders.some(order => order.key === key)) {
          report.skipped.push({ key, reason: "Изделие не найдено." });
        }
      });
      return report;
    } finally {
      lock.releaseLock();
    }
  }

  // Адресный возврат старых МДФ-фасадов из legacy-упаковки в Буфер Малярки.
  // Метод не вызывается при чтении очереди: изменение возможно только явно.
  static migrateLegacyFacadePackagingToMdfBuffer() {
    const targets = new Set([
      "36RM|Шагал|С/У гост",
      "23RM|Татанов|Стеллаж",
      "23RM|Татанов|Тумбы к стеллажу",
      "35RM|Пядухов|С/У",
      "11RM26|Татанов|Рабоч.зона Мастер-Спальня №2",
      "11RM26|Татанов|Шкаф Спальня №3 (Сын)",
      "34RM26|Татанов|ГАРДЕРОБНАЯ",
      "11RM26|Татанов|Спальня №1 (Дочь)",
      "11RM26|Татанов|Шкафы Мастер Спальня №2"
    ]);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row));
      const report = { migrated: [], skipped: [] };
      orders.filter(order => targets.has(order.key)).forEach(order => {
        const mdfStatus = String(order.mdfStatus || "").trim();
        const ready = order.hasMdfFacade &&
          String(order.stage || "").trim() === "Упаковка" &&
          String(order.facadePackagingStage || "").trim() === "Упаковка Фасада" &&
          !String(order.mdfStage || "").trim() &&
          (!mdfStatus || mdfStatus === "Готово") &&
          !String(order.mdfPaintingStage || "").trim() &&
          !String(order.mdfPaintingStatus || "").trim() &&
          !String(order.mdfPaintingOperationCode || "").trim() &&
          !String(order.mdfFacadePackagingStage || "").trim() &&
          !String(order.mdfFacadePackagingStatus || "").trim();
        if (!ready) {
          report.skipped.push({ key: order.key, reason: "Состояние уже изменено или не соответствует условиям миграции." });
          return;
        }
        SheetService.updateFields(order.key, {
          "Этап МДФ": "",
          "Статус МДФ": "Готово",
          "Этап упаковки фасада": "",
          "Статус упаковки фасада": "",
          "Этап Малярка МДФ": "В покраску",
          "Статус Малярка МДФ": "",
          "Операция Малярка МДФ": "",
          "Этап": "В покраску",
          "Статус": "",
          "Ответственный": "Производство"
        });
        report.migrated.push(order.key);
      });
      targets.forEach(key => {
        if (!orders.some(order => order.key === key)) {
          report.skipped.push({ key, reason: "Изделие не найдено." });
        }
      });
      return report;
    } finally {
      lock.releaseLock();
    }
  }

  // Запускает выбранную независимую ветку из Буфера. Общая запись изделия
  // остаётся сводным индикатором, а не заменой состояния потока.
  static startPaintingBranch(key, branch) {
    const normalizedBranch = String(branch || "").trim().toLowerCase();
    if (!["mdf", "veneer", "corpus-mdf", "painted-ldsp"].includes(normalizedBranch)) {
      throw new Error("Выберите ветку для запуска в Малярке.");
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);
      if (!order) throw new Error("Изделие не найдено. Обновите очередь и повторите действие.");
      const state = this.getPaintingBranchState(order, normalizedBranch);
      if (String(state.stage || "").trim() !== "В покраску") {
        throw new Error("Фасадная ветка уже запущена или не находится в Буфере.");
      }
      const branchOrder = this.getPaintingBranchOrder(order, normalizedBranch);
      this.validateProductionDays(branchOrder);
      const firstOperation = this.getActivePaintingOperations(branchOrder)[0];
      if (!firstOperation) throw new Error("В маршруте Малярки нет активных операций.");
      const painter = this.assignPainter(orders).painter;
      const sanding = this.assignSandingWorker(orders).worker;
      this.updateRouteFields(key, order, {
        [state.stageField]: firstOperation.stage,
        [state.statusField]: "",
        [state.operationField]: firstOperation.taskType,
        "Этап": firstOperation.stage,
        "Статус": "",
        "Маляр": painter,
        "Шлифовщик": sanding,
        "Плановая дата покраски": "",
        "Плановая дата шлифовки": "",
        "Малярка — ручной приоритет": false
      });
      return { success: true, branch: normalizedBranch, stage: firstOperation.stage };
    } finally {
      lock.releaseLock();
    }
  }

  // Ручной аварийный переход из Буфера: фасад передаётся в упаковку, а новый
  // корпусный поток возвращается в собственную сборку.
  static sendPaintingBranchToPackaging(key, branch) {
    const normalizedBranch = String(branch || "").trim().toLowerCase();
    if (!["mdf", "veneer", "corpus-mdf", "painted-ldsp"].includes(normalizedBranch)) {
      throw new Error("Выберите ветку, находящуюся в Буфере Малярки.");
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);
      if (!order) throw new Error("Изделие не найдено. Обновите очередь и повторите действие.");

      const state = this.getPaintingBranchState(order, normalizedBranch);
      if (String(state.stage || "").trim() !== "В покраску") {
        throw new Error("Ветку можно передать в упаковку только из Буфера Малярки.");
      }
      const isCorpus = ["corpus-mdf", "painted-ldsp"].includes(normalizedBranch);
      const config = isCorpus ? this.getAdditionalCorpusBranchConfig(normalizedBranch) : null;
      const fields = {
        [state.stageField]: "",
        [state.statusField]: "",
        [state.operationField]: ""
      };
      if (isCorpus) {
        fields[config.stageField] = "Сборка";
        fields[config.statusField] = "";
      } else {
        fields[state.packagingStageField] = "Упаковка Фасада";
        fields[state.packagingStatusField] = "";
      }
      this.updateRouteFields(key, order, fields);
      return { success: true, branch: normalizedBranch, stage: isCorpus ? "Сборка" : "Упаковка Фасада" };
    } finally {
      lock.releaseLock();
    }
  }

  static getManufacturingTasks(orders) {
    const tasks = [];
    orders.forEach(order => {
      [
        { branch: "ldsp", stage: order.ldspStage, status: order.ldspStatus, offset: 0 },
        { branch: "corpus-mdf", stage: order.corpusMdfStage, status: order.corpusMdfStatus, offset: 0.4 },
        { branch: "painted-ldsp", stage: order.paintedLdspStage, status: order.paintedLdspStatus, offset: 0.5 },
        { branch: "preparation-corpus-mdf", stage: order.corpusMdfPaintingStage, status: order.corpusMdfPaintingStatus, offset: 0.45, onlyStage: "Подготовка" },
        { branch: "mdf", stage: order.mdfStage, status: order.mdfStatus, offset: 1 },
        { branch: "veneer", stage: order.veneerStage, status: order.veneerStatus, offset: 2 },
        { branch: "preparation-mdf", stage: order.mdfPaintingStage, status: order.mdfPaintingStatus, offset: 2.1, onlyStage: "Подготовка" },
        { branch: "preparation-veneer", stage: order.veneerPaintingStage, status: order.veneerPaintingStatus, offset: 2.2, onlyStage: "Подготовка" },
        // Историческая фасадная упаковка живёт в собственном поле. Не подменяем
        // её МДФ-веткой: иначе действие из очереди ищет операцию не там.
        { branch: "facade-packaging", stage: order.facadePackagingStage, status: order.facadePackagingStatus, offset: 3 },
        { branch: "facade-packaging-mdf", stage: order.mdfFacadePackagingStage, status: order.mdfFacadePackagingStatus, offset: 4 },
        { branch: "facade-packaging-veneer", stage: order.veneerFacadePackagingStage, status: order.veneerFacadePackagingStatus, offset: 5 }
      ].forEach(operation => {
        if (!String(operation.stage || "").trim()) return;
        if (operation.onlyStage && String(operation.stage || "").trim() !== operation.onlyStage) return;
        if (operation.branch === "facade-packaging" && (
          String(order.mdfPaintingStage || "").trim() || String(order.veneerPaintingStage || "").trim() ||
          String(order.mdfFacadePackagingStage || "").trim() || String(order.veneerFacadePackagingStage || "").trim()
        )) return;
        tasks.push(Object.assign({}, order.toJSON(), {
          taskType: operation.branch,
          stage: operation.stage,
          status: operation.status,
          queue: Number(order.productionQueue || order.queue || 0) + operation.offset / 10
        }));
      });
      // Старый единый маршрут мог остановиться на «Подготовке» до появления
      // независимых фасадных веток. Показываем его в новом участке без записи
      // или миграции листа.
      if (String(order.stage || "").trim() === "Подготовка" &&
          !String(order.mdfPaintingStage || "").trim() && !String(order.veneerPaintingStage || "").trim()) {
        tasks.push(Object.assign({}, order.toJSON(), {
          taskType: "preparation-legacy",
          stage: "Подготовка",
          status: order.status,
          queue: Number(order.productionQueue || order.queue || 0) + 0.29
        }));
      }
    });
    return tasks;
  }

  // Обзор мастера строится из фактических веток, а не из общего этапа изделия.
  // В выдаче ветки сводятся в одну строку изделия, но сохраняют независимые состояния.
  static getMasterWorkTasks(orders) {
    const tasks = [];
    const completedStages = new Set(["Монтаж", "Завершен"]);
    const activeLegacyStages = new Set([
      "Запуск", "Распил", "Распил ЛДСП", "Распил МДФ", "Распил Шпон",
      "Кромка", "Кромка Шпон", "ЧПУ", "Присадка", "Присадка Шпон", "Фрезеровка",
      "Сборка", "В покраску", "Упаковка", "Упаковка Корпуса", "Упаковка Фасада",
      "СГП", "СГП Корпуса", "СГП Фасада"
    ].concat(this.getPaintingOperations().map(operation => String(operation.stage || "").trim())));
    const firstActiveState = states => states.find(state => String(state.stage || "").trim()) || { stage: "", status: "" };
    const add = (order, productFlow, stage, status, taskType, offset) => {
      const normalizedStage = String(stage || "").trim();
      if (!normalizedStage || completedStages.has(normalizedStage)) return;
      tasks.push(Object.assign({}, order.toJSON(), {
        key: `${order.key}::master:${taskType}`,
        orderKey: order.key,
        taskType,
        productFlow,
        stage: normalizedStage,
        status: String(status || "").trim(),
        queue: Number(order.productionQueue || order.queue || 0) + offset / 10
      }));
    };

    orders.forEach(order => {
      let hasActiveFlow = false;
      const addFlow = (productFlow, stage, status, taskType, offset) => {
        const before = tasks.length;
        add(order, productFlow, stage, status, taskType, offset);
        hasActiveFlow = hasActiveFlow || tasks.length > before;
      };
      addFlow("Корпус", order.ldspStage, order.ldspStatus, "ldsp", 0);
      const corpusMdfState = firstActiveState([
        { stage: order.corpusMdfPaintingStage, status: order.corpusMdfPaintingStatus },
        { stage: order.corpusMdfStage, status: order.corpusMdfStatus }
      ]);
      const paintedLdspState = firstActiveState([
        { stage: order.paintedLdspPaintingStage, status: order.paintedLdspPaintingStatus },
        { stage: order.paintedLdspStage, status: order.paintedLdspStatus }
      ]);
      addFlow("Корпус МДФ", corpusMdfState.stage, corpusMdfState.status, "corpus-mdf", 0.4);
      addFlow("Крашеный ЛДСП", paintedLdspState.stage, paintedLdspState.status, "painted-ldsp", 0.5);

      const hasIndependentFacadeState = Boolean(
        String(order.mdfPaintingStage || "").trim() || String(order.veneerPaintingStage || "").trim() ||
        String(order.mdfFacadePackagingStage || "").trim() || String(order.veneerFacadePackagingStage || "").trim()
      );
      const mdfState = firstActiveState([
        { stage: order.mdfFacadePackagingStage, status: order.mdfFacadePackagingStatus },
        { stage: order.mdfPaintingStage, status: order.mdfPaintingStatus },
        { stage: order.mdfStage, status: order.mdfStatus }
      ]);
      const veneerState = firstActiveState([
        { stage: order.veneerFacadePackagingStage, status: order.veneerFacadePackagingStatus },
        { stage: order.veneerPaintingStage, status: order.veneerPaintingStatus },
        { stage: order.veneerStage, status: order.veneerStatus }
      ]);
      addFlow("Фасад МДФ", mdfState.stage, mdfState.status, "mdf", 1);
      addFlow("Фасад Шпон", veneerState.stage, veneerState.status, "veneer", 2);

      if (!hasIndependentFacadeState) {
        const legacyFacadeFlow = order.hasVeneerFacade && !order.hasMdfFacade ? "Фасад Шпон" : "Фасад МДФ";
        addFlow(legacyFacadeFlow, order.facadePackagingStage, order.facadePackagingStatus, "facade-packaging", 3);
      }

      // Исторические/только что запущенные изделия ещё не имеют веток. Показываем
      // их один раз, не угадывая тип, чтобы мастер не терял их из сводки.
      if (!hasActiveFlow && activeLegacyStages.has(String(order.stage || "").trim())) {
        addFlow("Не определён", order.stage, order.status, "legacy", 9);
      }
    });
    const flowFields = {
      "Корпус": "corpusFlow",
      "Корпус МДФ": "corpusMdfFlow",
      "Крашеный ЛДСП": "paintedLdspFlow",
      "Фасад МДФ": "mdfFlow",
      "Фасад Шпон": "veneerFlow"
    };
    const grouped = new Map();
    tasks.forEach(task => {
      const orderKey = String(task.orderKey || task.key || "").trim();
      if (!grouped.has(orderKey)) {
        const row = Object.assign({}, task, {
          key: orderKey,
          orderKey,
          corpusFlow: null,
          corpusMdfFlow: null,
          paintedLdspFlow: null,
          mdfFlow: null,
          veneerFlow: null,
          unknownFlow: null
        });
        delete row.productFlow;
        delete row.stage;
        delete row.status;
        grouped.set(orderKey, row);
      }
      const row = grouped.get(orderKey);
      const flow = {
        stage: String(task.stage || "").trim(),
        status: String(task.status || "").trim()
      };
      const field = flowFields[task.productFlow];
      if (field) {
        // При старых данных может встретиться резервная ветка того же типа.
        // Первой остаётся приоритетная фактическая ветка, как и в прежнем списке.
        if (!row[field]) row[field] = flow;
      } else if (!row.unknownFlow) {
        row.unknownFlow = flow;
      }
      row.queue = Math.min(Number(row.queue || 0), Number(task.queue || 0));
    });
    // Монтаж и завершение не создают строку АРМ. Но если строка уже появилась за
    // счёт активной производственной/малярной ветки, её состояние показывается
    // в той же колонке типа изделия, а не в отдельной сводной колонке.
    const orderByKey = new Map(orders.map(order => [String(order.key || "").trim(), order]));
    const addCompletedFlow = (row, field, stage, status) => {
      const normalizedStage = String(stage || "").trim();
      if (row[field] || !completedStages.has(normalizedStage)) return;
      row[field] = {
        stage: normalizedStage,
        status: String(status || "").trim()
      };
    };
    grouped.forEach((row, orderKey) => {
      const order = orderByKey.get(orderKey);
      if (!order) return;
      addCompletedFlow(row, "corpusFlow", order.ldspStage, order.ldspStatus);
      const corpusMdfCompleted = firstActiveState([
        { stage: order.corpusMdfPaintingStage, status: order.corpusMdfPaintingStatus },
        { stage: order.corpusMdfStage, status: order.corpusMdfStatus }
      ]);
      const paintedLdspCompleted = firstActiveState([
        { stage: order.paintedLdspPaintingStage, status: order.paintedLdspPaintingStatus },
        { stage: order.paintedLdspStage, status: order.paintedLdspStatus }
      ]);
      addCompletedFlow(row, "corpusMdfFlow", corpusMdfCompleted.stage, corpusMdfCompleted.status);
      addCompletedFlow(row, "paintedLdspFlow", paintedLdspCompleted.stage, paintedLdspCompleted.status);

      const hasIndependentFacadeState = Boolean(
        String(order.mdfPaintingStage || "").trim() || String(order.veneerPaintingStage || "").trim() ||
        String(order.mdfFacadePackagingStage || "").trim() || String(order.veneerFacadePackagingStage || "").trim()
      );
      const mdfState = firstActiveState([
        { stage: order.mdfFacadePackagingStage, status: order.mdfFacadePackagingStatus },
        { stage: order.mdfPaintingStage, status: order.mdfPaintingStatus },
        { stage: order.mdfStage, status: order.mdfStatus }
      ]);
      const veneerState = firstActiveState([
        { stage: order.veneerFacadePackagingStage, status: order.veneerFacadePackagingStatus },
        { stage: order.veneerPaintingStage, status: order.veneerPaintingStatus },
        { stage: order.veneerStage, status: order.veneerStatus }
      ]);
      addCompletedFlow(row, "mdfFlow", mdfState.stage, mdfState.status);
      addCompletedFlow(row, "veneerFlow", veneerState.stage, veneerState.status);
      if (!hasIndependentFacadeState) {
        const legacyField = order.hasVeneerFacade && !order.hasMdfFacade ? "veneerFlow" : "mdfFlow";
        addCompletedFlow(row, legacyField, order.facadePackagingStage, order.facadePackagingStatus);
      }
    });
    const supplySignalsByKey = SupplyService.getMasterSupplySignals(orders);
    grouped.forEach((row, orderKey) => {
      const lookupKey = SheetService.buildOrderLookupKey(row.contract, row.client, row.product);
      row.supplySignals = supplySignalsByKey[lookupKey] || {};
    });
    return Array.from(grouped.values());
  }

  static setProformaApprovedDate(key, value) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      // Штатное согласование фиксирует дату автоматически. Поле в карточке —
      // контролируемая ручная коррекция исторических заказов, поэтому не
      // зависит от текущего этапа (у старых строк он может быть неактуален).
      const match = String(value || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (!match) throw new Error("Укажите корректную дату согласования Проформы.");
      const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
      if (date.getFullYear() !== Number(match[1]) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[3])) {
        throw new Error("Укажите корректную дату согласования Проформы.");
      }
      const storedDate = CalendarService.formatDate(date);
      SheetService.updateField(key, "Проформа Согласована", storedDate);
      return { success: true, patch: { key: key, values: { proformaApprovedDate: storedDate } } };
    } finally {
      lock.releaseLock();
    }
  }

  static setOrderDeadlineSettings(key, modeValue, manufacturingDaysValue, manualDateValue) {
    const mode = String(modeValue || "").trim();
    const modes = ["От согласования проформы", "Конкретная дата"];
    if (!modes.includes(mode)) throw new Error("Выберите способ определения дедлайна.");

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      const columns = SheetService.getOrderContext(key).columns;
      const deadlineFields = ["Дедлайн", "Способ определения дедлайна", "Срок изготовления, дней", "Дата дедлайна вручную"];
      // Миграция формул нужна только для старой структуры. На обычном клике
      // не перечитываем и не форматируем весь лист.
      if (deadlineFields.some(fieldName => !columns[fieldName])) {
        SheetService.ensureOrderDeadlineColumn();
        SheetService.resetOrderContexts();
      }

      const currentDays = Number(order.manufacturingDays) || 50;
      let manufacturingDays = currentDays;
      let manualDate = order.manualDeadlineDate || "";

      if (mode === "От согласования проформы") {
        manufacturingDays = Number(manufacturingDaysValue);
        if (!Number.isInteger(manufacturingDays) || manufacturingDays <= 0 || manufacturingDays > 3650) {
          throw new Error("Срок изготовления укажите целым числом от 1 до 3650 календарных дней.");
        }
      } else {
        const match = String(manualDateValue || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!match) throw new Error("Укажите корректную конкретную дату дедлайна.");
        const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
        if (date.getFullYear() !== Number(match[1]) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[3])) {
          throw new Error("Укажите корректную конкретную дату дедлайна.");
        }
        manualDate = CalendarService.formatDate(date);
      }

      SheetService.updateFields(key, {
        "Способ определения дедлайна": mode,
        "Срок изготовления, дней": manufacturingDays,
        "Дата дедлайна вручную": manualDate
      });
      return {
        success: true,
        patch: {
          key: key,
          values: {
            deadlineMode: mode,
            manufacturingDays: manufacturingDays,
            manualDeadlineDate: manualDate
          }
        }
      };
    } finally {
      lock.releaseLock();
    }
  }

  // Текущие операции, закреплённые за ролью «Шлифовка».
  static getSandingTasks(orders) {
    return this.getPaintingRouteSchedule(orders).filter(task => task.productionRole === "sanding");
  }

  static isSandingQueueOrder(order) {
    return this.getActivePaintingOperations(order).some(operation =>
      operation.role === "sanding" && operation.stage === String(order.stage || "").trim());
  }

  static getPaintingOperations() {
    return PaintingNormService.isHourlyEnabled()
      ? PaintingNormService.getOperations(true)
      : PaintingNormService.getLegacyOperations();
  }

  // Норматив подготовки остаётся частью профиля Малярки, но выполняется в
  // производственном участке. Этот набор нужен для проверки параметров.
  static getConfiguredPaintingOperations(order) {
    if (PaintingNormService.isHourlyEnabled()) {
      return PaintingNormService.getActiveOperationsForOrder(order);
    }
    const exclusions = (order && order.routeExclusions) || {};
    return this.getPaintingOperations().filter(operation =>
      !exclusions[operation.daysKey] &&
      (!operation.optionalGroup || (operation.optionalGroup === "lacquer" && Boolean(order && order.lacquerOperations)))
    );
  }

  static getActivePaintingOperations(order) {
    return this.getConfiguredPaintingOperations(order)
      .filter(operation => operation.taskType !== "preparation" && operation.stage !== "Подготовка");
  }

  static buildProductionRoleTasks(order, role) {
    const operations = this.getActivePaintingOperations(order);
    const currentIndex = operations.findIndex(operation => operation.stage === order.stage);
    if (currentIndex < 0 || String(order.status).toLowerCase() === "готово") return [];

    const responsible = role === "painting"
      ? this.getPainterName(order.painter)
      : this.getSandingWorkerName(order.sandingWorker);
    const queueBase = role === "painting"
      ? (order.primerQueue || order.paintingReserveQueue || 0)
      : (order.sandingQueue || order.whiteSandingReserveQueue || 0);
    const storedDate = role === "painting" ? order.paintingPlannedDate : order.sandingPlannedDate;

    return operations.slice(currentIndex).map((operation, relativeIndex) => {
      if (operation.role !== role) return null;
      const absoluteIndex = currentIndex + relativeIndex;
      const isCurrent = absoluteIndex === currentIndex;
      const previous = operations[absoluteIndex - 1];
      return Object.assign({}, order.toJSON(), {
        constructor: responsible,
        taskType: operation.taskType,
        stage: isCurrent ? operation.stage : operation.stage + " (резерв)",
        status: isCurrent ? order.status : "Резерв",
        duration: Number(order.productionDays[operation.daysKey]) || 0.5,
        queue: queueBase + absoluteIndex / 100,
        plannedDate: isCurrent ? (storedDate || "") : "",
        canStart: isCurrent,
        reserveAfterTaskType: isCurrent || !previous ? "" : previous.taskType,
        reserveOffsetDays: isCurrent || !previous
          ? 0
          : (Number(order.productionDays[previous.daysKey]) || 0.5)
      });
    }).filter(Boolean);
  }

  // В Малярке каждая строка заказа проходит единый последовательный маршрут.
  static getPaintingTasks(orders) {
    return this.getPaintingRouteSchedule(orders).filter(task => task.productionRole === "painting");
  }

  static getPaintingRouteSchedule(orders) {
    const branchOrders = this.getPaintingBranchOrders(orders);
    return PaintingNormService.isHourlyEnabled()
      ? this.getHourlyPaintingRouteSchedule(branchOrders)
      : this.getLegacyPaintingRouteSchedule(branchOrders);
  }

  // Новые ветки имеют собственное состояние. Уже начатые маршруты по старым
  // общим полям проходят планировщик без обёртки: это сохраняет их этап,
  // статус, назначенного сотрудника и сохранённую очередь.
  static getPaintingBranchOrders(orders) {
    const branches = [];
    const create = (order, branch, state) => {
      const config = this.getAdditionalCorpusBranchConfig(branch);
      const data = Object.assign({}, order.toJSON(), state, {
        key: order.key + "::paint:" + branch,
        orderKey: order.key,
        paintingBranch: branch,
        productFlow: config ? config.productFlow : branch === "veneer" ? "Фасад Шпон" : "Фасад МДФ",
        paintingArea: branch === "veneer"
          ? (Number(order.veneerArea) || Number(order.paintingArea) || 0)
          : branch === "painted-ldsp"
            ? (Number(order.ldspArea) || Number(order.paintingArea) || 0)
            : (Number(order.mdfArea) || Number(order.paintingArea) || 0)
      });
      const serialized = Object.assign({}, data);
      data.toJSON = () => Object.assign({}, serialized);
      branches.push(data);
    };
    orders.forEach(order => {
      let split = false;
      const mdfStage = String(order.mdfPaintingStage || "").trim();
      const veneerStage = String(order.veneerPaintingStage || "").trim();
      const corpusMdfStage = String(order.corpusMdfPaintingStage || "").trim();
      const paintedLdspStage = String(order.paintedLdspPaintingStage || "").trim();
      const hasPreparation = mdfStage === "Подготовка" || veneerStage === "Подготовка" || corpusMdfStage === "Подготовка";
      if (mdfStage && mdfStage !== "Подготовка") {
        create(order, "mdf", {
          stage: order.mdfPaintingStage,
          status: order.mdfPaintingStatus,
          paintingOperationCode: order.mdfPaintingOperationCode
        });
        split = true;
      }
      if (veneerStage && veneerStage !== "Подготовка") {
        create(order, "veneer", {
          stage: order.veneerPaintingStage,
          status: order.veneerPaintingStatus,
          paintingOperationCode: order.veneerPaintingOperationCode
        });
        split = true;
      }
      if (corpusMdfStage && corpusMdfStage !== "Подготовка") {
        create(order, "corpus-mdf", {
          stage: order.corpusMdfPaintingStage,
          status: order.corpusMdfPaintingStatus,
          paintingOperationCode: order.corpusMdfPaintingOperationCode
        });
        split = true;
      }
      if (paintedLdspStage) {
        create(order, "painted-ldsp", {
          stage: order.paintedLdspPaintingStage,
          status: order.paintedLdspPaintingStatus,
          paintingOperationCode: order.paintedLdspPaintingOperationCode
        });
        split = true;
      }
      // Общая старая ветка сохраняется для уже начатых маршрутов. Она не
      // дублируется после появления новых раздельных состояний.
      if (!split && !hasPreparation) {
        const stage = String(order.stage || "").trim();
        const hasPaintOperation = this.getActivePaintingOperations(order).some(operation => operation.stage === stage);
        if (stage === "В покраску" || hasPaintOperation) {
          branches.push(order);
        }
      }
    });
    return branches;
  }

  // Единый планировщик Малярки. В отличие от двух раздельных очередей,
  // он запускает только готовую следующую операцию маршрута и потому не
  // оставляет шлифовщика ждать грунтования предыдущего изделия.
  static getLegacyPaintingRouteSchedule(orders) {
    const first = { date: CalendarService.firstWorkDay(), halfDay: 0 };
    const resources = {};
    const states = [];
    const result = [];
    const copyPosition = position => ({ date: new Date(position.date), halfDay: position.halfDay });
    const comparePosition = (left, right) => {
      const leftDay = new Date(left.date).setHours(0, 0, 0, 0);
      const rightDay = new Date(right.date).setHours(0, 0, 0, 0);
      return leftDay - rightDay || Number(left.halfDay) - Number(right.halfDay);
    };
    const maxPosition = (left, right) => comparePosition(left, right) >= 0 ? copyPosition(left) : copyPosition(right);
    const datePosition = value => {
      const date = CalendarService.parseStoredDate(value);
      return date ? { date, halfDay: 0 } : copyPosition(first);
    };
    const advance = (position, duration) => {
      const next = copyPosition(position);
      const halfDays = Math.max(1, Math.round((Number(duration) || 0.5) * 2));
      let finish = copyPosition(next);
      for (let index = 0; index < halfDays; index++) {
        finish = copyPosition(next);
        next.halfDay += 1;
        if (next.halfDay === 2) {
          next.date = CalendarService.nextWorkDay(next.date);
          next.halfDay = 0;
        }
      }
      return { finish, next };
    };

    orders.forEach(order => {
      const operations = this.getActivePaintingOperations(order);
      const currentIndex = operations.findIndex(operation => operation.stage === String(order.stage || "").trim());
      if (currentIndex < 0 || String(order.status || "").trim().toLowerCase() === "готово") return;
      const preview = PaintingNormService.previewOrder(order);
      states.push({
        order,
        operations: operations.slice(currentIndex),
        index: 0,
        totalNormHours: preview.valid ? Math.round(preview.totalRawMinutes / 6) / 10 : null,
        readyAt: String(order.status || "").trim().toLowerCase() === "пауза"
          ? datePosition(order.sandingPlannedDate || order.paintingPlannedDate)
          : copyPosition(first)
      });
    });

    while (states.some(state => state.index < state.operations.length)) {
      const candidates = states.filter(state => state.index < state.operations.length).map(state => {
        const operation = state.operations[state.index];
        const responsible = operation.role === "painting"
          ? this.getPainterName(state.order.painter)
          : this.getSandingWorkerName(state.order.sandingWorker);
        const resource = operation.role + "|" + responsible;
        const availableAt = resources[resource] || first;
        const start = maxPosition(state.readyAt, availableAt);
        const isCurrent = state.index === 0;
        const active = isCurrent && String(state.order.status || "").trim().toLowerCase() === "в работе";
        const manual = Boolean(state.order.paintingManualPriority);
        return { state, operation, responsible, resource, start, active, manual };
      });
      candidates.sort((left, right) =>
        comparePosition(left.start, right.start) ||
        Number(right.active) - Number(left.active) ||
        Number(right.manual) - Number(left.manual) ||
        Number(left.state.order.productionQueue || left.state.order.queue || 0) - Number(right.state.order.productionQueue || right.state.order.queue || 0) ||
        String(left.state.order.key).localeCompare(String(right.state.order.key))
      );
      const candidate = candidates[0];
      const state = candidate.state;
      const operation = candidate.operation;
      const duration = Number(state.order.productionDays[operation.daysKey]) || 0.5;
      const scheduled = advance(candidate.start, duration);
      const isCurrent = state.index === 0;
      const task = Object.assign({}, state.order.toJSON(), {
        constructor: candidate.responsible,
        productionRole: operation.role,
        taskType: operation.taskType,
        stage: isCurrent ? operation.stage : operation.stage + " (резерв)",
        status: isCurrent ? state.order.status : "Резерв",
        duration,
        totalNormHours: state.totalNormHours,
        queue: result.length + 1,
        canStart: isCurrent,
        startDate: CalendarService.formatDate(candidate.start.date),
        finishDate: CalendarService.formatDate(scheduled.finish.date),
        readyDate: CalendarService.formatDate(scheduled.finish.date),
        weekday: CalendarService.getWeekday(candidate.start.date),
        dayPart: candidate.start.halfDay === 0 ? "Ⅰ половина" : "Ⅱ половина",
        scheduleStart: candidate.start.date,
        scheduleFinish: scheduled.finish.date
      });
      result.push(task);
      resources[candidate.resource] = scheduled.next;
      state.readyAt = scheduled.next;
      state.index += 1;
    }
    return result;
  }

  /**
   * Почасовой планировщик Малярки. Каждая операция занимает один ресурс,
   * последовательность операций изделия сохраняется, а рабочее время
   * ограничено окнами 09:00–13:00 и 14:00–18:00 по будним дням.
   */
  static getHourlyPaintingRouteSchedule(orders) {
    const config = PaintingNormService.readConfig();
    const first = this.normalizePaintingWorkMoment(new Date());
    const resources = {};
    const states = [];
    const result = [];

    orders.forEach(order => {
      const operations = PaintingNormService.getActiveOperationsForOrder(order, config);
      const currentCode = String(order.paintingOperationCode || "").trim();
      const currentIndex = operations.findIndex(operation =>
        (currentCode && operation.taskType === currentCode) ||
        (!currentCode && operation.stage === String(order.stage || "").trim())
      );
      if (currentIndex < 0 || String(order.status || "").trim().toLowerCase() === "готово") return;
      const preview = PaintingNormService.previewOrder(order);
      let readyAt = new Date(first);
      if (String(order.status || "").trim().toLowerCase() === "пауза") {
        const stored = CalendarService.parseStoredDate(order.sandingPlannedDate || order.paintingPlannedDate);
        if (stored) readyAt = this.normalizePaintingWorkMoment(stored);
      }
      states.push({
        order,
        operations: operations.slice(currentIndex),
        index: 0,
        readyAt,
        totalNormHours: preview.valid ? Math.round(preview.totalRawMinutes / 6) / 10 : null
      });
    });

    while (states.some(state => state.index < state.operations.length)) {
      const candidates = states.filter(state => state.index < state.operations.length).map(state => {
        const operation = state.operations[state.index];
        const responsible = operation.role === "painting"
          ? this.getPainterName(state.order.painter)
          : this.getSandingWorkerName(state.order.sandingWorker);
        const resource = operation.role + "|" + responsible;
        const availableAt = resources[resource] || first;
        const start = new Date(Math.max(state.readyAt.getTime(), availableAt.getTime()));
        const isCurrent = state.index === 0;
        return {
          state,
          operation,
          responsible,
          resource,
          start: this.normalizePaintingWorkMoment(start),
          active: isCurrent && String(state.order.status || "").trim().toLowerCase() === "в работе",
          manual: Boolean(state.order.paintingManualPriority)
        };
      });
      candidates.sort((left, right) =>
        left.start - right.start ||
        Number(right.active) - Number(left.active) ||
        Number(right.manual) - Number(left.manual) ||
        Number(left.state.order.productionQueue || left.state.order.queue || 0) - Number(right.state.order.productionQueue || right.state.order.queue || 0) ||
        String(left.state.order.key).localeCompare(String(right.state.order.key))
      );

      const candidate = candidates[0];
      const state = candidate.state;
      const operation = candidate.operation;
      const norm = PaintingNormService.calculateOperation(state.order, operation, config);
      // После активации неполная норма не должна возникать. Если данные были
      // изменены непосредственно в листе, строка остаётся видимой с ошибкой,
      // но не получает правдоподобную нулевую длительность.
      const durationMinutes = norm.valid ? norm.roundedMinutes : 5;
      const scheduled = this.advancePaintingWorkMinutes(candidate.start, durationMinutes);
      const isCurrent = state.index === 0;
      const startText = this.formatPaintingDateTime(candidate.start);
      const finishText = this.formatPaintingDateTime(scheduled.finish);
      result.push(Object.assign({}, state.order.toJSON(), {
        constructor: candidate.responsible,
        productionRole: operation.role,
        taskType: operation.taskType,
        paintingOperationCode: operation.taskType,
        stage: isCurrent ? operation.stage : operation.stage + " (резерв)",
        status: isCurrent ? state.order.status : "Резерв",
        duration: norm.valid ? norm.normHours : 0,
        durationMinutes,
        normRawMinutes: norm.valid ? norm.rawMinutes : null,
        normRoundedMinutes: norm.valid ? norm.roundedMinutes : null,
        normHours: norm.valid ? norm.normHours : null,
        totalNormHours: state.totalNormHours,
        normProfile: norm.valid ? norm.profileName : "",
        normError: norm.valid ? "" : norm.message,
        queue: result.length + 1,
        canStart: isCurrent,
        startDate: startText,
        finishDate: finishText,
        readyDate: finishText,
        weekday: CalendarService.getWeekday(candidate.start),
        dayPart: "",
        scheduleStart: candidate.start,
        scheduleFinish: scheduled.finish
      }));
      resources[candidate.resource] = scheduled.next;
      state.readyAt = scheduled.next;
      state.index += 1;
    }
    return result;
  }

  static normalizePaintingWorkMoment(value) {
    const date = new Date(value);
    date.setSeconds(0, 0);
    while (date.getDay() === 0 || date.getDay() === 6) {
      date.setDate(date.getDate() + 1);
      date.setHours(9, 0, 0, 0);
    }
    const minutes = date.getHours() * 60 + date.getMinutes();
    if (minutes < 9 * 60) date.setHours(9, 0, 0, 0);
    else if (minutes >= 13 * 60 && minutes < 14 * 60) date.setHours(14, 0, 0, 0);
    else if (minutes >= 18 * 60) {
      const next = CalendarService.nextWorkDay(date);
      date.setTime(next.getTime());
      date.setHours(9, 0, 0, 0);
    }
    return date;
  }

  static advancePaintingWorkMinutes(startValue, durationValue) {
    let current = this.normalizePaintingWorkMoment(startValue);
    let remaining = Math.max(5, Number(durationValue) || 5);
    let finish = new Date(current);
    while (remaining > 0) {
      const minute = current.getHours() * 60 + current.getMinutes();
      const windowEndMinute = minute < 13 * 60 ? 13 * 60 : 18 * 60;
      const available = windowEndMinute - minute;
      const used = Math.min(remaining, available);
      finish = new Date(current.getTime() + used * 60000);
      remaining -= used;
      if (remaining > 0) current = this.normalizePaintingWorkMoment(finish);
    }
    return { finish, next: this.normalizePaintingWorkMoment(finish) };
  }

  static formatPaintingDateTime(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
  }

  // «Плановая готовность» берётся из того же единого маршрута, что и Малярка.
  static getPaintingPlannedReadinessDates(orders) {
    const tasks = this.getPaintingRouteSchedule(orders);
    if (!tasks.length) return {};

    return tasks.reduce((result, task) => {
      const finish = task.scheduleFinish instanceof Date
        ? task.scheduleFinish
        : CalendarService.parseStoredDate(String(task.finishDate || "").slice(0, 10));
      if (!finish) return result;
      const resultKey = task.orderKey || task.key;
      const previousTask = result[resultKey + "__date"];
      const previous = previousTask ? new Date(previousTask) : null;
      if (!previous || finish > previous) result[resultKey] = task.finishDate;
      if (!previous || finish > previous) result[resultKey + "__date"] = finish.getTime();
      return result;
    }, {});
  }

  static isPaintingQueueOrder(order) {
    return this.getPaintingOperations().some(operation =>
      operation.role === "painting" && operation.stage === String(order.stage || "").trim());
  }

  static applyProductionReserveDates(tasks) {
    tasks.forEach(task => {
      if (!task.reserveAfterTaskType) return;
      const primary = tasks.find(item => item.key === task.key && item.taskType === task.reserveAfterTaskType);
      if (!primary || !primary.startDate) return;
      const start = CalendarService.parseStoredDate(primary.startDate);
      if (!start) return;
      task.plannedDate = this.addWorkingDuration(start, task.reserveOffsetDays);
    });
  }

  // Возвращает начало следующего доступного слота после указанной трудоёмкости.
  // Это сохраняет точность в полдня: 0,5 дня шлифовки + 0,5 дня грунта
  // сдвигают резерв на следующий рабочий день.
  static addWorkingDuration(startDate, duration) {
    let current = new Date(startDate);
    let halfDay = 0;
    const halfDays = Math.max(1, Math.round((Number(duration) || 0.5) * 2));
    for (let index = 0; index < halfDays; index++) {
      halfDay += 1;
      if (halfDay === 2) {
        current = CalendarService.nextWorkDay(current);
        halfDay = 0;
      }
    }
    return CalendarService.formatDate(current);
  }





  /**
   * Универсальное обновление поля заказа
   */
  static updateField(
    key,
    fieldName,
    value
  ) {


    SheetService.updateField(

      key,

      fieldName,

      value

    );


    return true;


  }

  static getOrderFiles(key) {
    const order = this.getOrderByKey(key);
    return GoogleDriveService.getOrderFiles(order);
  }

  static uploadOrderFile(key, section, file) {
    const order = this.getOrderByKey(key);
    return GoogleDriveService.uploadOrderFile(order, section, file);
  }

  static downloadOrderFile(key, section, fileId) {
    const order = this.getOrderByKey(key);
    return GoogleDriveService.downloadOrderFile(order, section, fileId);
  }

  static downloadMeasurementArchive(key) {
    return GoogleDriveService.downloadMeasurementArchive(this.getOrderByKey(key));
  }

  static getOrderFileThumbnail(key, section, fileId) {
    const order = this.getOrderByKey(key);
    return GoogleDriveService.getOrderFileThumbnail(order, section, fileId);
  }

  static deleteOrderFile(key, section, fileId) {
    const order = this.getOrderByKey(key);
    return GoogleDriveService.deleteOrderFile(order, section, fileId);
  }

  static getOrderByKey(key) {
    const context = SheetService.getOrderContext(key);
    const order = SheetService.objectsFromRows(context.rows)
      .map(row => OrderMapper.map(row))
      .find(item => item.key === key);
    if (!order) throw new Error("Заказ не найден.");
    return order;
  }

  static toNotificationItem(order, fields) {
    const source = Object.assign({}, order || {}, fields || {});
    return {
      key: source.key || "", contract: source.contract || "", client: source.client || "",
      product: source.product || "", manager: source.manager || "", constructor: source.constructor || "",
      stage: source.stage || "", responsible: source.responsible || "", plannedDate: source.plannedDate || "",
      deadline: source.deadline || "", comment: source.comment || ""
    };
  }

  static notifyAssignment(oldOrder, newOrder, eventType) {
    return NotificationService.onAssignmentChanged(
      this.toNotificationItem(oldOrder), this.toNotificationItem(newOrder), { eventType: eventType }
    );
  }





  static createOrder(orderData) {
    const order = {
      contract: String(orderData.contract || "").trim(),
      client: String(orderData.client || "").trim(),
      product: String(orderData.product || "").trim(),
      manager: String(orderData.manager || "").trim(),
      constructor: "",
      duration: 0.5
    };

    const labels = { contract: "Договор", client: "Клиент", product: "Изделие", manager: "Менеджер" };
    Object.keys(labels).forEach(field => {
      if (!order[field]) throw new Error("Заполните поле: " + labels[field]);
    });
    // Быстрая проверка до создания папки заказа. В appendOrder она повторяется
    // под блокировкой, чтобы исключить дубль при одновременном добавлении.
    SheetService.assertOrderKeyUnique(order);
    order.constructor = this.getInitialConstructor(order.product);
    order.key = [order.contract, order.client, order.product].join("|");

    // Хранилище файлов — дополнительная функция. Его отсутствие или
    // незавершённая настройка не должны блокировать создание заказа.
    try {
      order.cloudFolderUrl = GoogleDriveService.createOrderFolder(order).url;
    } catch (error) {
      Logger.log("Папка Google Drive не создана для " + order.contract + ": " + error.message);
      order.cloudFolderUrl = "";
    }

    // Новую запись создаёт сам менеджер, поэтому на этапе «ТЗ» уведомление
    // не требуется. Последующие передачи уведомляются отдельными переходами.
    return SheetService.appendOrder(order);
  }

  static getInitialConstructor(product) {
    if (/кух/i.test(String(product || ""))) return "Валентина";

    const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS)
      .map(row => OrderMapper.map(row));
    const constructors = this.getRoleResponsible("Конструкторы", "Валентина", "Ирина");

    if (!constructors.length) {
      throw new Error("В справочнике не найден ни один конструктор.");
    }

    const workload = constructors.reduce((result, constructor) => {
      result[constructor] = 0;
      return result;
    }, {});
    orders.filter(order => this.isConstructorQueueOrder(order)).forEach(order => {
      if (Object.prototype.hasOwnProperty.call(workload, order.constructor)) {
        workload[order.constructor] += Number(order.duration) || 0;
      }
    });

    return constructors.sort((left, right) =>
      workload[left] - workload[right] || left.localeCompare(right, "ru")
    )[0];
  }

  static moveOrderToDate(key, targetDate) {
    const date = CalendarService.parseInputDate(targetDate);
    return this.changeOrderQueue(key, null, date);
  }

  static assertConstructorActionsAvailable(order) {
    if (String(order && order.stage || "").trim() === "Согласование") {
      throw new Error("Заказ ожидает решения менеджера. Действия конструктора временно недоступны.");
    }
  }

  static clearManualOrderDate(key) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      SheetService.ensureOrderColumn("Ручной приоритет даты");
      // У одного изделия в листе могут существовать параллельные строки для
      // разных этапов (например, «Запуск» и «Технолог»). Обычный find по
      // составному ключу тогда брал первую строку и не давал согласовать
      // фактическую строку технолога. Сохраняем номер строки и выбираем
      // именно вариант этапа «Технолог».
      const orderRows = SheetService.getData(CONFIG.SHEETS.ORDERS);
      const orders = SheetService.objectsFromRows(orderRows)
        .map((row, index) => {
          const mapped = OrderMapper.map(row);
          mapped.sheetRow = index + 2;
          return mapped;
        });
      const normalizeStage = value => String(value || "")
        .replace(/\u00a0/g, " ")
        .trim()
        .toLowerCase();
      const order = orders.find(item =>
        item.key === key && normalizeStage(item.stage) === "технолог"
      );
      if (!order || !this.isConstructorQueueOrder(order)) {
        throw new Error("Заказ недоступен в очереди конструктора.");
      }
      this.assertConstructorActionsAvailable(order);

      order.plannedDate = "";
      order.manualPriority = false;
      const queueOrders = orders.filter(item => this.isConstructorQueueOrder(item));
      CalendarService.buildCalendar(queueOrders);
      SheetService.saveQueueAndCalendar(
        queueOrders.filter(item => item.constructor === order.constructor)
      );
      SheetService.updateFields(key, {
        "Плановая дата": "",
        "Ручной приоритет даты": false
      });
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static assignOrderConstructor(key, constructor) {

    const newConstructor = String(constructor || "").trim();

    if (!newConstructor) {
      throw new Error("Выберите конструктора.");
    }

    return this.changeOrderQueue(key, newConstructor, null);

  }

  static changeOrderDuration(key, value) {
    const duration = Number(value);
    if (duration < 0.5 || Math.round(duration * 2) !== duration * 2) {
      throw new Error("Длительность должна быть кратна 0,5 дня и не меньше 0,5.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS)
        .map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);

      if (!order) throw new Error("Заказ не найден.");
      this.assertConstructorActionsAvailable(order);

      order.duration = duration;
      SheetService.updateField(key, "Длительность", duration);

      if (this.isConstructorQueueOrder(order)) {
        const queueOrders = orders.filter(item => this.isConstructorQueueOrder(item));
        CalendarService.buildCalendar(queueOrders);
        SheetService.saveQueueAndCalendar(
          queueOrders.filter(item => item.constructor === order.constructor)
        );
      }

      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static changeConstructorTaskDays(key, taskType, value) {
    const days = Number(value);
    if (days < 0.5 || Math.round(days * 2) !== days * 2) {
      throw new Error("Дни должны быть кратны 0,5 и не меньше 0,5.");
    }

    const order = this.getOrderByKey(key);
    this.assertConstructorActionsAvailable(order);
    const field = taskType === "kd" || taskType === "kd-reserve"
      ? "Дни КД"
      : "Дни Проформы";
    SheetService.ensureOrderColumn(field);
    SheetService.updateField(key, field, days);

    return { success: true };
  }

  static advanceManagerStage(key, targetStage) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      SheetService.ensureOrderColumn("Проформа Согласована");
      SheetService.ensureOrderColumn("Ручной приоритет даты");
      ["Факт. начало работы", "Факт. половина начала", "Остаток длительности"]
        .forEach(field => SheetService.ensureOrderColumn(field));
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS)
        .map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);

      if (!order) throw new Error("Заказ не найден.");
      const notificationBefore = this.toNotificationItem(order);

      const transitions = {
        "ТЗ": "Технолог",
        "Согласование": "КД",
        "Технолог": "Проформа"
      };

      if (transitions[order.stage] !== targetStage) {
        throw new Error("Этот переход этапа недоступен для заказа.");
      }

      if (targetStage !== "КД") {
        SheetService.updateField(key, "Этап", targetStage);
        if (targetStage === "Технолог") {
          this.notifyAssignment(notificationBefore, Object.assign({}, order, { stage: targetStage }), "TECHNOLOGIST_REVIEW");
        }
        return { success: true };
      }

      // Новый КД получает приоритет над «Проформами», но не обгоняет
      // уже запланированные КД того же конструктора.
      order.stage = "КД";
      order.status = "";
      order.plannedDate = "";
      order.manualPriority = false;
      order.actualStartDate = "";
      order.actualStartHalf = "";
      order.remainingDuration = 0;
      order.queue = order.kdReserveQueue || order.queue;
      const queueOrders = orders.filter(item => this.isConstructorQueueOrder(item));
      this.normalizeConstructorQueue(queueOrders, order.constructor, order);

      CalendarService.buildCalendar(queueOrders);

      SheetService.updateFields(key, {
        "Этап": "КД",
        "Статус": "",
        "Очередность": order.queue,
        "Плановая дата": "",
        "Ручной приоритет даты": false,
        "Факт. начало работы": "",
        "Факт. половина начала": "",
        "Остаток длительности": "",
        "Проформа Согласована": CalendarService.formatDate(new Date())
      });
      SheetService.saveQueueAndCalendar(
        queueOrders.filter(item => item.constructor === order.constructor)
      );
      this.notifyAssignment(notificationBefore, order, "KD_ASSIGNED");

      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static reviewByTechnologist(key, action, constructor, proformaDaysValue, kdDaysValue) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      SheetService.ensureOrderColumn("Дни ТЗ");
      SheetService.ensureOrderColumn("Дни Проформы");
      SheetService.ensureOrderColumn("Дни КД");
      SheetService.ensureOrderColumn("Резерв КД — очередь");
      SheetService.ensureOrderColumn("Плановая дата");
      SheetService.ensureOrderColumn("Ручной приоритет даты");
      ["Факт. начало работы", "Факт. половина начала", "Остаток длительности"]
        .forEach(field => SheetService.ensureOrderColumn(field));
      // У одного изделия могут быть параллельные строки на разных этапах.
      // Поэтому сохраняем номер строки и выбираем именно этап «Технолог».
      const orderRows = SheetService.getData(CONFIG.SHEETS.ORDERS);
      const orders = SheetService.objectsFromRows(orderRows).map((row, index) => {
        const mapped = OrderMapper.map(row);
        mapped.sheetRow = index + 2;
        return mapped;
      });
      const normalizeStage = value => String(value || "")
        .replace(/\u00a0/g, " ")
        .trim()
        .toLowerCase();
      const order = orders.find(item =>
        item.key === key && normalizeStage(item.stage) === "технолог"
      );

      if (!order) {
        throw new Error("Заказ недоступен для согласования технологом.");
      }
      const notificationBefore = this.toNotificationItem(order);

      if (action === "revision") {
        SheetService.updateFieldsAtRow(order.sheetRow, {
          "Этап": "ТЗ",
          "Статус": "",
          "Плановая дата": "",
          "Ручной приоритет даты": false,
          "Факт. начало работы": "",
          "Факт. половина начала": "",
          "Остаток длительности": ""
        });
        this.notifyAssignment(notificationBefore, Object.assign({}, order, { stage: "ТЗ", status: "" }), "TECHNOLOGIST_REVISION");
        return { success: true };
      }

      if (action !== "approved") {
        throw new Error("Неизвестное решение технолога.");
      }

      const selectedConstructor = String(constructor || "").trim();
      const proformaDays = Number(proformaDaysValue);
      const kdDays = Number(kdDaysValue);
      if (!selectedConstructor) throw new Error("Выберите конструктора.");
      if ([proformaDays, kdDays].some(value => value < 0.5 || Math.round(value * 2) !== value * 2)) {
        throw new Error("Дни Проформы и КД должны быть кратны 0,5 и не меньше 0,5.");
      }

      order.stage = "Проформа";
      order.status = "";
      order.constructor = selectedConstructor;
      order.duration = proformaDays;
      order.startDate = "";
      order.plannedDate = "";
      order.manualPriority = false;
      order.actualStartDate = "";
      order.actualStartHalf = "";
      order.remainingDuration = 0;

      const queueOrders = orders.filter(item => this.isConstructorQueueOrder(item));
      const peers = queueOrders.filter(item => item.constructor === selectedConstructor);
      const lastQueue = peers.reduce((max, item) => Math.max(max, Number(item.queue) || 0), 0);
      order.queue = lastQueue + 10;
      order.kdReserveQueue = lastQueue + 11;

      CalendarService.buildCalendar(queueOrders);
      SheetService.updateFieldsAtRow(order.sheetRow, {
        "Этап": order.stage,
        "Статус": order.status,
        "Конструктор": selectedConstructor,
        "Очередность": order.queue,
        "Начало": order.startDate,
        "Окончание": order.finishDate,
        "Длительность": proformaDays,
        "Дни Проформы": proformaDays,
        "Дни КД": kdDays,
        "Резерв КД — очередь": order.kdReserveQueue,
        "Плановая дата": "",
        "Ручной приоритет даты": false,
        "Факт. начало работы": "",
        "Факт. половина начала": "",
        "Остаток длительности": ""
      });
      // Текущая строка сохранена по её номеру выше. Остальные заказы очереди
      // можно обновлять обычным составным ключом.
      SheetService.saveQueueAndCalendar(
        queueOrders.filter(item => item.constructor === selectedConstructor && item !== order)
      );
      this.notifyAssignment(notificationBefore, order, "PROFORMA_ASSIGNED");

      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static changeOrderQueue(key, newConstructor, targetDate) {

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      const orders = SheetService
        .getObjects(CONFIG.SHEETS.ORDERS)
        .map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);

      if (!order) {
        throw new Error("Заказ не найден.");
      }
      const notificationBefore = this.toNotificationItem(order);

      this.assertConstructorActionsAvailable(order);

      const previousConstructor = order.constructor;
      const queueOrders = orders.filter(item => this.isConstructorQueueOrder(item));
      CalendarService.buildCalendar(queueOrders);
      const date = targetDate || new Date(order.scheduleStart);
      const nextConstructor = newConstructor || previousConstructor;

      // Ручная дата задаёт приоритет в календаре, но сохраняет место заказа
      // в числовой очереди. Переназначение конструктора работает отдельно.
      if (targetDate && !newConstructor) {
        SheetService.ensureOrderColumn("Ручной приоритет даты");
        order.plannedDate = CalendarService.formatDate(targetDate);
        order.manualPriority = true;
        CalendarService.buildCalendar(queueOrders);
        SheetService.saveQueueAndCalendar(
          queueOrders.filter(item => item.constructor === previousConstructor)
        );
        SheetService.updateFields(key, {
          "Плановая дата": order.plannedDate,
          "Ручной приоритет даты": true
        });
        return { success: true };
      }

      const targetQueue = queueOrders
        .filter(item => item !== order && item.constructor === nextConstructor)
        .sort((left, right) => left.queue - right.queue);
      const insertIndex = targetQueue.findIndex(item => date <= item.scheduleFinish);
      const position = insertIndex === -1 ? targetQueue.length : insertIndex;
      const previous = targetQueue[position - 1];
      const next = targetQueue[position];

      order.constructor = nextConstructor;
      order.queue = this.getQueueValue(previous, next);

      if (targetDate) {
        order.plannedDate = CalendarService.formatDate(targetDate);
      }
      CalendarService.buildCalendar(queueOrders);

      const affectedOrders = queueOrders.filter(item =>
        item.constructor === previousConstructor || item.constructor === nextConstructor
      );
      SheetService.saveQueueAndCalendar(affectedOrders);

      if (targetDate) {
        SheetService.updateField(key, "Плановая дата", order.plannedDate);
      }

      if (newConstructor && nextConstructor !== previousConstructor) {
        this.notifyAssignment(notificationBefore, order, "CONSTRUCTOR_REASSIGNED");
      }

      return { success: true };
    } finally {
      lock.releaseLock();
    }

  }

  static getQueueValue(previous, next) {

    if (!previous && !next) return 10;
    if (!previous) return next.queue - 1;
    if (!next) return previous.queue + 10;

    return (previous.queue + next.queue) / 2;

  }

  static normalizeConstructorQueue(queueOrders, constructor, newestOrder) {
    const peers = queueOrders
      .filter(item => item.constructor === constructor && item !== newestOrder)
      .sort((left, right) => left.queue - right.queue);
    const kdOrders = peers.filter(item => item.stage === "КД");
    const proformaOrders = peers.filter(item =>
      item.stage === "Проформа" || item.stage === "Согласование"
    );
    const newOrderGroup = newestOrder.stage === "КД" ? kdOrders : proformaOrders;

    // Новый КД ставится после других КД, а новая Проформа — после Проформ.
    newOrderGroup.push(newestOrder);
    kdOrders.concat(proformaOrders).forEach((item, index) => {
      item.queue = (index + 1) * 10;
    });
  }

  static isConstructorQueueOrder(order) {
    const stage = String(order.stage || "").trim().toLowerCase();
    const status = String(order.status || "").trim().toLowerCase();
    return (stage === "проформа" || stage === "согласование" || stage === "кд") && status !== "готово";
  }

  static getConstructorWorkDuration(order) {
    const stageDuration = String(order.stage || "").trim() === "КД"
      ? order.kdDays
      : order.proformaDays;
    return Number(order.remainingDuration) || Number(stageDuration) || Number(order.duration) || 0.5;
  }

  static calculateRemainingConstructorDuration(order, today, currentHalfDay) {
    const baseDuration = this.getConstructorWorkDuration(order);
    const actualStart = CalendarService.parseStoredDate(order.actualStartDate || order.startDate) || today;
    const completedHalfDays = CalendarService.countCompletedHalfDays(
      actualStart,
      order.actualStartHalf,
      today,
      currentHalfDay
    );
    return Math.max(0.5, Math.round((baseDuration - completedHalfDays * 0.5) * 2) / 2);
  }

  static changeWorkStatus(key, action, comment) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      ["Факт. начало работы", "Факт. половина начала", "Остаток длительности"]
        .forEach(field => SheetService.ensureOrderColumn(field));
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS)
        .map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);

      if (!order) {
        throw new Error("Заказ не найден. Обновите очередь и повторите действие.");
      }
      const notificationBefore = this.toNotificationItem(order);
      if (!this.isConstructorQueueOrder(order)) {
        const stage = String(order.stage || "не указан").trim() || "не указан";
        const status = String(order.status || "без статуса").trim() || "без статуса";
        throw new Error(
          "Заказ уже изменился в другой очереди: этап «" + stage +
          "», статус «" + status + "». Список будет обновлён."
        );
      }
      this.assertConstructorActionsAvailable(order);

      const queueOrders = orders.filter(item => this.isConstructorQueueOrder(item));
      CalendarService.buildCalendar(queueOrders);
      const today = CalendarService.firstWorkDay();
      const currentHalfDay = CalendarService.getCurrentHalfDay();

      if (action === "start") {
        order.status = "В работе";
        order.startDate = CalendarService.formatDate(today);
        order.actualStartDate = CalendarService.formatDate(today);
        order.actualStartHalf = currentHalfDay === 0 ? "Ⅰ" : "Ⅱ";
        order.plannedDate = "";
        order.manualPriority = false;
        this.placeOrderAtDate(order, queueOrders, today);
        if (order.stage === "Проформа") {
          const realKd = orders
            .filter(item => item.constructor === order.constructor && item.stage === "КД" && String(item.status).toLowerCase() !== "готово")
            .sort((left, right) => left.queue - right.queue);
          const lastKd = realKd[realKd.length - 1];
          order.kdReserveQueue = lastKd ? lastKd.queue + 0.1 : order.queue + 0.1;
        }
      } else if (action === "pause") {
        const pauseComment = String(comment || "").trim();
        if (!pauseComment) throw new Error("Укажите причину паузы.");
        order.remainingDuration = this.calculateRemainingConstructorDuration(
          order,
          today,
          currentHalfDay
        );
        order.status = "Пауза";
        const nextDay = CalendarService.nextWorkDay(today);
        order.startDate = CalendarService.formatDate(nextDay);
        order.plannedDate = CalendarService.formatDate(nextDay);
        order.manualPriority = false;
        this.placeOrderAtDate(order, queueOrders, nextDay, true);
        order.comment = pauseComment;
      } else if (action === "approval") {
        if (String(order.stage || "").trim() !== "Проформа") {
          throw new Error("На согласование можно передать только заказ на этапе «Проформа».");
        }
        // Менеджер сразу видит заказ на согласовании, а конструктор продолжает
        // работу над той же задачей до решения менеджера.
        order.stage = "Согласование";
        order.status = "В работе";
        order.startDate = CalendarService.formatDate(today);
        order.actualStartDate = CalendarService.formatDate(today);
        order.actualStartHalf = currentHalfDay === 0 ? "Ⅰ" : "Ⅱ";
        order.plannedDate = "";
        order.manualPriority = false;
        this.placeOrderAtDate(order, queueOrders, today, true);
      } else if (action === "stop") {
        order.status = "Готово";
        order.stage = this.getNextStage(order.stage);
        const enteredManufacturing = order.stage === "Очередь";
        const fields = {
          "Статус": enteredManufacturing ? "" : order.status,
          "Этап": order.stage,
          "Факт. начало работы": "",
          "Факт. половина начала": "",
          "Остаток длительности": ""
        };
        if (enteredManufacturing || order.stage === "В покраску") {
          const maxProductionQueue = orders.reduce((max, item) =>
            Math.max(max, Number(item.productionQueue) || 0), 0);
          fields["Очередь производства"] = maxProductionQueue + 10;
        }
        if (enteredManufacturing) fields["Ответственный"] = "Производство";
        SheetService.updateFields(order.key, fields);

        const remaining = queueOrders.filter(item => item !== order);
        CalendarService.buildCalendar(remaining);
        SheetService.saveQueueAndCalendar(remaining.filter(item => item.constructor === order.constructor));
        return { success: true };
      } else {
        throw new Error("Неизвестное действие со статусом.");
      }

      CalendarService.buildCalendar(queueOrders);
      SheetService.saveQueueAndCalendar(queueOrders.filter(item => item.constructor === order.constructor));
      const fields = {
        "Статус": order.status,
        "Этап": order.stage,
        "Плановая дата": order.plannedDate,
        "Ручной приоритет даты": order.manualPriority,
        "Факт. начало работы": order.actualStartDate || "",
        "Факт. половина начала": order.actualStartHalf || "",
        "Остаток длительности": order.remainingDuration || "",
        "Резерв КД — очередь": order.kdReserveQueue
      };
      if (action === "pause") fields["Комментарий"] = order.comment;
      SheetService.updateFields(order.key, fields);
      if (action === "approval") {
        this.notifyAssignment(notificationBefore, order, "PROFORMA_APPROVAL");
      }
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static changeProductionStatus(key, action) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS)
        .map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);
      const manufacturingStages = ["Запуск", "Распил", "Присадка", "Фрезеровка", "Сборка"];
      const allowedStages = ["В покраску"].concat(manufacturingStages);
      if (!order || !allowedStages.includes(order.stage)) {
        throw new Error("Заказ недоступен в выбранной производственной очереди.");
      }

      if (action === "start") {
        const active = orders.find(item => item.key !== key && item.stage === order.stage && String(item.status).toLowerCase() === "в работе");
        if (active) throw new Error(`На этапе «${order.stage}» уже есть заказ в работе.`);
        SheetService.updateField(key, "Статус", "В работе");
      } else if (action === "pause") {
        SheetService.updateField(key, "Статус", "Пауза");
      } else if (action === "stop" || action === "send") {
        if (order.stage !== "В покраску") {
          const currentIndex = manufacturingStages.indexOf(order.stage);
          const nextStage = currentIndex < manufacturingStages.length - 1
            ? manufacturingStages[currentIndex + 1]
            : (order.hasMdfFacade || order.hasVeneerFacade ? "В покраску" : "Упаковка");
          const fields = { "Статус": "", "Этап": nextStage, "Ответственный": "Производство" };
          if (String(order.stage || "").trim() === "Фрезеровка" && (order.hasMdfFacade || order.hasVeneerFacade)) {
            fields["Этап"] = "В покраску";
            if (order.hasMdfFacade) {
              fields["Этап Малярка МДФ"] = "Подготовка";
              fields["Статус Малярка МДФ"] = "";
              fields["Операция Малярка МДФ"] = "preparation";
            }
            if (order.hasVeneerFacade) {
              fields["Этап Малярка Шпон"] = "Подготовка";
              fields["Статус Малярка Шпон"] = "";
              fields["Операция Малярка Шпон"] = "preparation";
            }
            this.updateRouteFields(key, order, fields);
            return { success: true };
          }
          if (nextStage === "В покраску") {
            const firstOperation = this.getActivePaintingOperations(order)[0];
            if (order.hasMdfFacade) {
              fields["Этап Малярка МДФ"] = firstOperation ? firstOperation.stage : "Подготовка";
              fields["Статус Малярка МДФ"] = "";
              fields["Операция Малярка МДФ"] = firstOperation ? firstOperation.taskType : "preparation";
            }
            if (order.hasVeneerFacade) {
              fields["Этап Малярка Шпон"] = firstOperation ? firstOperation.stage : "Подготовка";
              fields["Статус Малярка Шпон"] = "";
              fields["Операция Малярка Шпон"] = firstOperation ? firstOperation.taskType : "preparation";
            }
            const maxProductionQueue = orders.reduce((max, item) =>
              Math.max(max, Number(item.productionQueue) || 0), 0);
            fields["Очередь производства"] = maxProductionQueue + 10;
          } else if (nextStage === "Упаковка") {
            fields["Этап ЛДСП"] = "Упаковка Корпуса";
            fields["Статус ЛДСП"] = "";
          }
          this.updateRouteFields(key, order, fields);
          return { success: true };
        }
        this.validateProductionDays(order);
        const firstPaintingOperation = this.getActivePaintingOperations(order)[0];
        if (!firstPaintingOperation) throw new Error("В маршруте Малярки нет активных операций.");
        const paintingAssignment = this.assignPainter(orders);
        const sandingAssignment = this.assignSandingWorker(orders);
        ["Маляр", "Очередь грунт", "Плановая дата покраски", "Шлифовщик", "Очередь шлифовки", "Плановая дата шлифовки", "Малярка — ручной приоритет"]
          .forEach(header => SheetService.ensureOrderColumn(header));
        this.updateRouteFields(key, order, {
          "Статус": "",
          "Этап": firstPaintingOperation.stage,
          "Операция Малярки": firstPaintingOperation.taskType,
          "Маляр": paintingAssignment.painter,
          "Очередь грунт": paintingAssignment.primerQueue,
          "Плановая дата покраски": "",
          "Шлифовщик": sandingAssignment.worker,
          "Очередь шлифовки": sandingAssignment.primaryQueue,
          "Плановая дата шлифовки": "",
          "Малярка — ручной приоритет": false
        });
      } else {
        throw new Error("Неизвестное действие буфера Малярки.");
      }
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static launchManufacturing(key) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      if (String(order.stage || "").trim() !== "Запуск") {
        throw new Error("Заказ уже покинул очередь запуска.");
      }
      if (!order.hasLdspFacade && !order.hasMdfFacade && !order.hasVeneerFacade && !order.hasNoFacade) {
        throw new Error("Укажите материал фасада или выберите «Без фасада».");
      }
      ["Этап ЛДСП", "Статус ЛДСП", "Корпус МДФ", "Этап Корпус МДФ", "Статус Корпус МДФ",
        "Этап Малярка Корпус МДФ", "Статус Малярка Корпус МДФ", "Операция Малярка Корпус МДФ",
        "Крашеный ЛДСП", "Этап Крашеный ЛДСП", "Статус Крашеный ЛДСП",
        "Этап Малярка Крашеный ЛДСП", "Статус Малярка Крашеный ЛДСП", "Операция Малярка Крашеный ЛДСП",
        "Этап МДФ", "Статус МДФ", "Фасад МДФ",
        "Этап Шпон", "Статус Шпон", "Фасад Шпон"]
        .forEach(field => SheetService.ensureOrderColumn(field));
      this.updateRouteFields(key, order, {
        "Этап": "Распил",
        "Статус": "",
        "Ответственный": "Производство",
        "Этап ЛДСП": order.hasCorpusLdsp === false ? "" : "Распил ЛДСП",
        "Статус ЛДСП": "",
        "Этап Корпус МДФ": order.hasCorpusMdf ? "Распил МДФ" : "",
        "Статус Корпус МДФ": "",
        "Этап Малярка Корпус МДФ": "",
        "Статус Малярка Корпус МДФ": "",
        "Операция Малярка Корпус МДФ": "",
        "Этап Крашеный ЛДСП": order.hasPaintedLdsp ? "Распил ЛДСП" : "",
        "Статус Крашеный ЛДСП": "",
        "Этап Малярка Крашеный ЛДСП": "",
        "Статус Малярка Крашеный ЛДСП": "",
        "Операция Малярка Крашеный ЛДСП": "",
        "Этап МДФ": order.hasMdfFacade ? "Распил МДФ" : "",
        "Статус МДФ": "",
        "Этап Шпон": order.hasVeneerFacade ? "Распил Шпон" : "",
        "Статус Шпон": "",
        "Этап Малярка МДФ": "",
        "Статус Малярка МДФ": "",
        "Операция Малярка МДФ": "",
        "Этап Малярка Шпон": "",
        "Статус Малярка Шпон": "",
        "Операция Малярка Шпон": "",
        "Этап упаковки фасада МДФ": "",
        "Статус упаковки фасада МДФ": "",
        "Этап упаковки фасада Шпон": "",
        "Статус упаковки фасада Шпон": ""
      });
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  // Состав корпусов фиксируется до запуска. В запущенном изделии допускается
  // только добавление новой независимой ветки — исключение изменило бы историю
  // уже начатого маршрута.
  static setCorpusTypes(key, selection) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      const requested = selection || {};
      const values = {
        hasCorpusLdsp: Boolean(requested.hasCorpusLdsp),
        hasCorpusMdf: Boolean(requested.hasCorpusMdf),
        hasPaintedLdsp: Boolean(requested.hasPaintedLdsp)
      };
      const fields = {
        "Корпус ЛДСП": values.hasCorpusLdsp,
        "Корпус МДФ": values.hasCorpusMdf,
        "Крашеный ЛДСП": values.hasPaintedLdsp
      };
      ["Корпус ЛДСП", "Корпус МДФ", "Крашеный ЛДСП",
        "Этап ЛДСП", "Статус ЛДСП", "Этап Корпус МДФ", "Статус Корпус МДФ",
        "Этап Малярка Корпус МДФ", "Статус Малярка Корпус МДФ", "Операция Малярка Корпус МДФ",
        "Этап Крашеный ЛДСП", "Статус Крашеный ЛДСП",
        "Этап Малярка Крашеный ЛДСП", "Статус Малярка Крашеный ЛДСП", "Операция Малярка Крашеный ЛДСП"]
        .forEach(field => SheetService.ensureOrderColumn(field));

      const isLaunched = !["Очередь", "Запуск"].includes(String(order.stage || "").trim());
      const current = {
        hasCorpusLdsp: order.hasCorpusLdsp !== false,
        hasCorpusMdf: Boolean(order.hasCorpusMdf),
        hasPaintedLdsp: Boolean(order.hasPaintedLdsp)
      };
      if (isLaunched) {
        Object.keys(values).forEach(name => {
          if (current[name] && !values[name]) {
            throw new Error("После запуска ветку корпуса нельзя отменить.");
          }
        });
        if (!current.hasCorpusLdsp && values.hasCorpusLdsp && !String(order.ldspStage || "").trim()) {
          fields["Этап ЛДСП"] = "Распил ЛДСП";
          fields["Статус ЛДСП"] = "";
        }
        ["corpus-mdf", "painted-ldsp"].forEach(branch => {
          const config = this.getAdditionalCorpusBranchConfig(branch);
          if (!current[config.selectedProperty] && values[config.selectedProperty]) {
            fields[config.stageField] = config.initialStage;
            fields[config.statusField] = "";
            fields[config.paintingStageField] = "";
            fields[config.paintingStatusField] = "";
            fields[config.paintingOperationField] = "";
          }
        });
      }
      this.updateRouteFields(key, order, fields);
      return Object.assign({ success: true }, values);
    } finally {
      lock.releaseLock();
    }
  }

  static changeManufacturingStatus(key, branch, action, comment) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      let normalizedBranch = String(branch || "").trim().toLowerCase();
      const order = this.getOrderByKey(key);
      if (!normalizedBranch && ["Распил ЛДСП", "Кромка", "Присадка", "Сборка",
          "Упаковка Корпуса", "СГП Корпуса", "Монтаж"].includes(String(order.stage || "").trim())) {
        SheetService.ensureOrderColumn("Этап ЛДСП");
        SheetService.ensureOrderColumn("Статус ЛДСП");
        SheetService.updateFields(key, {
          "Этап ЛДСП": order.stage,
          "Статус ЛДСП": order.status || ""
        });
        order.ldspStage = order.stage;
        order.ldspStatus = order.status || "";
        normalizedBranch = "ldsp";
      }
      const branchConfig = {
        ldsp: { stageField: "Этап ЛДСП", statusField: "Статус ЛДСП" },
        "corpus-mdf": { stageField: "Этап Корпус МДФ", statusField: "Статус Корпус МДФ", stageProperty: "corpusMdfStage", statusProperty: "corpusMdfStatus" },
        "painted-ldsp": { stageField: "Этап Крашеный ЛДСП", statusField: "Статус Крашеный ЛДСП", stageProperty: "paintedLdspStage", statusProperty: "paintedLdspStatus" },
        mdf: { stageField: "Этап МДФ", statusField: "Статус МДФ" },
        veneer: { stageField: "Этап Шпон", statusField: "Статус Шпон", stageProperty: "veneerStage", statusProperty: "veneerStatus" },
        "preparation-mdf": { stageField: "Этап Малярка МДФ", statusField: "Статус Малярка МДФ", stageProperty: "mdfPaintingStage", statusProperty: "mdfPaintingStatus", operationField: "Операция Малярка МДФ" },
        "preparation-corpus-mdf": { stageField: "Этап Малярка Корпус МДФ", statusField: "Статус Малярка Корпус МДФ", stageProperty: "corpusMdfPaintingStage", statusProperty: "corpusMdfPaintingStatus", operationField: "Операция Малярка Корпус МДФ" },
        "preparation-veneer": { stageField: "Этап Малярка Шпон", statusField: "Статус Малярка Шпон", stageProperty: "veneerPaintingStage", statusProperty: "veneerPaintingStatus", operationField: "Операция Малярка Шпон" },
        "preparation-legacy": { stageField: "Этап", statusField: "Статус", stageProperty: "stage", statusProperty: "status", operationField: "Операция Малярки" },
        "facade-packaging-mdf": { stageField: "Этап упаковки фасада МДФ", statusField: "Статус упаковки фасада МДФ", stageProperty: "mdfFacadePackagingStage", statusProperty: "mdfFacadePackagingStatus" },
        "facade-packaging-veneer": { stageField: "Этап упаковки фасада Шпон", statusField: "Статус упаковки фасада Шпон", stageProperty: "veneerFacadePackagingStage", statusProperty: "veneerFacadePackagingStatus" },
        // Историческое имя остаётся рабочим для незавершённых старых строк.
        "facade-packaging": { stageField: "Этап упаковки фасада", statusField: "Статус упаковки фасада", stageProperty: "facadePackagingStage", statusProperty: "facadePackagingStatus" },
        "legacy-installation": { stageField: "Этап", statusField: "Статус", stageProperty: "stage", statusProperty: "status" }
      }[normalizedBranch];
      if (!branchConfig) throw new Error("Неизвестная операция производства.");

      if (!branchConfig.stageProperty) {
        branchConfig.stageProperty = normalizedBranch === "ldsp" ? "ldspStage" : "mdfStage";
        branchConfig.statusProperty = normalizedBranch === "ldsp" ? "ldspStatus" : "mdfStatus";
      }
      const currentStage = order[branchConfig.stageProperty];
      const currentStatus = order[branchConfig.statusProperty];
      if (!String(currentStage || "").trim()) throw new Error("Операция не найдена.");
      if (normalizedBranch.startsWith("preparation-") && String(currentStage || "").trim() !== "Подготовка") {
        throw new Error("Подготовка доступна только для фасадной ветки на этапе «Подготовка».");
      }

      let nextStatus;
      if (action === "start") {
        if (String(currentStatus || "").trim().toLowerCase() === "в работе") return { success: true };
        nextStatus = "В работе";
      } else if (action === "pause") {
        if (String(currentStage || "").trim() === "Монтаж") {
          const pauseComment = String(comment || "").trim();
          if (!pauseComment) throw new Error("Укажите причину паузы.");
          const commentField = this.getInstallationCommentField(normalizedBranch);
          if (!commentField) throw new Error("Для этой ветки монтажа не найдено поле комментария.");
          SheetService.updateFields(key, {
            [branchConfig.statusField]: "Пауза",
            [commentField]: pauseComment
          });
          return { success: true };
        }
        nextStatus = "Пауза";
      } else if (action === "stop") {
        const transitions = {
          "Распил ЛДСП": "Кромка",
          "Распил МДФ": "Фрезеровка",
          "Кромка": "Присадка",
          "Присадка": "Сборка",
          "Фрезеровка": "Подготовка",
          "Распил Шпон": "Кромка Шпон",
          "Кромка Шпон": "Присадка Шпон",
          "Присадка Шпон": "Подготовка",
          "Подготовка": "В покраску",
          "Сборка": "Упаковка Корпуса",
          "Упаковка Корпуса": "СГП Корпуса",
          "Упаковка Фасада": "СГП Фасада",
          "СГП Корпуса": "Монтаж",
          "СГП Фасада": "Монтаж",
          "Монтаж": "Завершен"
        };
        let nextStage = transitions[String(currentStage || "").trim()];
        const fields = {
          [branchConfig.stageField]: nextStage || currentStage,
          [branchConfig.statusField]: nextStage ? "" : "Готово"
        };
        if (nextStage === "Подготовка" && ["mdf", "veneer", "corpus-mdf"].includes(normalizedBranch)) {
          const prefix = normalizedBranch === "veneer" ? "Шпон" : normalizedBranch === "corpus-mdf" ? "Корпус МДФ" : "МДФ";
          fields["Этап Малярка " + prefix] = "Подготовка";
          fields["Статус Малярка " + prefix] = "";
          fields["Операция Малярка " + prefix] = "preparation";
          // Производственная ветка передаётся в отдельный участок подготовки.
          fields[branchConfig.stageField] = "";
          fields[branchConfig.statusField] = "";
          fields["Этап"] = "В покраску";
          fields["Статус"] = "";
          fields["Ответственный"] = "Производство";
        }
        if (nextStage === "В покраску" && normalizedBranch.startsWith("preparation-")) {
          const firstOperation = this.getActivePaintingOperations(order)[0];
          if (!firstOperation) throw new Error("В маршруте Малярки нет активных операций.");
          fields[branchConfig.stageField] = "В покраску";
          fields[branchConfig.statusField] = "";
          if (branchConfig.operationField) fields[branchConfig.operationField] = firstOperation.taskType;
          // Из Буфера запуск остаётся ручной: первая операция записывается
          // как следующая, но её этап ещё не запускается.
          fields["Этап"] = "В покраску";
          fields["Статус"] = "";
          fields["Ответственный"] = "Производство";
        }
        if (normalizedBranch === "painted-ldsp" && String(currentStage || "").trim() === "Присадка") {
          const firstOperation = this.getActivePaintingOperations(order)[0];
          if (!firstOperation) throw new Error("В маршруте Малярки нет активных операций.");
          fields[branchConfig.stageField] = "";
          fields[branchConfig.statusField] = "";
          fields["Этап Малярка Крашеный ЛДСП"] = "В покраску";
          fields["Статус Малярка Крашеный ЛДСП"] = "";
          fields["Операция Малярка Крашеный ЛДСП"] = firstOperation.taskType;
        }
        if (["СГП Корпуса", "СГП Фасада"].includes(nextStage)) {
          fields["Этап"] = "СГП";
          fields["Статус"] = "";
          fields["Ответственный"] = "Производство";
        }
        if (nextStage === "Монтаж") {
          fields["Этап"] = "Монтаж";
          fields["Статус"] = "";
          fields["Ответственный"] = "Производство";
        }
        if (nextStage === "Завершен") {
          fields[branchConfig.statusField] = "Готово";
          if (!String(order.ldspStage || "").trim()) Object.assign(fields, this.getLegacyCompletionFields(order, normalizedBranch));
        }
        this.updateRouteFields(key, order, fields);
        return { success: true };
      } else {
        throw new Error("Неизвестное действие производства.");
      }
      SheetService.updateField(key, branchConfig.statusField, nextStatus);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  // СГП не является рабочей операцией: ветка адресно передаётся в монтаж
  // без промежуточного статуса. Общий этап пересчитывается updateRouteFields.
  static sendSgpToInstallation(key, branch) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const normalizedBranch = String(branch || "").trim().toLowerCase();
      const allowedBranches = new Set([
        "ldsp", "facade-packaging", "facade-packaging-mdf", "facade-packaging-veneer",
        "mdf", "veneer", "corpus-mdf", "painted-ldsp"
      ]);
      if (!allowedBranches.has(normalizedBranch)) {
        throw new Error("На монтаж можно передать только корпусную или фасадную ветку СГП.");
      }

      const order = this.getOrderByKey(key);
      const stream = this.getInstallationStreams(order)
        .find(item => item.taskType === normalizedBranch);
      const branchConfig = this.getInstallationBranchConfig(normalizedBranch);
      const expectedStage = ["ldsp", "corpus-mdf", "painted-ldsp"].includes(normalizedBranch)
        ? "СГП Корпуса" : "СГП Фасада";
      if (!branchConfig || !stream || String(stream.stage || "").trim() !== expectedStage) {
        throw new Error("Передать на монтаж можно только с этапа «" + expectedStage + "».");
      }

      this.updateRouteFields(key, order, {
        [branchConfig.stageField]: "Монтаж",
        [branchConfig.statusField]: "",
        ...this.getRouteStatusResetFields(normalizedBranch)
      });
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static saveInstallationComment(key, branch, comment) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const normalizedBranch = String(branch || "").trim().toLowerCase();
      const note = String(comment || "").trim();
      if (!note) throw new Error("Укажите комментарий монтажа.");
      const fieldName = this.getInstallationCommentField(normalizedBranch);
      if (!fieldName) throw new Error("Неизвестная ветка монтажа.");

      const order = this.getOrderByKey(key);
      const stream = this.getInstallationStreams(order)
        .find(item => item.taskType === normalizedBranch);
      if (!stream || !["Монтаж", "Завершен"].includes(String(stream.stage || "").trim())) {
        throw new Error("Комментарий можно изменить только у существующего потока монтажа.");
      }
      SheetService.ensureOrderColumn(fieldName);
      SheetService.updateField(key, fieldName, note);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static saveClaimComment(key, comment) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const note = String(comment || "").trim();
      if (!note) throw new Error("Укажите суть рекламации.");
      const order = this.getOrderByKey(key);
      if (String(order.stage || "").trim() !== "Рекламация") {
        throw new Error("Комментарий можно изменить только у изделия в рекламации.");
      }
      SheetService.ensureOrderColumn("Комментарий рекламации");
      SheetService.updateField(key, "Комментарий рекламации", note);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static changeManufacturingStage(key, branch, targetStage) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const allowedStages = ["Запуск", "Распил", "Распил ЛДСП", "Распил МДФ", "Распил Шпон",
        "Кромка", "Кромка Шпон", "Присадка", "Присадка Шпон", "Фрезеровка", "Подготовка", "Сборка", "В покраску",
        "Упаковка Корпуса", "Упаковка Фасада", "СГП Корпуса", "СГП Фасада", "Монтаж", "Завершен"];
      const stage = String(targetStage || "").trim();
      if (!allowedStages.includes(stage)) throw new Error("Выберите производственный этап из списка.");

      const order = this.getOrderByKey(key);
      const normalizedBranch = String(branch || "").trim().toLowerCase();
      if (String(order.stage || "").trim() === "Очередь" && (normalizedBranch || stage !== "Запуск")) {
        throw new Error("Сначала переведите изделие из очереди на этап «Запуск».");
      }
      if (String(order.stage || "").trim() === "Запуск" && stage !== "Запуск" &&
          !order.hasLdspFacade && !order.hasMdfFacade && !order.hasVeneerFacade && !order.hasNoFacade) {
        throw new Error("Укажите материал фасада или выберите «Без фасада».");
      }
      const hasBranches = this.hasManufacturingTasks(order);
      const baseStages = ["Очередь", "Запуск", "Распил", "Кромка", "Присадка", "Фрезеровка", "Подготовка", "Сборка"];
      if (!hasBranches && normalizedBranch !== "legacy-installation" && !baseStages.includes(String(order.stage || "").trim())) {
        throw new Error("Заказ уже не находится в разделе Производство.");
      }
      if (normalizedBranch && !["ldsp", "corpus-mdf", "painted-ldsp", "mdf", "veneer", "preparation-mdf", "preparation-veneer", "preparation-corpus-mdf", "preparation-legacy", "facade-packaging", "facade-packaging-mdf", "facade-packaging-veneer", "legacy-installation"].includes(normalizedBranch)) {
        throw new Error("Неизвестная производственная операция.");
      }
      if (stage === "Подготовка" && !["mdf", "veneer", "corpus-mdf", "preparation-mdf", "preparation-veneer", "preparation-corpus-mdf", "preparation-legacy"].includes(normalizedBranch)) {
        throw new Error("В «Подготовку» можно перевести только ветку МДФ или шпона.");
      }
      if (stage === "В покраску" && !["preparation-mdf", "preparation-veneer", "preparation-corpus-mdf", "painted-ldsp", "preparation-legacy"].includes(normalizedBranch)) {
        throw new Error("В Буфер Малярки можно передать только завершённую подготовку или крашеный ЛДСП после присадки.");
      }
      const branchStages = {
        ldsp: order.ldspStage,
        "corpus-mdf": order.corpusMdfStage,
        "painted-ldsp": order.paintedLdspStage,
        mdf: order.mdfStage,
        veneer: order.veneerStage,
        "preparation-mdf": order.mdfPaintingStage,
        "preparation-veneer": order.veneerPaintingStage,
        "preparation-corpus-mdf": order.corpusMdfPaintingStage,
        "preparation-legacy": order.stage,
        "facade-packaging": order.facadePackagingStage,
        "facade-packaging-mdf": order.mdfFacadePackagingStage,
        "facade-packaging-veneer": order.veneerFacadePackagingStage,
        "legacy-installation": order.stage
      };
      const branchStatuses = {
        ldsp: order.ldspStatus,
        "corpus-mdf": order.corpusMdfStatus,
        "painted-ldsp": order.paintedLdspStatus,
        mdf: order.mdfStatus,
        veneer: order.veneerStatus,
        "preparation-mdf": order.mdfPaintingStatus,
        "preparation-veneer": order.veneerPaintingStatus,
        "preparation-corpus-mdf": order.corpusMdfPaintingStatus,
        "preparation-legacy": order.status,
        "facade-packaging": order.facadePackagingStatus,
        "facade-packaging-mdf": order.mdfFacadePackagingStatus,
        "facade-packaging-veneer": order.veneerFacadePackagingStatus,
        "legacy-installation": order.status
      };
      if (stage === "Завершен") {
        // Проверка должна использовать тот же эффективный поток, который
        // отображается в «Монтажах». У переходных МДФ/шпоновых записей
        // статус мог сохраниться в совместимом поле, пока этап уже находится
        // в новой отдельной ветке упаковки.
        const stream = this.getInstallationStreams(order)
          .find(item => item.taskType === normalizedBranch);
        const currentBranchStage = String(stream ? stream.stage : (branchStages[normalizedBranch] || order.stage || "")).trim();
        const currentBranchStatus = String(stream ? stream.status : (branchStatuses[normalizedBranch] || order.status || "")).trim().toLowerCase();
        if (currentBranchStage !== "Монтаж" || currentBranchStatus !== "готово") {
          throw new Error("Вручную завершить можно только монтаж со статусом «Готово».");
        }
      }
      if (normalizedBranch) {
        const currentBranchStage = branchStages[normalizedBranch];
        if (!String(currentBranchStage || "").trim()) throw new Error("Операция уже покинула эту очередь.");
      }

      ["Этап ЛДСП", "Статус ЛДСП", "Этап Корпус МДФ", "Статус Корпус МДФ",
        "Этап Малярка Корпус МДФ", "Статус Малярка Корпус МДФ", "Операция Малярка Корпус МДФ",
        "Этап Крашеный ЛДСП", "Статус Крашеный ЛДСП",
        "Этап Малярка Крашеный ЛДСП", "Статус Малярка Крашеный ЛДСП", "Операция Малярка Крашеный ЛДСП",
        "Этап МДФ", "Статус МДФ", "Этап упаковки фасада", "Статус упаковки фасада"]
        .forEach(field => SheetService.ensureOrderColumn(field));

      if (stage === "Запуск") {
        this.updateRouteFields(key, order, {
          "Этап": stage,
          "Статус": "",
          "Ответственный": "Производство",
          "Этап ЛДСП": "",
          "Статус ЛДСП": "",
          "Этап Корпус МДФ": "", "Статус Корпус МДФ": "",
          "Этап Малярка Корпус МДФ": "", "Статус Малярка Корпус МДФ": "", "Операция Малярка Корпус МДФ": "",
          "Этап Крашеный ЛДСП": "", "Статус Крашеный ЛДСП": "",
          "Этап Малярка Крашеный ЛДСП": "", "Статус Малярка Крашеный ЛДСП": "", "Операция Малярка Крашеный ЛДСП": "",
          "Этап МДФ": "",
          "Статус МДФ": "",
          "Этап Шпон": "",
          "Статус Шпон": "",
          "Этап Малярка МДФ": "", "Статус Малярка МДФ": "", "Операция Малярка МДФ": "",
          "Этап Малярка Шпон": "", "Статус Малярка Шпон": "", "Операция Малярка Шпон": "",
          "Этап упаковки фасада МДФ": "", "Статус упаковки фасада МДФ": "",
          "Этап упаковки фасада Шпон": "", "Статус упаковки фасада Шпон": "",
          "Этап упаковки фасада": "", "Статус упаковки фасада": ""
        });
        return { success: true };
      }

      if (stage === "Распил") {
        this.updateRouteFields(key, order, {
          "Этап": "Распил",
          "Статус": "",
          "Ответственный": "Производство",
          "Этап ЛДСП": order.hasCorpusLdsp === false ? "" : "Распил ЛДСП",
          "Статус ЛДСП": "",
          "Этап Корпус МДФ": order.hasCorpusMdf ? "Распил МДФ" : "", "Статус Корпус МДФ": "",
          "Этап Малярка Корпус МДФ": "", "Статус Малярка Корпус МДФ": "", "Операция Малярка Корпус МДФ": "",
          "Этап Крашеный ЛДСП": order.hasPaintedLdsp ? "Распил ЛДСП" : "", "Статус Крашеный ЛДСП": "",
          "Этап Малярка Крашеный ЛДСП": "", "Статус Малярка Крашеный ЛДСП": "", "Операция Малярка Крашеный ЛДСП": "",
          "Этап МДФ": order.hasMdfFacade ? "Распил МДФ" : "",
          "Статус МДФ": "",
          "Этап Шпон": order.hasVeneerFacade ? "Распил Шпон" : "",
          "Статус Шпон": "",
          "Этап Малярка МДФ": "", "Статус Малярка МДФ": "", "Операция Малярка МДФ": "",
          "Этап Малярка Шпон": "", "Статус Малярка Шпон": "", "Операция Малярка Шпон": "",
          "Этап упаковки фасада МДФ": "", "Статус упаковки фасада МДФ": "",
          "Этап упаковки фасада Шпон": "", "Статус упаковки фасада Шпон": "",
          "Этап упаковки фасада": "", "Статус упаковки фасада": ""
        });
        return { success: true };
      }

      const targetBranch = normalizedBranch ||
        (["Упаковка Фасада", "СГП Фасада"].includes(stage) ? "facade-packaging-mdf" :
          ["Распил Шпон", "Кромка Шпон", "Присадка Шпон"].includes(stage) ? "veneer" :
          ["Распил МДФ", "Фрезеровка"].includes(stage) ? "mdf" : "ldsp");
      const targetConfig = {
        ldsp: { stage: "Этап ЛДСП", status: "Статус ЛДСП" },
        "corpus-mdf": { stage: "Этап Корпус МДФ", status: "Статус Корпус МДФ" },
        "painted-ldsp": { stage: "Этап Крашеный ЛДСП", status: "Статус Крашеный ЛДСП" },
        mdf: { stage: "Этап МДФ", status: "Статус МДФ" },
        veneer: { stage: "Этап Шпон", status: "Статус Шпон" },
        "preparation-mdf": { stage: "Этап Малярка МДФ", status: "Статус Малярка МДФ", operation: "Операция Малярка МДФ" },
        "preparation-veneer": { stage: "Этап Малярка Шпон", status: "Статус Малярка Шпон", operation: "Операция Малярка Шпон" },
        "preparation-corpus-mdf": { stage: "Этап Малярка Корпус МДФ", status: "Статус Малярка Корпус МДФ", operation: "Операция Малярка Корпус МДФ" },
        "preparation-legacy": { stage: "Этап", status: "Статус", operation: "Операция Малярки" },
        "facade-packaging": { stage: "Этап упаковки фасада", status: "Статус упаковки фасада" },
        "facade-packaging-mdf": { stage: "Этап упаковки фасада МДФ", status: "Статус упаковки фасада МДФ" },
        "facade-packaging-veneer": { stage: "Этап упаковки фасада Шпон", status: "Статус упаковки фасада Шпон" },
        "legacy-installation": { stage: "Этап", status: "Статус" }
      }[targetBranch];
      const stageField = targetConfig.stage;
      const statusField = targetConfig.status;
      const fields = {
        "Этап": "Распил",
        "Статус": "",
        "Ответственный": "Производство",
        [stageField]: stage,
        [statusField]: ""
      };
      if (stage !== "Завершен") Object.assign(fields, this.getRouteStatusResetFields(targetBranch));
      if (targetBranch === "mdf") fields["Фасад МДФ"] = true;
      if (targetBranch === "veneer") fields["Фасад Шпон"] = true;
      if (stage === "Подготовка" && ["mdf", "veneer", "corpus-mdf"].includes(targetBranch)) {
        const prefix = targetBranch === "veneer" ? "Шпон" : targetBranch === "corpus-mdf" ? "Корпус МДФ" : "МДФ";
        fields["Этап Малярка " + prefix] = "Подготовка";
        fields["Статус Малярка " + prefix] = "";
        fields["Операция Малярка " + prefix] = "preparation";
        fields[stageField] = "";
        fields["Этап"] = "В покраску";
      }
      if (stage === "В покраску") {
        const firstOperation = this.getActivePaintingOperations(order)[0];
        fields[stageField] = "В покраску";
        fields[statusField] = "";
        if (targetConfig.operation) fields[targetConfig.operation] = firstOperation ? firstOperation.taskType : "";
        fields["Этап"] = "В покраску";
      }
      if (stage === "В покраску" && targetBranch === "painted-ldsp") {
        fields["Этап Малярка Крашеный ЛДСП"] = "В покраску";
        fields["Статус Малярка Крашеный ЛДСП"] = "";
        fields["Операция Малярка Крашеный ЛДСП"] = this.getActivePaintingOperations(order)[0]?.taskType || "";
        fields["Этап Крашеный ЛДСП"] = "";
        fields["Статус Крашеный ЛДСП"] = "";
      }
      if (stage === "Упаковка Фасада") fields["Этап"] = "Упаковка";
      if (["СГП Корпуса", "СГП Фасада"].includes(stage)) fields["Этап"] = "СГП";
      if (stage === "Завершен") {
        fields[statusField] = "Готово";
        if (!String(order.ldspStage || "").trim()) Object.assign(fields, this.getLegacyCompletionFields(order, targetBranch));
      }
      this.updateRouteFields(key, order, fields);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static moveOrderToOverallStage(key, targetStage) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      const currentStage = String(order.stage || "").trim();
      if (!this.getOverallManufacturingStages().includes(currentStage)) {
        throw new Error("Ручной перевод доступен только для изделия в производственном маршруте.");
      }
      const fields = this.buildOverallStageMoveFields(order, targetStage);
      SheetService.updateFields(key, fields);
      return { success: true, previousStage: currentStage, stage: String(targetStage || "").trim() };
    } finally {
      lock.releaseLock();
    }
  }

  static sendOrderToClaims(key, branch, comment) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      const normalizedBranch = String(branch || "").trim().toLowerCase();
      const stream = this.getInstallationStreams(order).find(item => item.taskType === normalizedBranch);
      const config = this.getInstallationBranchConfig(normalizedBranch);
      if (!config || !stream || !["Монтаж", "Завершен"].includes(String(stream.stage || "").trim())) {
        throw new Error("В рекламацию можно передать только существующую монтажную ветку.");
      }
      if (String(order.stage || "").trim() === "Рекламация" &&
        String(stream.stage || "").trim() === "Завершен" &&
        String(stream.status || "").trim().toLowerCase() === "готово") {
        return { success: true };
      }
      const installationComment = this.getInstallationComment(order, normalizedBranch);
      const transferComment = String(comment || "").trim();
      const claimCommentLines = String(order.claimComment || "").split(/\r?\n/)
        .map(value => value.trim())
        .filter(Boolean);
      if (installationComment && !claimCommentLines.includes(installationComment)) {
        claimCommentLines.push(installationComment);
      }
      if (transferComment && !claimCommentLines.includes(transferComment)) claimCommentLines.push(transferComment);
      if (claimCommentLines.length) SheetService.ensureOrderColumn("Комментарий рекламации");
      const fields = {
        [config.stageField]: "Завершен",
        [config.statusField]: "Готово",
        "Этап": "Рекламация",
        "Статус": "",
        "Ответственный": "Производство"
      };
      if (claimCommentLines.length) fields["Комментарий рекламации"] = claimCommentLines.join("\n");
      SheetService.updateFields(key, fields);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static completeClaim(key) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      if (String(order.stage || "").trim() !== "Рекламация") {
        throw new Error("Завершить можно только изделие на этапе «Рекламация».");
      }
      const overview = this.getClaimOverview(order);
      if (!overview.claimReady) {
        throw new Error("Завершить рекламацию можно только после перевода всех обязательных веток в «Рекламация» или «Завершен · Готово».");
      }
      const fields = { "Этап": "Завершен", "Статус": "Готово", "Ответственный": "Производство" };
      if (order.manualClaim) fields["Ручная рекламация"] = false;
      this.getInstallationStreams(order).forEach(stream => {
        const config = this.getInstallationBranchConfig(stream.taskType);
        if (config && String(stream.stage || "").trim() === "Монтаж" && String(stream.status || "").trim().toLowerCase() === "рекламация") {
          fields[config.stageField] = "Завершен";
          fields[config.statusField] = "Готово";
        }
      });
      SheetService.updateFields(key, fields);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  // Ручное восстановление только исторической корпусной ветки, которая уже
  // физически смонтирована, но отсутствует в полях раздельного маршрута.
  static restoreCompletedCorpusInstallation(key) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      if (String(order.ldspStage || "").trim()) {
        throw new Error("Корпусная ветка уже существует и не требует восстановления.");
      }
      SheetService.ensureOrderColumns(["Этап ЛДСП", "Статус ЛДСП"]);
      this.updateRouteFields(key, order, {
        "Этап ЛДСП": "Монтаж",
        "Статус ЛДСП": "Готово",
        "Этап": "Монтаж",
        "Статус": "",
        "Ответственный": "Производство"
      });
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static setMdfFacade(key, enabled) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      const hasMdfFacade = Boolean(enabled);
      const fields = { "Фасад МДФ": hasMdfFacade };
      if (hasMdfFacade) fields["Без фасада"] = false;
      SheetService.ensureOrderColumn("Фасад МДФ");

      if (hasMdfFacade && String(order.stage || "").trim() === "Распил" && !String(order.mdfStage || "").trim()) {
        fields["Этап МДФ"] = "Распил МДФ";
        fields["Статус МДФ"] = "";
      }

      if (!hasMdfFacade && String(order.mdfStage || "").trim()) {
        fields["Этап МДФ"] = "";
        fields["Статус МДФ"] = "";
      }

      SheetService.updateFields(key, fields);
      return { success: true, hasMdfFacade: hasMdfFacade };
    } finally {
      lock.releaseLock();
    }
  }

  static setVeneerFacade(key, enabled) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      const hasVeneerFacade = Boolean(enabled);
      const fields = { "Фасад Шпон": hasVeneerFacade };
      if (hasVeneerFacade) fields["Без фасада"] = false;
      SheetService.ensureOrderColumn("Фасад Шпон");

      if (hasVeneerFacade && String(order.stage || "").trim() === "Распил" && !String(order.veneerStage || "").trim()) {
        fields["Этап Шпон"] = "Распил Шпон";
        fields["Статус Шпон"] = "";
      }
      if (!hasVeneerFacade && String(order.veneerStage || "").trim()) {
        fields["Этап Шпон"] = "";
        fields["Статус Шпон"] = "";
      }

      SheetService.updateFields(key, fields);
      return { success: true, hasVeneerFacade: hasVeneerFacade };
    } finally {
      lock.releaseLock();
    }
  }

  static setFacadeMaterials(key, selection) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      const requested = selection || {};
      const hasNoFacade = Boolean(requested.hasNoFacade);
      const materials = {
        hasLdspFacade: hasNoFacade ? false : Boolean(requested.hasLdspFacade),
        hasMdfFacade: hasNoFacade ? false : Boolean(requested.hasMdfFacade),
        hasVeneerFacade: hasNoFacade ? false : Boolean(requested.hasVeneerFacade),
        hasNoFacade: hasNoFacade
      };
      if (materials.hasLdspFacade || materials.hasMdfFacade || materials.hasVeneerFacade) {
        materials.hasNoFacade = false;
      }

      ["Фасад ЛДСП", "Фасад МДФ", "Фасад Шпон", "Без фасада", "Этап МДФ", "Статус МДФ", "Этап Шпон", "Статус Шпон",
        "Этап Малярка МДФ", "Статус Малярка МДФ", "Операция Малярка МДФ",
        "Этап Малярка Шпон", "Статус Малярка Шпон", "Операция Малярка Шпон"]
        .forEach(field => SheetService.ensureOrderColumn(field));
      const fields = {
        "Фасад ЛДСП": materials.hasLdspFacade,
        "Фасад МДФ": materials.hasMdfFacade,
        "Фасад Шпон": materials.hasVeneerFacade,
        "Без фасада": materials.hasNoFacade
      };
      if (materials.hasMdfFacade) {
        Object.assign(fields, this.createMissingFacadeBranchFields(order, "mdf"));
      } else if (String(order.mdfStage || "").trim()) {
        fields["Этап МДФ"] = "";
        fields["Статус МДФ"] = "";
      }

      if (materials.hasVeneerFacade) {
        Object.assign(fields, this.createMissingFacadeBranchFields(order, "veneer"));
      } else if (String(order.veneerStage || "").trim()) {
        fields["Этап Шпон"] = "";
        fields["Статус Шпон"] = "";
      }

      const createsOrRemovesRoute = Object.keys(fields).some(field => [
        "Этап МДФ", "Статус МДФ", "Этап Шпон", "Статус Шпон",
        "Этап Малярка МДФ", "Статус Малярка МДФ", "Операция Малярка МДФ",
        "Этап Малярка Шпон", "Статус Малярка Шпон", "Операция Малярка Шпон"
      ].includes(field));
      if (createsOrRemovesRoute) this.updateRouteFields(key, order, fields);
      else SheetService.updateFields(key, fields);
      return Object.assign({ success: true }, materials);
    } finally {
      lock.releaseLock();
    }
  }

  // Флажок материала не должен создавать "пустой" фасад без маршрута. До
  // запуска ветка стартует штатно с распила; для уже идущего изделия она
  // появляется в Буфере Малярки и ожидает ручного запуска покраски.
  static createMissingFacadeBranchFields(order, branch) {
    const config = {
      mdf: {
        manufacturingStage: order.mdfStage,
        paintingStage: order.mdfPaintingStage,
        packagingStage: order.mdfFacadePackagingStage,
        manufacturingStageField: "Этап МДФ",
        manufacturingStatusField: "Статус МДФ",
        paintingStageField: "Этап Малярка МДФ",
        paintingStatusField: "Статус Малярка МДФ",
        paintingOperationField: "Операция Малярка МДФ",
        initialStage: "Распил МДФ"
      },
      veneer: {
        manufacturingStage: order.veneerStage,
        paintingStage: order.veneerPaintingStage,
        packagingStage: order.veneerFacadePackagingStage,
        manufacturingStageField: "Этап Шпон",
        manufacturingStatusField: "Статус Шпон",
        paintingStageField: "Этап Малярка Шпон",
        paintingStatusField: "Статус Малярка Шпон",
        paintingOperationField: "Операция Малярка Шпон",
        initialStage: "Распил Шпон"
      }
    }[String(branch || "").trim().toLowerCase()];
    if (!config) throw new Error("Неизвестный тип фасадной ветки.");

    const hasRoute = [config.manufacturingStage, config.paintingStage, config.packagingStage]
      .some(value => String(value || "").trim());
    if (hasRoute) return {};

    const overallStage = String(order.stage || "").trim();
    if (overallStage === "Распил") {
      return {
        [config.manufacturingStageField]: config.initialStage,
        [config.manufacturingStatusField]: ""
      };
    }

    // До запуска достаточно сохранить выбор материала: launchManufacturing
    // создаст производственную ветку вместе с остальными маршрутами.
    const launchedStages = ["Кромка", "ЧПУ", "Подготовка", "Сборка", "В покраску", "Упаковка", "СГП", "Монтаж"];
    if (!launchedStages.includes(overallStage)) return {};

    const firstOperation = this.getActivePaintingOperations(order)[0];
    if (!firstOperation) throw new Error("В маршруте Малярки нет активных операций.");
    return {
      [config.paintingStageField]: "В покраску",
      [config.paintingStatusField]: "",
      [config.paintingOperationField]: firstOperation.taskType
    };
  }

  // Адресное восстановление исторической ветки: выбор шпона уже сохранён в
  // карточке 39RM, но до появления автоматического создания маршрутов ветка
  // не была заведена. Команда безопасна для повторного запуска.
  static migrateMissingVeneerFacadeToPaintingBuffer(dryRun) {
    const key = "39RM|Малыш|Прихожая";
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const order = this.getOrderByKey(key);
      if (!order) return { dryRun: Boolean(dryRun), skipped: "Изделие не найдено." };
      if (!order.hasVeneerFacade) return { dryRun: Boolean(dryRun), skipped: "Шпон не выбран." };
      const fields = this.createMissingFacadeBranchFields(order, "veneer");
      if (!Object.keys(fields).length) {
        return { dryRun: Boolean(dryRun), alreadyApplied: true };
      }
      if (!dryRun) this.updateRouteFields(key, order, fields);
      return { dryRun: Boolean(dryRun), migrated: true, fields: fields };
    } finally {
      lock.releaseLock();
    }
  }

  // Площади материалов используются для оценки загрузки производственных
  // потоков. Они не меняют состав фасадов и маршрут заказа.
  static setMaterialAreas(key, areas) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const source = areas || {};
      const definitions = [
        { key: "ldspArea", column: "ЛДСП м2", title: "ЛДСП" },
        { key: "mdfArea", column: "МДФ м2", title: "МДФ" },
        { key: "lhdfArea", column: "ЛХДФ м2", title: "ЛХДФ" },
        { key: "veneerArea", column: "Шпон м2", title: "Шпон" }
      ];
      const fields = {};
      const result = { success: true };

      definitions.forEach(definition => {
        const raw = String(source[definition.key] == null ? "" : source[definition.key]).trim();
        if (!raw) {
          fields[definition.column] = "";
          result[definition.key] = 0;
          return;
        }
        const value = Number(raw.replace(",", "."));
        if (!isFinite(value) || value < 0) {
          throw new Error("Укажите площадь «" + definition.title + "» числом не меньше нуля.");
        }
        fields[definition.column] = value;
        result[definition.key] = value;
      });

      SheetService.ensureOrderColumns(definitions.map(definition => definition.column));
      SheetService.updateFields(key, fields);
      return result;
    } finally {
      lock.releaseLock();
    }
  }

  static saveProductionDays(key, days) {
    const order = this.getOrderByKey(key);
    const operations = this.getPaintingOperations();
    const editableStages = ["В покраску"].concat(operations.map(operation => operation.stage));
    const operationCodeIsActive = operations.some(operation => operation.taskType === String(order.paintingOperationCode || ""));
    if (!editableStages.includes(order.stage) && !operationCodeIsActive) throw new Error("Нормативы операций уже недоступны для редактирования.");
    const exclusions = days.exclusions || {};
    const lacquerOperations = Boolean(days.lacquerOperations);
    const paintingArea = Number(days.paintingArea);
    if (!(paintingArea > 0)) {
      throw new Error("Укажите площадь покраски больше нуля.");
    }
    if (PaintingNormService.isHourlyEnabled()) {
      const facadeType = String(days.facadeType == null ? order.facadeType : days.facadeType).trim();
      const allowedTypes = ["", "Гладкий", "Лето", "Зима", "Фрезеровка", "Сложная фрезеровка", "Волна", "Шпон", "Глянец", "Шпон глянец"];
      if (!allowedTypes.includes(facadeType)) throw new Error("Выберите тип фасада из списка.");
      const polishingArea = Number(days.polishingArea);
      const backPaintingArea = Number(days.backPaintingArea);
      const frontPaintingArea = Number(days.frontPaintingArea);
      const fields = {
        "Площадь покраски, м2": paintingArea,
        "Площадь тыльной стороны, м2": backPaintingArea > 0 ? backPaintingArea : "",
        "Площадь лица и торцов, м2": frontPaintingArea > 0 ? frontPaintingArea : "",
        "Площадь полировки, м2": polishingArea > 0 ? polishingArea : "",
        "Тип фасада": facadeType,
        "Лаковые операции": lacquerOperations
      };
      operations.forEach(operation => {
        if (operation.exclusionColumn) fields[operation.exclusionColumn] = Boolean(exclusions[operation.daysKey]);
      });
      SheetService.ensureOrderColumns(Object.keys(fields));
      SheetService.updateFields(key, fields);
      const updated = Object.assign({}, order, {
        paintingArea,
        backPaintingArea: backPaintingArea > 0 ? backPaintingArea : 0,
        frontPaintingArea: frontPaintingArea > 0 ? frontPaintingArea : 0,
        polishingArea: polishingArea > 0 ? polishingArea : 0,
        facadeType,
        lacquerOperations,
        routeExclusions: Object.assign({}, order.routeExclusions || {}, exclusions)
      });
      return { success: true, preview: PaintingNormService.previewOrder(updated) };
    }
    // Площадь нужна для учёта загрузки ещё до согласования нормативов.
    // Если ни одна длительность пока не внесена, сохраняем только площадь:
    // не подставляем значения по умолчанию и не меняем маршрут изделия.
    const hasDurationValues = operations.some(operation =>
      String(days[operation.daysKey] == null ? "" : days[operation.daysKey]).trim() !== ""
    );
    if (!hasDurationValues) {
      SheetService.ensureOrderColumn("Площадь покраски, м2");
      SheetService.updateField(key, "Площадь покраски, м2", paintingArea);
      return { success: true, paintingAreaOnly: true };
    }
    const values = operations.map(operation => Number(days[operation.daysKey]));
    if (values.some((value, index) =>
      !exclusions[operations[index].daysKey] &&
      (!operations[index].optionalGroup || lacquerOperations) &&
      (value < 0.5 || Math.round(value * 2) !== value * 2))) {
      throw new Error("Все длительности должны быть кратны 0,5 дня и не меньше 0,5.");
    }
    const fields = { "Площадь покраски, м2": paintingArea, "Лаковые операции": lacquerOperations };
    SheetService.ensureOrderColumn("Площадь покраски, м2");
    SheetService.ensureOrderColumn("Лаковые операции");
    SheetService.ensureOrderDurationNumberFormat(operations.map(operation => operation.column));
    operations.forEach((operation, index) => {
      SheetService.ensureOrderColumn(operation.column);
      fields[operation.column] = exclusions[operation.daysKey] ||
        (operation.optionalGroup && !lacquerOperations) ? "" : values[index];
      if (operation.exclusionColumn) {
        SheetService.ensureOrderColumn(operation.exclusionColumn);
        fields[operation.exclusionColumn] = Boolean(exclusions[operation.daysKey]);
      }
    });
    SheetService.updateFields(key, fields);
    return { success: true };
  }

  static changeFacadeType(key, facadeType) {
    const allowedTypes = ["", "Гладкий", "Лето", "Зима", "Фрезеровка", "Сложная фрезеровка", "Волна", "Шпон", "Глянец", "Шпон глянец"];
    const value = String(facadeType || "").trim();
    if (!allowedTypes.includes(value)) throw new Error("Выберите тип фасада из списка.");

    const order = this.getOrderByKey(key);
    const paintingOperations = this.getPaintingOperations();
    const paintingStages = ["В покраску"].concat(paintingOperations.map(operation => operation.stage));
    if (!paintingStages.includes(String(order.stage || "").trim()) &&
        !paintingOperations.some(operation => operation.taskType === String(order.paintingOperationCode || ""))) {
      throw new Error("Тип фасада можно изменить только для изделия в разделе Малярка.");
    }

    if (PaintingNormService.isHourlyEnabled() && String(order.stage || "").trim() !== "В покраску" &&
        !PaintingNormService.getProfileForFacadeType(value)) {
      throw new Error("Для выбранного типа фасада не настроен нормативный профиль.");
    }

    SheetService.ensureOrderColumn("Тип фасада");
    SheetService.updateField(key, "Тип фасада", value);
    return { success: true, facadeType: value };
  }

  static validateProductionDays(order) {
    if (PaintingNormService.isHourlyEnabled()) {
      const preview = PaintingNormService.previewOrder(order);
      if (!(Number(order.paintingArea) > 0)) throw new Error("Перед запуском укажите площадь покраски больше нуля.");
      if (!preview.profile) throw new Error("Для выбранного типа фасада не настроен нормативный профиль.");
      if (!preview.valid) {
        const invalid = preview.operations.find(operation => !operation.valid);
        throw new Error((invalid && invalid.message) || "Заполните нормы маршрута Малярки.");
      }
      return;
    }
    const values = this.getConfiguredPaintingOperations(order)
      .map(operation => Number((order.productionDays || {})[operation.daysKey]));
    if (values.some(value => value < 0.5 || Math.round(value * 2) !== value * 2)) {
      throw new Error("Перед завершением укажите длительность всех включённых операций Малярки.");
    }
  }

  static assignSandingWorker(orders) {
    const workers = this.getRoleResponsible("Шлифовка", "Амир", "Сергей");
    const tasks = this.getSandingTasks(orders);
    const load = workers.reduce((result, worker) => {
      result[worker] = tasks
        .filter(task => task.constructor === worker)
        .reduce((sum, task) => sum + (Number(task.duration) || 0), 0);
      return result;
    }, {});
    const worker = workers.slice().sort((left, right) => load[left] - load[right])[0];
    const maxQueue = tasks.reduce((max, task) => Math.max(max, Number(task.queue) || 0), 0);
    return { worker, primaryQueue: maxQueue + 10, whiteQueue: maxQueue + 11 };
  }

  static getPaintingBranchState(order, branch) {
    const name = String(branch || "legacy").trim().toLowerCase();
    const corpusConfig = this.getAdditionalCorpusBranchConfig(name);
    if (corpusConfig) return {
      branch: name,
      stage: order[corpusConfig.paintingStageProperty],
      status: order[corpusConfig.paintingStatusProperty],
      operationCode: order[corpusConfig.paintingOperationProperty],
      stageField: corpusConfig.paintingStageField,
      statusField: corpusConfig.paintingStatusField,
      operationField: corpusConfig.paintingOperationField,
      returnStageField: corpusConfig.stageField,
      returnStatusField: corpusConfig.statusField,
      returnToAssembly: true
    };
    if (name === "mdf") return {
      branch: name, stage: order.mdfPaintingStage, status: order.mdfPaintingStatus,
      operationCode: order.mdfPaintingOperationCode, stageField: "Этап Малярка МДФ",
      statusField: "Статус Малярка МДФ", operationField: "Операция Малярка МДФ",
      packagingStageField: "Этап упаковки фасада МДФ", packagingStatusField: "Статус упаковки фасада МДФ"
    };
    if (name === "veneer") return {
      branch: name, stage: order.veneerPaintingStage, status: order.veneerPaintingStatus,
      operationCode: order.veneerPaintingOperationCode, stageField: "Этап Малярка Шпон",
      statusField: "Статус Малярка Шпон", operationField: "Операция Малярка Шпон",
      packagingStageField: "Этап упаковки фасада Шпон", packagingStatusField: "Статус упаковки фасада Шпон"
    };
    return {
      branch: "legacy", stage: order.stage, status: order.status,
      operationCode: order.paintingOperationCode, stageField: "Этап", statusField: "Статус",
      operationField: "Операция Малярки", packagingStageField: "Этап упаковки фасада",
      packagingStatusField: "Статус упаковки фасада"
    };
  }

  static getPaintingBranchOrder(order, branch) {
    const state = this.getPaintingBranchState(order, branch);
    return Object.assign({}, order, {
      stage: state.stage,
      status: state.status,
      paintingOperationCode: state.operationCode
    });
  }

  static changeSandingStatus(key, taskType, action, comment, branch) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      SheetService.ensureOrderColumn("Малярка — ручной приоритет");
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);
      if (!order) throw new Error("Заказ не найден. Обновите очередь и повторите действие.");
      const state = this.getPaintingBranchState(order, branch);
      const branchOrder = this.getPaintingBranchOrder(order, state.branch);
      const operations = this.getActivePaintingOperations(branchOrder);
      const operationIndex = operations.findIndex(item => item.taskType === taskType);
      const operation = operations[operationIndex];
      const currentMatches = order && operation && (
        branchOrder.stage === operation.stage || String(branchOrder.paintingOperationCode || "") === operation.taskType
      );
      if (!order || !operation || operation.role !== "sanding" || !currentMatches) {
        throw new Error("Операция недоступна в очереди шлифовки.");
      }

      const allTasks = this.getSandingTasks(orders);
      const currentTask = allTasks.find(task => (task.orderKey || task.key) === key && task.paintingBranch === state.branch && task.taskType === taskType);
      const today = CalendarService.firstWorkDay();
      const queueField = "Очередь шлифовки";
      const fields = {};

      if (action === "start") {
        // Запуск остаётся доступным всегда. Если он не соответствует
        // прогнозу, изделие получает приоритет до конца маршрута.
        fields["Малярка — ручной приоритет"] = this.isManualPaintingStart(orders, key, taskType);
        fields[state.statusField] = "В работе";
        fields["Плановая дата шлифовки"] = "";
      } else if (action === "pause") {
        const pauseComment = String(comment || "").trim();
        if (!pauseComment) throw new Error("Укажите причину паузы.");
        const nextDay = CalendarService.nextWorkDay(today);
        currentTask.plannedDate = CalendarService.formatDate(nextDay);
        currentTask.queue = this.getSandingQueueValue(currentTask, allTasks, nextDay, true);
        fields[queueField] = currentTask.queue;
        fields[state.statusField] = "Пауза";
        fields["Комментарий"] = pauseComment;
        fields["Плановая дата шлифовки"] = currentTask.plannedDate;
      } else if (action === "stop") {
        const nextOperation = operations.slice(operationIndex + 1)
          .find(item => !(order.routeExclusions || {})[item.daysKey]);
        fields[state.statusField] = "";
        fields[state.stageField] = nextOperation ? nextOperation.stage : "";
        fields[state.operationField] = nextOperation ? nextOperation.taskType : "";
        if (!nextOperation) {
          if (state.returnToAssembly) {
            fields[state.returnStageField] = "Сборка";
            fields[state.returnStatusField] = "";
          } else {
            fields[state.packagingStageField] = "Упаковка Фасада";
            fields[state.packagingStatusField] = "";
          }
        }
        fields["Плановая дата шлифовки"] = "";
      } else {
        throw new Error("Неизвестное действие шлифовки.");
      }

      this.updateRouteFields(key, order, fields);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static changeSandingWorker(key, worker) {
    const order = this.getOrderByKey(key);
    const operations = this.getActivePaintingOperations(order);
    const currentIndex = operations.findIndex(operation => operation.stage === String(order.stage || "").trim());
    const hasRemainingSanding = currentIndex >= 0 && operations
      .slice(currentIndex)
      .some(operation => operation.role === "sanding");
    if (!hasRemainingSanding) {
      throw new Error("У изделия не осталось операций шлифовки для переназначения.");
    }
    const availableWorkers = this.getRoleResponsible("Шлифовка", "Амир", "Сергей");
    const selectedWorker = String(worker || "").trim();
    if (!availableWorkers.includes(selectedWorker)) {
      throw new Error("Выберите шлифовщика из справочника.");
    }
    SheetService.ensureOrderColumn("Шлифовщик");
    SheetService.updateField(key, "Шлифовщик", selectedWorker);
    return { success: true };
  }

  static isManualPaintingStart(orders, key, taskType) {
    const selected = orders.find(order => order.key === key);
    if (!selected || selected.paintingManualPriority) return Boolean(selected && selected.paintingManualPriority);

    const schedule = this.getPaintingRouteSchedule(orders);
    const target = schedule.find(task => (task.orderKey || task.key) === key && task.taskType === taskType && task.canStart);
    if (!target) return false;
    const expected = schedule
      .filter(task => task.canStart && task.constructor === target.constructor &&
        String(task.status || "").trim().toLowerCase() !== "в работе")
      .sort((left, right) => {
        const leftStart = new Date(left.scheduleStart || 0).getTime();
        const rightStart = new Date(right.scheduleStart || 0).getTime();
        return leftStart - rightStart || Number(left.dayPart !== "Ⅰ половина") - Number(right.dayPart !== "Ⅰ половина") ||
          Number(left.queue || 0) - Number(right.queue || 0);
      })[0];
    return Boolean(expected && (expected.key !== key || expected.taskType !== taskType));
  }

  static getSandingQueueValue(currentTask, allTasks, date, firstInQueue) {
    if (!PaintingNormService.isHourlyEnabled()) CalendarService.buildCalendar(allTasks);
    const peers = allTasks
      .filter(task => task !== currentTask && task.constructor === currentTask.constructor)
      .sort((left, right) => left.queue - right.queue);
    const index = firstInQueue ? 0 : peers.findIndex(task => date <= task.scheduleFinish);
    const position = index === -1 ? peers.length : index;
    return this.getQueueValue(peers[position - 1], peers[position]);
  }

  static assignPainter(orders) {
    // Покраска выполняется одним маляром, поэтому это единая очередь.
    const painters = this.getRoleResponsible("Покраска", "Денис");
    const tasks = this.getPaintingTasks(orders);
    const load = painters.reduce((result, painter) => {
      result[painter] = tasks.filter(task => task.constructor === painter)
        .reduce((sum, task) => sum + (Number(task.duration) || 0), 0);
      return result;
    }, {});
    const painter = painters.slice().sort((left, right) => load[left] - load[right])[0];
    const maxQueue = tasks.reduce((max, task) => Math.max(max, Number(task.queue) || 0), 0);
    return { painter, primerQueue: maxQueue + 10, paintingQueue: maxQueue + 11 };
  }

  static getPainterName(value) {
    const name = String(value || "").trim();
    const painters = this.getRoleResponsible("Покраска", "Денис");
    // Нормализуем старые или отсутствующие значения к единственному списку
    // ответственных из справочника.
    return painters.includes(name) ? name : painters[0];
  }

  static getRoleResponsible(dictionaryName, ...fallback) {
    const dictionaries = DictionaryService.getAll();
    const expected = String(dictionaryName || "").trim().toLowerCase();
    if (expected === "шлифовка") {
      const available = this.getAvailableSandingWorkers(dictionaries);
      return available.length ? available : fallback;
    }
    const actualKey = Object.keys(dictionaries).find(key =>
      String(key || "").trim().toLowerCase() === expected
    );
    const values = ((actualKey ? dictionaries[actualKey] : []) || [])
      .map(value => String(value || "").trim())
      .filter(Boolean);
    return values.length ? values : fallback;
  }

  static getAvailableSandingWorkers(dictionaries) {
    const source = dictionaries || DictionaryService.getAll();
    const dictionaryKey = Object.keys(source).find(key =>
      String(key || "").trim().toLowerCase() === "шлифовка"
    );
    const allWorkers = ((dictionaryKey ? source[dictionaryKey] : []) || [])
      .map(value => String(value || "").trim())
      .filter(Boolean);
    if (!allWorkers.length) return ["Амир"];

    const stored = PropertiesService.getScriptProperties().getProperty("availableSandingWorkers");
    if (!stored) return allWorkers;
    try {
      const selected = JSON.parse(stored);
      const available = (Array.isArray(selected) ? selected : [])
        .map(value => String(value || "").trim())
        .filter(value => allWorkers.includes(value));
      return available.length ? available : allWorkers;
    } catch (error) {
      return allWorkers;
    }
  }

  static setAvailableSandingWorkers(workers) {
    const dictionaries = DictionaryService.getAll();
    const dictionaryKey = Object.keys(dictionaries).find(key =>
      String(key || "").trim().toLowerCase() === "шлифовка"
    );
    const allWorkers = ((dictionaryKey ? dictionaries[dictionaryKey] : []) || [])
      .map(value => String(value || "").trim())
      .filter(Boolean);
    const selected = [...new Set((Array.isArray(workers) ? workers : [])
      .map(value => String(value || "").trim())
      .filter(value => allWorkers.includes(value)))];
    if (!selected.length) throw new Error("Укажите хотя бы одного доступного шлифовщика.");
    PropertiesService.getScriptProperties().setProperty("availableSandingWorkers", JSON.stringify(selected));
    return { availableWorkers: selected, count: selected.length };
  }

  static getSandingWorkerName(value) {
    const worker = String(value || "").trim();
    const workers = this.getRoleResponsible("Шлифовка", "Амир", "Сергей");
    return workers.includes(worker) ? worker : workers[0];
  }

  static changePaintingStatus(key, taskType, action, comment, branch) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      SheetService.ensureOrderColumn("Малярка — ручной приоритет");
      const orders = SheetService.getObjects(CONFIG.SHEETS.ORDERS).map(row => OrderMapper.map(row));
      const order = orders.find(item => item.key === key);
      if (!order) throw new Error("Заказ не найден. Обновите очередь и повторите действие.");
      const state = this.getPaintingBranchState(order, branch);
      const branchOrder = this.getPaintingBranchOrder(order, state.branch);
      const operations = this.getActivePaintingOperations(branchOrder);
      const operationIndex = operations.findIndex(item => item.taskType === taskType);
      const operation = operations[operationIndex];
      const currentMatches = order && operation && (
        branchOrder.stage === operation.stage || String(branchOrder.paintingOperationCode || "") === operation.taskType
      );
      if (!order || !operation || operation.role !== "painting" || !currentMatches) {
        throw new Error("Операция недоступна в очереди покраски.");
      }

      const allTasks = this.getPaintingTasks(orders);
      const currentTask = allTasks.find(task => (task.orderKey || task.key) === key && task.paintingBranch === state.branch && task.taskType === taskType);
      const today = CalendarService.firstWorkDay();
      const queueField = "Очередь грунт";
      const fields = {};
      if (action === "start") {
        // Запуск не блокируется: внеочередное изделие становится первым в
        // прогнозе для всех оставшихся операций.
        fields["Малярка — ручной приоритет"] = this.isManualPaintingStart(orders, key, taskType);
        fields[state.statusField] = "В работе";
        fields["Плановая дата покраски"] = "";
      } else if (action === "pause") {
        const pauseComment = String(comment || "").trim();
        if (!pauseComment) throw new Error("Укажите причину паузы.");
        const nextDay = CalendarService.nextWorkDay(today);
        currentTask.plannedDate = CalendarService.formatDate(nextDay);
        currentTask.queue = this.getPaintingQueueValue(currentTask, allTasks, nextDay, true);
        fields[queueField] = currentTask.queue;
        fields[state.statusField] = "Пауза";
        fields["Комментарий"] = pauseComment;
        fields["Плановая дата покраски"] = currentTask.plannedDate;
      } else if (action === "stop") {
        const nextOperation = operations.slice(operationIndex + 1)
          .find(item => !(order.routeExclusions || {})[item.daysKey]);
        fields[state.statusField] = "";
        fields[state.stageField] = nextOperation ? nextOperation.stage : "";
        fields[state.operationField] = nextOperation ? nextOperation.taskType : "";
        if (!nextOperation) {
          if (state.returnToAssembly) {
            fields[state.returnStageField] = "Сборка";
            fields[state.returnStatusField] = "";
          } else {
            fields[state.packagingStageField] = "Упаковка Фасада";
            fields[state.packagingStatusField] = "";
          }
        }
        fields["Плановая дата покраски"] = "";
      } else {
        throw new Error("Неизвестное действие покраски.");
      }
      this.updateRouteFields(key, order, fields);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  }

  static getPaintingQueueValue(currentTask, allTasks, date, firstInQueue) {
    if (!PaintingNormService.isHourlyEnabled()) CalendarService.buildCalendar(allTasks);
    const peers = allTasks.filter(task => task !== currentTask && task.constructor === currentTask.constructor)
      .sort((left, right) => left.queue - right.queue);
    const index = firstInQueue ? 0 : peers.findIndex(task => date <= task.scheduleFinish);
    const position = index === -1 ? peers.length : index;
    return this.getQueueValue(peers[position - 1], peers[position]);
  }

  static placeOrderAtDate(order, queueOrders, date, firstInQueue) {
    const peers = queueOrders
      .filter(item => item !== order && item.constructor === order.constructor)
      .sort((left, right) => left.queue - right.queue);
    const index = firstInQueue ? 0 : peers.findIndex(item => date <= item.scheduleFinish);
    const position = index === -1 ? peers.length : index;
    order.queue = this.getQueueValue(peers[position - 1], peers[position]);
  }

  static getNextStage(stage) {
    const normalized = String(stage || "").trim();
    if (normalized === "КД") return "Очередь";
    return { "Проформа": "Согласование" }[normalized] || stage;
  }

  /**
   * Изменяет ответственного
   */
  static changeResponsible(
    key,
    responsible
  ) {


    return this.updateField(

      key,

      "Ответственный",

      responsible

    );


  }





  /**
   * Изменяет этап
   */
  static changeStage(
    key,
    stage
  ) {


    return this.updateField(

      key,

      "Этап",

      stage

    );


  }


}
