/**
 * Единый контур снабжения: спецификации, закупочные статусы и поставщики.
 */
class SupplyService {
  static get HEADERS() {
    return [
      "ID", "Договор", "Клиент", "Изделие", "Артикул", "Наименование",
      "Количество", "Единица", "Поставщик", "Цена", "Статус закупки",
      "Поставщик вручную", "Файл источника", "Обновлено", "Способ сопоставления", "Производитель", "Ссылка поставщика", "Категория",
      "Лист источника", "Строка заголовка", "Поставщик источника", "Версия импорта", "Исключено из версии",
      "Закупочное количество", "Единица закупки", "Статус сопоставления", "Источник позиции", "Комментарий заявки"
    ];
  }

  static get STATUSES() {
    return ["Новый", "Заказ поставщику", "Счёт получен", "Оплачено", "Товар получен"];
  }

  // Группы индикаторов снабжения для обзорной таблицы АРМ мастера. Состав
  // фиксирован: UI и Excel используют те же ключи и не делают собственную
  // классификацию строк снабжения.
  static get MASTER_SIGNAL_GROUPS() {
    return [
      { key: "ldsp", title: "ЛДСП", shape: "square" },
      { key: "blum", title: "Блюм", shape: "circle" },
      { key: "hardware", title: "Фурнитура", shape: "circle" },
      { key: "modus", title: "Модус", shape: "triangle" },
      { key: "duslar", title: "Дюслар", shape: "triangle" },
      { key: "glass", title: "Стекло", shape: "diamond" },
      { key: "other", title: "Прочее", shape: "diamond" }
    ];
  }

  // Возвращает светофоры только для уже отобранных строк АРМ. Чтение обоих
  // листов выполняется пакетно, поэтому snapshot не превращается в N+1 запрос.
  static getMasterSupplySignals(orders) {
    const sourceOrders = Array.isArray(orders) ? orders : [];
    const groups = this.MASTER_SIGNAL_GROUPS;
    const orderByKey = new Map();
    sourceOrders.forEach(order => {
      const key = SheetService.buildOrderLookupKey(order.contract, order.client, order.product);
      if (key !== "||") orderByKey.set(key, order);
    });
    const positionsByOrder = new Map(Array.from(orderByKey.keys()).map(key => [key, groups.reduce((result, group) => {
      result[group.key] = [];
      return result;
    }, {})]));
    if (!positionsByOrder.size) return {};

    const addPosition = (row, columns, groupResolver) => {
      const key = SheetService.buildOrderLookupKey(row[columns.contract], row[columns.client], row[columns.product]);
      const bucket = positionsByOrder.get(key);
      if (!bucket) return;
      const group = groupResolver(row, columns);
      if (group && bucket[group]) bucket[group].push(String(row[columns.status] || "Новый").trim());
    };
    const supplySheet = SheetService.getSheet(CONFIG.SHEETS.SUPPLY);
    if (supplySheet && supplySheet.getLastRow() > 1) {
      const rows = supplySheet.getDataRange().getValues();
      const headers = rows[0].map(value => String(value || "").trim());
      const columns = {
        contract: headers.indexOf("Договор"), client: headers.indexOf("Клиент"), product: headers.indexOf("Изделие"),
        supplier: headers.indexOf("Поставщик"), name: headers.indexOf("Наименование"), manufacturer: headers.indexOf("Производитель"),
        category: headers.indexOf("Категория"), status: headers.indexOf("Статус закупки"), excluded: headers.indexOf("Исключено из версии"), origin: headers.indexOf("Источник позиции")
      };
      if (Object.entries(columns).every(([key, index]) => key === "origin" || index >= 0)) {
        rows.slice(1).forEach(row => {
          if (String(row[columns.excluded] || "").trim() === "Да" || (columns.origin >= 0 && String(row[columns.origin] || "") === "Ручная заявка")) return;
          addPosition(row, columns, (item, fields) => this.getMasterSupplyGroup({
            supplier: item[fields.supplier], name: item[fields.name], manufacturer: item[fields.manufacturer], category: item[fields.category]
          }));
        });
      }
    }

    const plateSheet = SheetService.getSheet(CONFIG.SHEETS.PLATE_MATERIALS);
    if (plateSheet && plateSheet.getLastRow() > 1) {
      const rows = plateSheet.getDataRange().getValues();
      const headers = rows[0].map(value => String(value || "").trim());
      const columns = {
        contract: headers.indexOf("Договор"), client: headers.indexOf("Клиент"), product: headers.indexOf("Изделие"),
        category: headers.indexOf("Категория"), status: headers.indexOf("Статус закупки")
      };
      if (Object.values(columns).every(index => index >= 0)) {
        rows.slice(1).forEach(row => {
          if (this.normalize(row[columns.category]) !== "лдсп") return;
          addPosition(row, columns, () => "ldsp");
        });
      }
    }

    const result = {};
    positionsByOrder.forEach((positions, key) => {
      const order = orderByKey.get(key);
      result[key] = groups.reduce((signals, group) => {
        const statuses = positions[group.key];
        signals[group.key] = this.getMasterSupplySignal(statuses, group.key === "ldsp" && !statuses.length && this.hasMasterLdspPassedSaw(order));
        return signals;
      }, {});
    });
    return result;
  }

  static getMasterSupplyGroup(position) {
    const supplier = this.normalize(position.supplier);
    const name = this.normalize(position.name);
    const manufacturer = this.normalize(position.manufacturer);
    if (supplier === "ом гласс") return "glass";
    if (/\bblum\b|блюм/.test(manufacturer) || /\bblum\b|блюм/.test(name)) return "blum";
    if (supplier === "мдм" || supplier.startsWith("мдм ")) return "hardware";
    if (/\bmodus\b/.test(supplier) || supplier.includes("модус")) return "modus";
    if (/\bduslar\b/.test(supplier) || supplier.includes("дюслар")) return "duslar";
    if (supplier === "расходники" || this.normalize(position.category) === "расходники") return null;
    return "other";
  }

  static getMasterSupplySignal(statuses, ldspSawPassed) {
    const source = Array.isArray(statuses) ? statuses : [];
    if (!source.length) return ldspSawPassed
      ? { state: "green", label: "Получено" }
      : { state: "none", label: "Нет позиций" };
    const normalized = source.map(status => String(status || "Новый").trim().replace(/^Счет получен$/, "Счёт получен"));
    const allReceived = normalized.every(status => status === "Товар получен");
    const hasNoInvoice = normalized.some(status => !["Счёт получен", "Оплачено", "Товар получен"].includes(status));
    if (hasNoInvoice) return { state: "red", label: "Счёт не получен" };
    if (normalized.some(status => status === "Счёт получен")) return { state: "lilac", label: "Счёт получен — ожидает оплаты" };
    if (allReceived) return { state: "green", label: "Получено" };
    return { state: "yellow", label: "Оплачено" };
  }

  static hasMasterLdspPassedSaw(order) {
    if (!order) return false;
    const afterSaw = new Set([
      "Кромка", "Присадка", "Сборка", "В покраску", "Упаковка", "Упаковка Корпуса",
      "СГП", "СГП Корпуса", "Монтаж", "Завершен"
    ]);
    return [order.ldspStage, order.paintedLdspStage].some(stage => afterSaw.has(String(stage || "").trim())) ||
      Boolean(String(order.paintedLdspPaintingStage || "").trim());
  }

  static getSnapshot() {
    SheetService.ensureOrderDeadlineColumn();
    // Миграция справочника определений выполняется при первом открытии
    // снабжения, чтобы новая колонка артикула появилась без отдельной команды.
    this.ensureConsumablesSheet();
    const sheet = this.ensureSheet();
    SupplierOrderService.reconcileMaterialStatuses();
    const values = sheet.getDataRange().getValues();
    const deadlines = this.getOrderDirectory().reduce((result, order) => {
      result[order.key] = order.deadline || "";
      return result;
    }, {});
    const items = values.slice(1).filter(row => row[0]).map(row => ({
      rowId: String(row[0]), contract: row[1], client: row[2], product: row[3],
      deadline: deadlines[[row[1], row[2], row[3]].join("|")] || "",
      article: row[4], name: row[5], quantity: row[6], unit: row[7],
      supplier: row[8], price: row[9], purchaseStatus: row[10] || "Новый",
      supplierManual: row[11] === true, sourceFile: row[12], updatedAt: row[13],
      matchMethod: row[14], manufacturer: row[15] || "", productUrl: row[16] || "", category: row[17] || "",
      purchaseQuantity: row[23] === "" || row[23] === null ? "" : Number(row[23]), purchaseUnit: row[24] || "", matchStatus: row[25] || "",
      excludedFromVersion: String(row[22] || "") === "Да", manualRequest: String(row[26] || "") === "Ручная заявка", requestComment: String(row[27] || "")
    }));
    // Для ручного выбора показываем реестр поставщиков, даже если у него ещё
    // нет позиций в очереди. «Расходники» — системный вариант выбора.
    const suppliers = [...new Set(
      this.getSupplierRegistry().filter(item => item.active).map(item => item.name)
        .concat(items.map(item => String(item.supplier || "").trim()).filter(Boolean), ["Расходники"])
    )]
      .sort((a, b) => a.localeCompare(b, "ru"));
    return JSON.parse(JSON.stringify({
      section: "supply",
      orders: items,
      dictionaries: {
        suppliers: suppliers,
        purchaseStatuses: this.STATUSES,
        supplierOrders: SupplierOrderService.getOrders()
      }
    }));
  }

  static ensureSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SUPPLY);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.SUPPLY);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, this.HEADERS.length)
        .setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
    } else if (sheet.getLastColumn() < this.HEADERS.length || sheet.getRange(1, this.HEADERS.length).getDisplayValue() !== this.HEADERS[this.HEADERS.length - 1]) {
      sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]);
      sheet.getRange(1, 1, 1, this.HEADERS.length)
        .setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
    }
    this.ensureStatusValidation(sheet);
    return sheet;
  }

  static ensureStatusValidation(sheet) {
    // В старых строках проверка данных была создана до появления статуса
    // «Заказ поставщику». Обновляем её один раз после публикации, не меняя
    // сами значения очереди при обычном чтении.
    const cache = CacheService.getScriptCache();
    const cacheKey = "supply-status-validation-v2";
    if (cache.get(cacheKey)) return;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(this.STATUSES, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 11, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
    cache.put(cacheKey, "1", 21600);
  }

  static findOrderCandidates(fileName) {
    const text = this.normalize(fileName);
    const orders = this.getOrderDirectory();
    const contractMatches = orders.filter(order => {
      const contract = this.compact(order.contract);
      return contract && this.compact(text).includes(contract);
    });
    const pool = contractMatches.length ? contractMatches : orders;
    const scored = pool.map(order => {
      const productTokens = this.tokens(order.product);
      const clientTokens = this.tokens(order.client);
      const tokens = new Set(this.tokens(text));
      const productScore = productTokens.filter(token => tokens.has(token)).length * 3;
      const clientScore = clientTokens.filter(token => tokens.has(token)).length * 2;
      return Object.assign({}, order, { score: productScore + clientScore + (contractMatches.includes(order) ? 10 : 0) });
    }).filter(order => order.score > 0)
      .sort((a, b) => b.score - a.score || String(a.product).localeCompare(String(b.product), "ru"));
    if (!scored.length) throw new Error("По имени файла не найден заказ в листе «Изделия».");
    return scored.slice(0, 10).map(order => ({
      key: order.key, contract: order.contract, client: order.client,
      product: order.product, score: order.score
    }));
  }

  // Селектор импорта всегда показывает весь справочник изделий. Совпадения по
  // имени файла нужны только для удобства: они выводятся первыми и выбираются
  // по умолчанию, но не должны скрывать остальные заказы.
  static getImportOrderCandidates(fileName) {
    const orders = this.getOrderDirectory();
    if (!orders.length) throw new Error("В листе «Изделия» нет доступных заказов.");
    let preferred = [];
    try {
      preferred = this.findOrderCandidates(fileName);
    } catch (error) {
      preferred = [];
    }
    const preferredKeys = {};
    preferred.forEach(order => { preferredKeys[order.key] = true; });
    const rest = orders.filter(order => !preferredKeys[order.key]).sort((left, right) =>
      [left.contract, left.client, left.product].join("|").localeCompare(
        [right.contract, right.client, right.product].join("|"), "ru"));
    return preferred.concat(rest).map(order => ({
      key: order.key,
      contract: order.contract,
      client: order.client,
      product: order.product
    }));
  }

  static getOrderDirectory() {
    const rows = SheetService.objectsFromRows(SheetService.getData(CONFIG.SHEETS.ORDERS));
    return rows.filter(row => row["Договор"] && row["Изделие"]).map(row => ({
      contract: String(row["Договор"]).trim(), client: String(row["Клиент"] || "").trim(),
      product: String(row["Изделие"]).trim(),
      key: [row["Договор"], row["Клиент"], row["Изделие"]].join("|"),
      deadline: row["Дедлайн"] || ""
    }));
  }

  static getManualRequestOrderDirectory() {
    const rows = SheetService.objectsFromRows(SheetService.getData(CONFIG.SHEETS.ORDERS));
    const unique = new Map();
    rows.forEach(row => {
      const contract = String(row["Договор"] || "").trim();
      const client = String(row["Клиент"] || "").trim();
      const product = String(row["Изделие"] || "").trim();
      if (!product || !client) return;
      const key = [contract, client, product].join("|");
      if (!unique.has(key)) unique.set(key, { key, contract, client, product });
    });
    return Array.from(unique.values()).sort((a, b) => a.key.localeCompare(b.key, "ru"));
  }

  static saveManualRequest(data) {
    const input = data || {};
    const name = String(input.name || "").trim();
    const article = String(input.article || "").trim();
    const unit = String(input.unit || "").trim();
    const comment = String(input.comment || "").trim();
    const selectedClient = String(input.client || "").trim();
    const rawQuantity = String(input.quantity == null ? "" : input.quantity).trim();
    const quantity = rawQuantity ? Number(rawQuantity.replace(",", ".")) : "";
    if (!name) throw new Error("Укажите наименование.");
    if (rawQuantity && (!Number.isFinite(quantity) || quantity <= 0)) throw new Error("Количество должно быть больше нуля.");
    const key = String(input.orderKey || "");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const sheet = this.ensureSheet();
      if (input.rowId) {
        const row = this.findRow(sheet, input.rowId);
        if (String(sheet.getRange(row, 27).getValue() || "") !== "Ручная заявка") throw new Error("Можно изменять только ручные заявки.");
        const hasProduct = Boolean(String(sheet.getRange(row, 4).getValue() || "").trim());
        const previousClient = String(sheet.getRange(row, 3).getValue() || "").trim();
        if (!hasProduct && selectedClient && selectedClient !== previousClient && !EquipmentService.getClients().includes(selectedClient)) throw new Error("Выбранный клиент не найден.");
        sheet.getRange(row, 5, 1, 4).setValues([[article, name, quantity, unit]]);
        if (!hasProduct) sheet.getRange(row, 3).setValue(selectedClient);
        sheet.getRange(row, 28).setValue(comment);
        sheet.getRange(row, 14).setValue(new Date());
        return { rowId: String(input.rowId) };
      }
      const order = key ? this.getManualRequestOrderDirectory().find(item => item.key === key) : null;
      if (key && !order) throw new Error("Выбранное изделие не найдено.");
      if (!order && selectedClient && !EquipmentService.getClients().includes(selectedClient)) throw new Error("Выбранный клиент не найден.");
      const row = Array(this.HEADERS.length).fill("");
      row[0] = Utilities.getUuid();
      if (order) { row[1] = order.contract; row[2] = order.client; row[3] = order.product; }
      else row[2] = selectedClient;
      row[4] = article; row[5] = name; row[6] = quantity; row[7] = unit;
      row[10] = "Новый"; row[13] = new Date(); row[26] = "Ручная заявка"; row[27] = comment;
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
      return { rowId: row[0] };
    } finally { lock.releaseLock(); }
  }

  static getMaterialOrderDirectory() {
    const rows = this.ensureSheet().getDataRange().getValues().slice(1);
    const unique = {};
    rows.forEach(row => {
      const contract = String(row[1] || "").trim();
      const client = String(row[2] || "").trim();
      const product = String(row[3] || "").trim();
      if (!contract || !product) return;
      const key = [contract, client, product].join("|");
      unique[key] = { key: key, contract: contract, client: client, product: product };
    });
    return Object.keys(unique).map(key => unique[key]).sort((left, right) =>
      [left.contract, left.client, left.product].join("|").localeCompare([right.contract, right.client, right.product].join("|"), "ru"));
  }

  static importSpecification(file, orderKey) {
    const preview = this.previewSpecification(file, orderKey);
    return this.saveSpecification(preview.materials, file.name, orderKey, preview.importInfo);
  }

  static analyzeSpecification(file, orderKey) {
    const format = this.getSpecificationFormat(file);
    if (format === "xml") return this.analyzeXmlSpecification(file);
    const order = this.getOrderDirectory().find(item => item.key === String(orderKey || ""));
    if (!order) throw new Error("Выбранный заказ больше не найден в листе «Изделия».");
    const candidates = this.readUploadedWorkbook(file, workbook => this.discoverMaterialTables(workbook));
    if (!candidates.length) throw new Error("В файле не найдены закупочные таблицы с наименованием и количеством.");
    return { format: format, candidates: candidates };
  }

  static previewSpecification(file, orderKey, selections) {
    const format = this.getSpecificationFormat(file);
    const order = this.getOrderDirectory().find(item => item.key === String(orderKey || ""));
    if (!order) throw new Error("Выбранный заказ больше не найден в листе «Изделия».");
    if (format === "xml") return this.previewXmlSpecification(file, order);
    const materials = this.readUploadedWorkbook(file, workbook => this.extractMaterials(workbook, selections));
    if (!materials.length) throw new Error("В файле не найдены строки материалов с наименованием и количеством.");
    const priceIndex = this.buildPriceIndex();
    const itktMatches = {};
    const materialRoutes = this.getMaterialRoutesIndex();
    const materialMappings = this.getMaterialMappingsIndex();
    const supplierRegistry = this.getSupplierRegistry();
    const supplierAliases = this.getSupplierAliasIndex();
    return {
      materials: materials.map(material => {
        const materialRoute = this.getMaterialRoute(material, materialRoutes);
        const plateMatch = PlateCatalogService.match(material);
        const isPlate = Boolean(materialRoute.plate || plateMatch.category || this.isEdgeMaterial(material.name));
        const isConsumable = !isPlate && Boolean(materialRoute.consumable);
        const knownMatch = isConsumable ? null : this.matchKnownMaterial(material, materialMappings);
        // Название производителя из спецификации — наиболее точный источник.
        // Если оно внесено в справочник, не заменяем его результатом поиска
        // по прайсу или сайту другого поставщика.
        const mappedSupplier = this.matchSupplierByManufacturer(material, supplierAliases);
        const siteMatch = isConsumable || knownMatch || mappedSupplier ? null : (this.isBlumMaterial(material) ? (this.matchItktBlum(material, itktMatches) || this.matchBlumRus(material)) :
          (this.isDuslarMaterial(material) ? this.matchDuslar(material) :
            (this.isSheetMaterial(material) ? this.matchExpoTorg(material) : null)));
        const configuredSiteMatch = isConsumable || knownMatch || mappedSupplier || siteMatch ? null :
          this.matchConfiguredSite(material, supplierRegistry, itktMatches);
        const match = isConsumable || knownMatch || mappedSupplier ? { supplier: "", price: "", method: "", productUrl: "" } :
          (siteMatch || configuredSiteMatch || this.matchSupplier(material, priceIndex));
        return Object.assign({}, material, {
          supplier: isConsumable ? "Расходники" : (knownMatch ? knownMatch.supplier : (match.supplier || mappedSupplier)),
          price: isConsumable ? "" : (knownMatch ? knownMatch.price : match.price),
          matchMethod: isConsumable ? "Расходник" : (knownMatch ? knownMatch.method : (match.method || (mappedSupplier ? "По названию поставщика" : ""))),
          productUrl: isConsumable ? "" : (knownMatch ? knownMatch.productUrl : (match.productUrl || "")),
          category: isConsumable ? "Расходники" : (isPlate ? "Плитный материал" : ""),
          plateCategory: plateMatch.category || (this.isEdgeMaterial(material.name) ? "Кромка" : ""), plateDecor: plateMatch.decor, plateThickness: plateMatch.thickness
        });
      })
    };
  }

  static saveSpecification(materials, fileName, orderKey, importInfo) {
    const order = this.getOrderDirectory().find(item => item.key === String(orderKey || ""));
    if (!order) throw new Error("Выбранный заказ больше не найден в листе «Изделия».");
    if (!Array.isArray(materials) || !materials.length) throw new Error("Нет материалов для импорта.");
    const isXml = importInfo && importInfo.format === "xml";
    const preparedMaterials = (materials || []).map(material => this.applyConsumableSelection(material));
    if (isXml) this.ensureXmlSuppliers(preparedMaterials);
    const activeSuppliers = this.getSupplierRegistry().filter(item => item.active);
    const routedToPlate = preparedMaterials.filter(material => material.category === "Плитный материал");
    if (routedToPlate.length) PlateMaterialService.mergeSpecificationMaterials(routedToPlate, fileName, order);
    const cleanMaterials = preparedMaterials.filter(material => material.category !== "Плитный материал").map(material => ({
      article: String(material.article || "").trim(), name: String(material.name || "").trim(),
      quantity: this.toNumber(material.quantity), unit: String(material.unit || "").trim(),
      manufacturer: String(material.manufacturer || "").trim(), supplier: this.resolveXmlSupplier(material, activeSuppliers, isXml),
      price: this.toNumber(material.price), matchMethod: String(material.matchMethod || "").trim(),
      productUrl: String(material.productUrl || "").trim(), category: String(material.category || "").trim(),
      purchaseQuantity: this.toNumber(material.purchaseQuantity), purchaseUnit: String(material.purchaseUnit || "").trim(), matchStatus: String(material.matchStatus || "").trim(),
      sourceSheet: isXml && importInfo.manualOrderSelection ? "XML · изделие выбрано вручную" : String(material.sourceSheet || "").trim(), sourceHeaderRow: material.sourceHeaderRow || "",
      sourceSupplier: String(material.sourceSupplier || "").trim(),
      supplierManual: Boolean(isXml && material.supplierManual), consumableManual: Boolean(material.consumableManual), xmlHardware: Boolean(material.xmlHardware)
    })).filter(material => material.name && Number.isFinite(material.quantity) && material.quantity > 0);
    if (!cleanMaterials.length) return {
      imported: 0, unresolved: [], routedToPlate: routedToPlate.length,
      quantityWarnings: isXml && Array.isArray(importInfo.quantityWarnings) ? importInfo.quantityWarnings : []
    };
    const sheet = this.ensureSheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const rows = sheet.getDataRange().getValues();
      const existing = {};
      rows.slice(1).forEach(row => {
        if ([row[1], row[2], row[3]].join("|") !== order.key || String(row[26] || "") === "Ручная заявка") return;
        existing[String(row[0])] = row;
      });
      const now = new Date();
      const nextIds = new Set();
      const orderedMaterialIds = new Set(SupplierOrderService.getOrderedMaterialIds());
      const importVersion = Utilities.getUuid();
      const importedRows = cleanMaterials.map(material => {
        const calculatedRowId = this.buildRowId(order.key, material);
        const old = existing[calculatedRowId] || this.findExistingManualXmlRow(existing, material);
        const rowId = old ? String(old[0]) : calculatedRowId;
        nextIds.add(rowId);
        const manual = Boolean((old && old[11] === true) || material.supplierManual);
        return [
          rowId, order.contract, order.client, order.product, manual && old ? old[4] : material.article,
          material.name, material.quantity, material.unit,
          manual && old ? old[8] : material.supplier,
          manual && old ? old[9] : (Number.isFinite(material.price) ? material.price : ""),
          old ? (old[10] || "Новый") : "Новый",
          manual,
          fileName,
          now,
          manual ? "Вручную" : material.matchMethod,
          material.manufacturer,
          material.productUrl,
          material.category,
          material.sourceSheet || "", material.sourceHeaderRow || "", material.sourceSupplier || "", importVersion, "",
          Number.isFinite(material.purchaseQuantity) && material.purchaseQuantity > 0 ? material.purchaseQuantity : "", material.purchaseUnit || "", material.matchStatus || "", "", ""
        ];
      });
      const keep = rows.slice(1).filter(row => {
        const sameOrder = [row[1], row[2], row[3]].join("|") === order.key;
        if (!sameOrder || String(row[26] || "") === "Ручная заявка" || nextIds.has(String(row[0]))) return true;
        if (!orderedMaterialIds.has(String(row[0]))) return false;
        row[22] = "Да";
        row[13] = now;
        return true;
      });
      const otherOrders = keep.filter(row => [row[1], row[2], row[3]].join("|") !== order.key);
      const preservedOrderedRows = keep.filter(row => [row[1], row[2], row[3]].join("|") === order.key && !nextIds.has(String(row[0])));
      const finalRows = otherOrders.concat(importedRows, preservedOrderedRows);
      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, this.HEADERS.length).clearContent();
      if (finalRows.length) sheet.getRange(2, 1, finalRows.length, this.HEADERS.length).setValues(finalRows);
      this.rememberManualConsumableRoutes(cleanMaterials.filter(material => material.consumableManual));
      // Глобальное правило создаётся только после ручной корректировки
      // строки в XML-предпросмотре. Автоматический импорт его не меняет.
      if (isXml) this.rememberConfirmedXmlMaterialMappings(cleanMaterials.filter(material => material.supplierManual));
      const unresolved = importedRows.filter(row => !row[8]).map(row => row[5]);
      return {
        imported: importedRows.length, unresolved: unresolved.slice(0, 50), routedToPlate: routedToPlate.length,
        manuallySelectedOrder: Boolean(isXml && importInfo.manualOrderSelection),
        sourceOrder: isXml ? importInfo.sourceOrder || null : null,
        quantityWarnings: isXml && Array.isArray(importInfo.quantityWarnings) ? importInfo.quantityWarnings : []
      };
    } finally {
      lock.releaseLock();
    }
  }

  static resolveXmlSupplier(material, activeSuppliers, isXml) {
    const selected = String(material && material.supplier || "").trim();
    if (!isXml || !selected) return selected;
    if (this.normalize(selected) === this.normalize("Расходники")) return "Расходники";
    const aliasIndex = this.getSupplierAliasIndex(activeSuppliers);
    const canonical = aliasIndex[this.normalize(selected)] || "";
    // XML агента нельзя использовать для неявного создания поставщиков:
    // неизвестное исходное значение остаётся невыбранным до решения снабженца.
    if (material.xmlHardware && !canonical) throw new Error("Для позиции «" + String(material.name || "") + "» выберите поставщика из справочника.");
    return canonical || selected;
  }

  static applyConsumableSelection(material) {
    const selected = String(material && material.supplier || "").trim();
    if (this.normalize(selected) !== this.normalize("Расходники")) return material;
    return Object.assign({}, material, {
      supplier: "Расходники", category: "Расходники", price: "", productUrl: "",
      matchMethod: material.consumableManual ? "Расходник · вручную" : (material.matchMethod || "Расходник")
    });
  }

  static findExistingManualXmlRow(existing, material) {
    if (!material || !material.supplierManual) return null;
    const sourceSheet = this.normalize(material.sourceSheet);
    const sourceRow = String(material.sourceHeaderRow || "").trim();
    return Object.keys(existing || {}).map(key => existing[key]).find(row => Boolean(row[11]) &&
      (sourceRow ? String(row[19] || "").trim() === sourceRow : this.normalize(row[18]) === sourceSheet) &&
      this.normalizeName(row[5]) === this.normalizeName(material.name) && this.normalize(row[7]) === this.normalize(material.unit)) || null;
  }

  static readUploadedWorkbook(file, reader) {
    const bytes = Utilities.base64Decode(String(file.base64 || ""));
    if (!bytes.length || bytes.length > 8 * 1024 * 1024) throw new Error("Файл пуст или превышает 8 МБ.");
    const blob = Utilities.newBlob(bytes, file.mimeType || "application/vnd.ms-excel", file.name);
    const temp = Drive.Files.insert({ title: "TEMP_SUPPLY_" + Date.now() }, blob, { convert: true });
    try {
      return reader(SpreadsheetApp.openById(temp.id));
    } finally {
      Drive.Files.trash(temp.id);
    }
  }

  static getSpecificationFormat(file) {
    if (!file) throw new Error("Выберите файл спецификации в формате XLS или XML.");
    const name = String(file.name || "").trim();
    const mimeType = String(file.mimeType || "").toLowerCase();
    if (/\.xml$/i.test(name) || /(?:application|text)\/xml/.test(mimeType)) return "xml";
    if (/\.xls$/i.test(name) || mimeType === "application/vnd.ms-excel") return "xls";
    throw new Error("Выберите файл спецификации в формате XLS или XML.");
  }

  static readXmlSpecification(file) {
    const bytes = Utilities.base64Decode(String(file && file.base64 || ""));
    if (!bytes.length || bytes.length > 8 * 1024 * 1024) throw new Error("XML-файл пуст или превышает 8 МБ.");
    let document;
    try {
      document = XmlService.parse(Utilities.newBlob(bytes).getDataAsString("UTF-8").replace(/^\uFEFF/, ""));
    } catch (error) {
      throw new Error("Не удалось прочитать XML спецификации: " + error.message);
    }
    const root = document.getRootElement();
    if (!root) throw new Error("В XML отсутствует корневой элемент.");
    if (["hardwareProcurement", "procurementSpecification", "procurement"].indexOf(root.getName()) >= 0) return this.readHardwareProcurementXml(root, file);
    if (root.getName() !== "supply-specification") throw new Error("Поддерживаются XML <hardwareProcurement>, <procurementSpecification>, <procurement> или <supply-specification>.");
    if (String(root.getAttribute("version") && root.getAttribute("version").getValue() || "") !== "1") {
      throw new Error("Поддерживается только XML-формат supply-specification версии 1.");
    }
    const text = (element, name) => String(element && element.getChild(name) && element.getChild(name).getText() || "").trim();
    const orderElement = root.getChild("order");
    const sourceOrder = { contract: text(orderElement, "contract"), client: text(orderElement, "client"), product: text(orderElement, "product") };
    if (!orderElement || !sourceOrder.contract || !sourceOrder.client || !sourceOrder.product) {
      throw new Error("XML должен содержать договор, клиента и изделие в блоке <order>.");
    }
    const materialElements = root.getChild("materials") && root.getChild("materials").getChildren("material") || [];
    if (!materialElements.length) throw new Error("В XML нет ни одной позиции <material>.");
    const aggregated = {};
    const quantityWarnings = [];
    materialElements.forEach((element, index) => {
      const sourceQuantity = text(element, "quantity");
      const material = {
        article: text(element, "article"), name: text(element, "name"), quantity: this.toNumber(sourceQuantity), unit: text(element, "unit"),
        supplier: text(element, "supplier"), manufacturer: text(element, "manufacturer"), price: this.toNumber(text(element, "price")),
        productUrl: text(element, "product-url"), category: text(element, "category"), plateCategory: text(element, "plate-category"),
        plateDecor: text(element, "decor"), plateThickness: this.toNumber(text(element, "thickness-mm")),
        sourceSupplier: text(element, "supplier"), sourceSheet: "XML", sourceHeaderRow: "", matchMethod: "XML: внешняя система",
        xmlCandidateVerified: true, xmlVerifiedSupplier: text(element, "supplier"), xmlVerifiedArticle: text(element, "article")
      };
      if (!material.name) {
        quantityWarnings.push("XML-позиция №" + (index + 1) + " без наименования пропущена.");
        return;
      }
      const quantity = this.normalizeXmlSourceQuantity(sourceQuantity, material.name);
      material.quantity = quantity.value;
      if (quantity.warning) quantityWarnings.push(quantity.warning);
      if (!material.unit) throw new Error("В XML-позиции «" + material.name + "» не указана единица измерения.");
      if (!material.supplier) throw new Error("В XML-позиции «" + material.name + "» не указан поставщик.");
      if (text(element, "price") && (!Number.isFinite(material.price) || material.price < 0)) throw new Error("В XML-позиции «" + material.name + "» указана некорректная цена.");
      if (material.productUrl && !/^https?:\/\//i.test(material.productUrl)) throw new Error("В XML-позиции «" + material.name + "» указана некорректная ссылка на товар.");
      if (["", "Расходники", "Плитный материал"].indexOf(material.category) < 0) throw new Error("В XML-позиции «" + material.name + "» неизвестная категория.");
      if (material.category === "Плитный материал" && material.plateCategory && ["ЛДСП", "ЛМДФ", "ЛХДФ", "МДФ", "Кромка"].indexOf(material.plateCategory) < 0) {
        throw new Error("В XML-позиции «" + material.name + "» неизвестная категория плиты.");
      }
      const identity = this.normalizeArticle(material.article) || (this.normalizeName(material.name) + "|" + this.normalize(material.unit));
      if (aggregated[identity] && !this.sameSupplier(aggregated[identity].supplier, material.supplier)) {
        throw new Error("В XML одинаковая позиция «" + material.name + "» указана у разных поставщиков.");
      }
      if (aggregated[identity]) aggregated[identity].quantity += material.quantity;
      else aggregated[identity] = material;
    });
    return { sourceOrder: sourceOrder, materials: Object.keys(aggregated).map(key => aggregated[key]), quantityWarnings: quantityWarnings };
  }

  static readHardwareProcurementXml(root, file) {
    const version = String(root.getAttribute("version") && root.getAttribute("version").getValue() || "");
    const isProcurementV05 = root.getName() === "procurement";
    if ((isProcurementV05 && version !== "0.5.0") || (!isProcurementV05 && version !== "1.0")) {
      throw new Error(isProcurementV05 ? "Поддерживается XML procurement версии 0.5.0." : "Поддерживается XML фурнитуры версии 1.0.");
    }
    const text = (element, name) => String(element && element.getChild(name) && element.getChild(name).getText() || "").trim();
    const attribute = (element, name) => String(element && element.getAttribute(name) && element.getAttribute(name).getValue() || "").trim();
    const source = root.getChild("source") || (isProcurementV05 ? root.getChild("metadata") : null);
    const sourceFile = text(source, "fileName") || text(source, "sourceFile") || attribute(root, "sourceFile") || String(file && file.name || "");
    const sourceSheet = text(source, "sheetName") || text(source, "sourceSheet") || attribute(root, "sourceSheet") || "Фурнитура";
    const sourceOrderElement = root.getChild("order");
    const sourceOrder = sourceOrderElement ? { contract: text(sourceOrderElement, "contract"), client: text(sourceOrderElement, "client"), product: text(sourceOrderElement, "product") } : null;
    const items = root.getChild("items") && root.getChild("items").getChildren("item") || [];
    if (!items.length) throw new Error("В XML фурнитуры нет ни одной позиции <item>.");
    const quantityWarnings = [];
    const materials = items.map((item, index) => {
      const original = item.getChild("original");
      const sourceRow = text(item, "sourceRow") || attribute(item, "sourceRow") || String(index + 1);
      const importAction = text(item, "importAction").toUpperCase();
      // Агент явно отделяет заголовки и материалы, не входящие в закупку
      // фурнитуры. Эти строки не являются ошибкой XML и не должны получать
      // искусственное количество при импорте.
      if (importAction === "EXCLUDE") {
        const excludedName = text(original, "name") || "без наименования";
        quantityWarnings.push("XML-позиция из строки " + sourceRow + " «" + excludedName + "» исключена агентом и не импортирована.");
        return null;
      }
      const matching = item.getChild("matching") || item.getChild("match");
      const matchStatus = attribute(item, "matchStatus") || attribute(matching, "status") || text(matching, "status") || "UNVERIFIED";
      // Агент выпускает фактический <matching><candidate>, а описание
      // плагина — <match><candidates><candidate>. Поддерживаем оба контракта.
      const candidatesNode = matching && matching.getChild("candidates");
      const candidates = matching ? matching.getChildren("candidate").concat(candidatesNode ? candidatesNode.getChildren("candidate") : []) : [];
      const preferred = candidates.find(candidate => attribute(candidate, "preferred").toLowerCase() === "true") || null;
      // В XML procurement 0.5 подтверждённые данные лежат прямо в <match>,
      // а не во вложенном <candidate>. Оба представления равнозначны.
      const directMatch = matching && ["EXACT", "ANALOG"].indexOf(matchStatus) >= 0 &&
        (text(matching, "supplier") || text(matching, "supplierId")) ? matching : null;
      const candidateValue = (name) => attribute(preferred, name) || text(preferred, name) || text(directMatch, name);
      const candidateSupplier = candidateValue("supplier") || candidateValue("supplierId");
      const eligible = ["EXACT", "ANALOG"].indexOf(matchStatus) >= 0 && Boolean(preferred || directMatch);
      const purchase = item.getChild("purchase");
      const purchaseStatus = attribute(purchase, "status") || text(purchase, "status");
      const purchaseQuantity = this.toNumber(text(purchase, "purchaseQuantity"));
      const purchaseUnit = text(purchase, "unit") || text(purchase, "purchaseUnit");
      const sourceQuantity = text(original, "requiredQuantity") || text(original, "quantity");
      const material = {
        article: eligible ? (candidateValue("article") || text(original, "article")) : text(original, "article"),
        name: text(original, "name"), quantity: this.toNumber(sourceQuantity), unit: text(original, "unit"),
        supplier: eligible ? candidateSupplier : "", manufacturer: eligible ? candidateValue("manufacturer") : "",
        productUrl: eligible ? candidateValue("url") : "", price: "", category: "", plateCategory: "", plateDecor: "", plateThickness: "",
        purchaseQuantity: ["AS_SPECIFIED", "CALCULATED"].indexOf(purchaseStatus) >= 0 ? purchaseQuantity : "",
        purchaseUnit: ["AS_SPECIFIED", "CALCULATED"].indexOf(purchaseStatus) >= 0 ? purchaseUnit : "",
        matchStatus: matchStatus, matchMethod: eligible ? (matchStatus === "ANALOG" ? "XML: Аналог — проверить" : "XML: точное сопоставление") : "XML: " + matchStatus,
        sourceSupplier: text(item, "sourceSupplier") || text(original, "supplier"), sourceSheet: sourceSheet, sourceHeaderRow: sourceRow,
        xmlHardware: true, xmlCandidateVerified: Boolean(eligible), xmlVerifiedSupplier: eligible ? candidateSupplier : "", xmlVerifiedArticle: eligible ? (candidateValue("article") || text(original, "article")) : ""
      };
      if (!material.name) {
        quantityWarnings.push("XML-позиция из строки " + sourceRow + " без наименования пропущена.");
        return null;
      }
      const quantity = this.normalizeXmlSourceQuantity(sourceQuantity, material.name, true);
      material.quantity = quantity.value;
      if (quantity.warning) quantityWarnings.push(quantity.warning);
      if (!material.unit) {
        material.unit = "шт";
        quantityWarnings.push("В позиции «" + material.name + "» единица измерения не указана; установлено «шт». Проверьте строку перед формированием заказа.");
      }
      if (eligible && !material.supplier) throw new Error("В XML-позиции «" + material.name + "» нет поставщика подтверждённого кандидата.");
      if (["AS_SPECIFIED", "CALCULATED"].indexOf(purchaseStatus) >= 0 && (!Number.isFinite(purchaseQuantity) || purchaseQuantity <= 0 || !purchaseUnit)) {
        throw new Error("В XML-позиции «" + material.name + "» некорректное закупочное количество.");
      }
      const plateMatch = PlateCatalogService.match(material);
      const isPlate = Boolean(plateMatch.category || this.isEdgeMaterial(material.name));
      if (isPlate) {
        material.category = "Плитный материал";
        material.plateCategory = plateMatch.category || (this.isEdgeMaterial(material.name) ? "Кромка" : "");
        material.plateDecor = plateMatch.decor || "";
        material.plateThickness = plateMatch.thickness || "";
      }
      return material;
    }).filter(Boolean);
    if (!materials.length) throw new Error("В XML фурнитуры нет позиций с наименованием.");
    return { sourceOrder: sourceOrder, sourceFile: sourceFile, sourceSheet: sourceSheet, materials: materials, hardwareXml: true, quantityWarnings: quantityWarnings };
  }

  static analyzeXmlSpecification(file) {
    const specification = this.readXmlSpecification(file);
    const resolution = this.resolveXmlOrder(specification);
    return {
      format: "xml", sourceOrder: specification.sourceOrder, candidates: resolution.candidates,
      matchedOrderKey: resolution.matchedOrderKey, requiresManualSelection: !resolution.matchedOrderKey,
      materialCount: specification.materials.length, sourceFile: specification.sourceFile || "", matchSource: resolution.matchSource,
      quantityWarnings: specification.quantityWarnings || []
    };
  }

  static resolveXmlOrder(specification) {
    const orders = this.getOrderDirectory();
    const matches = specification.sourceOrder ? orders.filter(order => this.sameOrder(order, specification.sourceOrder)) : [];
    if (matches.length === 1) return {
      candidates: matches.map(order => ({ key: order.key, contract: order.contract, client: order.client, product: order.product })),
      matchedOrderKey: matches[0].key, matchSource: "xml-key"
    };
    let filenameCandidates = [];
    if (specification.sourceFile) {
      try {
        filenameCandidates = this.findOrderCandidates(specification.sourceFile);
      } catch (error) { filenameCandidates = []; }
    }
    const first = filenameCandidates[0];
    const second = filenameCandidates[1];
    // Одного кандидата либо строго лучшего по договору, клиенту и изделию
    // достаточно для безопасного автоподбора. При равном балле оставляем выбор человеку.
    const isUniqueFilenameMatch = first && (!second || Number(first.score) > Number(second.score));
    return {
      candidates: filenameCandidates.length ? filenameCandidates.map(item => ({ key: item.key, contract: item.contract, client: item.client, product: item.product })) : orders.map(order => ({ key: order.key, contract: order.contract, client: order.client, product: order.product })),
      matchedOrderKey: isUniqueFilenameMatch ? first.key : "", matchSource: isUniqueFilenameMatch ? "source-file" : ""
    };
  }

  static previewXmlSpecification(file, order) {
    const specification = this.readXmlSpecification(file);
    const resolution = this.resolveXmlOrder(specification);
    const manuallySelected = resolution.matchedOrderKey !== order.key;
    const activeSuppliers = this.getSupplierRegistry().filter(item => item.active);
    const aliases = this.getSupplierAliasIndex(activeSuppliers);
    const materialMappings = this.getMaterialMappingsIndex();
    const materialRoutes = this.getMaterialRoutesIndex();
    const materials = specification.materials.map(material => {
      const confirmedMapping = this.matchKnownMaterial(material, materialMappings);
      const verifiedSupplier = material.xmlCandidateVerified ? (aliases[this.normalize(material.supplier)] || material.supplier) : "";
      const suggestedSupplier = material.xmlHardware && !material.xmlCandidateVerified ? (aliases[this.normalize(material.sourceSupplier)] || "") : "";
      const route = this.getMaterialRoute(material, materialRoutes);
      const plateMatch = PlateCatalogService.match(material);
      const isPlate = Boolean(route.plate || plateMatch.category || this.isEdgeMaterial(material.name));
      const isConsumable = !isPlate && Boolean(route.consumable);
      return Object.assign({}, material, {
        supplier: isConsumable ? "Расходники" : (confirmedMapping ? confirmedMapping.supplier : (verifiedSupplier || material.supplier || suggestedSupplier)),
        price: isConsumable ? "" : (confirmedMapping && Number.isFinite(confirmedMapping.price) ? confirmedMapping.price : material.price),
        matchMethod: isConsumable ? "Расходник" : (confirmedMapping ? confirmedMapping.method : material.matchMethod),
        productUrl: isConsumable ? "" : (confirmedMapping ? confirmedMapping.productUrl : material.productUrl),
        category: isConsumable ? "Расходники" : (isPlate ? "Плитный материал" : material.category),
        plateCategory: plateMatch.category || material.plateCategory || (this.isEdgeMaterial(material.name) ? "Кромка" : ""),
        plateDecor: plateMatch.decor || material.plateDecor || "",
        plateThickness: plateMatch.thickness || material.plateThickness || "",
        xmlVerifiedSupplier: verifiedSupplier || material.xmlVerifiedSupplier,
        supplierManual: false,
        xmlSuggestedSupplier: suggestedSupplier,
        mappingApplied: Boolean(confirmedMapping)
      });
    });
    return {
      materials: materials,
      importInfo: { format: "xml", sourceOrder: specification.sourceOrder, sourceFile: specification.sourceFile || "", manualOrderSelection: manuallySelected, matchSource: resolution.matchSource, quantityWarnings: specification.quantityWarnings || [] },
      format: "xml"
    };
  }

  static sameOrder(first, second) {
    if (!first || !second || !first.contract || !first.client || !first.product || !second.contract || !second.client || !second.product) return false;
    return this.normalize(first && first.contract) === this.normalize(second && second.contract) &&
      this.normalize(first && first.client) === this.normalize(second && second.client) &&
      this.normalize(first && first.product) === this.normalize(second && second.product);
  }

  static ensureXmlSuppliers(materials) {
    const existing = this.getSupplierRegistry();
    const known = existing.map(item => item.name);
    const aliases = this.getSupplierAliasIndex(existing);
    [...new Set((materials || []).filter(item => !item.xmlHardware || item.xmlCandidateVerified)
      .map(item => {
        const name = String(item.supplier || "").trim();
        return aliases[this.normalize(name)] || name;
      }).filter(name => name && name !== "Расходники"))].forEach(name => {
      if (known.some(item => this.sameSupplier(item, name))) return;
      this.saveSupplierRegistryItem({ name: name, active: true, note: "Добавлен из XML спецификации" });
      known.push(name);
    });
  }

  static discoverMaterialTables(workbook) {
    return workbook.getSheets().map((sheet, index) => {
      const values = sheet.getDataRange().getDisplayValues();
      const header = this.findHeader(values);
      if (!header || !this.isPurchasingTable(sheet.getName(), values, header)) return null;
      const count = this.countMaterialRows(values, header);
      if (!count) return null;
      return {
        id: String(index), sheetName: sheet.getName(), headerRow: header.row + 1, rowCount: count,
        confidence: header.confidence, warnings: header.warnings,
        supplierColumn: header.supplier,
        columns: values[header.row].map((value, column) => ({ index: column, label: String(value || "Колонка " + (column + 1)) }))
      };
    }).filter(Boolean);
  }

  static extractMaterials(workbook, selections) {
    const aggregated = {};
    const selected = new Map((selections || []).filter(item => item && item.selected !== false).map(item => [String(item.id), item]));
    const candidates = this.discoverMaterialTables(workbook).filter(item => selected.size ? selected.has(item.id) : true);
    if (!candidates.length) throw new Error("Выберите хотя бы одну закупочную таблицу.");
    candidates.forEach(candidate => {
      const sheet = workbook.getSheets()[Number(candidate.id)];
      const values = sheet.getDataRange().getDisplayValues();
      const header = this.findHeader(values);
      if (!header) return;
      const selectedConfig = selected.get(candidate.id) || {};
      if (Number.isFinite(Number(selectedConfig.supplierColumn))) header.supplier = Number(selectedConfig.supplierColumn);
      const sheetMaterials = [];
      for (let rowIndex = header.row + 1; rowIndex < values.length; rowIndex++) {
        const row = values[rowIndex];
        const name = String(row[header.name] || "").trim();
        const quantity = this.toNumber(row[header.quantity]);
        // Импортируем исключительно строки тела таблицы: строка шапки,
        // подзаголовки и итоги не могут стать материалами.
        if (!name || !Number.isFinite(quantity) || quantity <= 0 || this.isNonMaterialRow(name)) continue;
        sheetMaterials.push({
          article: header.article >= 0 ? String(row[header.article] || "").trim() : "",
          name: name,
          quantity: quantity,
          unit: header.unit >= 0 ? String(row[header.unit] || "").trim() : "",
          manufacturer: header.supplier >= 0 ? String(row[header.supplier] || "").trim() : this.detectManufacturerFromName(name),
          sourceSupplier: header.supplier >= 0 ? String(row[header.supplier] || "").trim() : "",
          sourceSheet: sheet.getName(), sourceHeaderRow: header.row + 1
        });
      }
      sheetMaterials.forEach(material => {
        const key = this.normalizeArticle(material.article) || (this.normalizeName(material.name) + "|" + this.normalize(material.unit));
        if (!aggregated[key]) aggregated[key] = material;
        else aggregated[key].quantity += material.quantity;
      });
    });
    return Object.keys(aggregated).map(key => aggregated[key]);
  }

  static countMaterialRows(values, header) {
    let count = 0;
    for (let rowIndex = header.row + 1; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      const name = String(row[header.name] || "").trim();
      const quantity = this.toNumber(row[header.quantity]);
      if (name && Number.isFinite(quantity) && quantity > 0 && !this.isNonMaterialRow(name)) count++;
    }
    return count;
  }

  static isPurchasingTable(sheetName, values, header) {
    const sheet = this.normalize(sheetName);
    if (/малярк|сбороч|раскро|лдсп|мдф/.test(sheet)) return false;
    const headers = values[header.row].map(value => this.normalize(value)).join(" |");
    if (/длина|ширина|готовая деталь|заготовка/.test(headers)) return false;
    return header.article >= 0 || header.explicitUnit || header.supplier >= 0 || /фурнитур|материал|покупн/.test(sheet);
  }

  static findHeader(values) {
    const aliases = {
      article: ["артикул", "код", "код товара", "обозначение"],
      name: ["наименование", "материал", "наименование материала", "название"],
      quantity: ["кол во", "количество", "количество в заказе", "кол.", "кол-во", "qty"],
      unit: ["ед", "ед.", "единица", "ед изм", "ед. изм."],
      supplier: ["производитель", "изготовитель", "бренд", "торговая марка", "марка", "фирма", "поставщик", "примечание"]
    };
    for (let r = 0; r < Math.min(values.length, 40); r++) {
      const normalized = values[r].map(value => this.normalize(value));
      // В файлах БАЗИС заголовки часто уточняются: например,
      // «Количество в заказе» и «Наименование материала».
      const locate = names => normalized.findIndex(value => names.some(name => {
        const alias = this.normalize(name);
        return value === alias || value.indexOf(alias + " ") === 0 || value.endsWith(" " + alias);
      }));
      const name = locate(aliases.name);
      const quantity = locate(aliases.quantity);
      if (name >= 0 && quantity >= 0) {
        let unit = locate(aliases.unit);
        const explicitUnit = unit >= 0;
        // В XLS из Базис колонка единицы измерения часто находится сразу
        // после количества, но её заголовок визуально пустой.
        if (unit < 0 && quantity + 1 < values[r].length) unit = quantity + 1;
        const supplierOnHeader = locate(aliases.supplier);
        // В XLS БАЗИС шапка часто объединена и «Производитель» может быть
        // на другой строке в верхней части таблицы. Ищем по всей шапке.
        let supplier = supplierOnHeader >= 0 ? supplierOnHeader : this.findHeaderColumn(values, aliases.supplier, quantity + 1);
        // Старый экспорт БАЗИС: после количества идут единица и короткий код поставщика без заголовков.
        if (supplier < 0 && !explicitUnit && quantity + 2 < values[r].length) supplier = quantity + 2;
        return {
          row: r, name: name, quantity: quantity, article: locate(aliases.article), unit: unit,
          supplier: supplier, explicitUnit: explicitUnit,
          confidence: (locate(aliases.article) >= 0 ? "Высокая" : "Средняя"),
          warnings: supplier < 0 ? ["Поставщик не определён: можно выбрать колонку вручную."] : []
        };
      }
    }
    return null;
  }

  static isNonMaterialRow(name) {
    const value = this.normalize(name);
    return !value || /^(итого|всего|наименование|наименование материала|материалы|фурнитура|примечание)$/.test(value) ||
      /^(итого|всего)\b/.test(value);
  }

  static findNearbyHeaderColumn(values, row, aliases, minColumn) {
    const start = Math.max(0, row - 2);
    const end = Math.min(values.length - 1, row + 2);
    for (let r = start; r <= end; r++) {
      const column = values[r].findIndex((value, index) => {
        if (Number.isFinite(minColumn) && index < minColumn) return false;
        const text = this.normalize(value);
        return aliases.some(alias => {
          const normalizedAlias = this.normalize(alias);
          return text === normalizedAlias || text.indexOf(normalizedAlias + " ") === 0 || text.endsWith(" " + normalizedAlias);
        });
      });
      if (column >= 0) return column;
    }
    return -1;
  }

  static findHeaderColumn(values, aliases, minColumn) {
    const maxRows = Math.min(values.length, 40);
    for (let row = 0; row < maxRows; row++) {
      for (let column = Math.max(0, minColumn || 0); column < values[row].length; column++) {
        const text = this.normalize(values[row][column]);
        if (aliases.some(alias => {
          const value = this.normalize(alias);
          return text === value || text.indexOf(value + " ") === 0 || text.endsWith(" " + value) || text.indexOf(" " + value + " ") >= 0;
        })) return column;
      }
    }
    return -1;
  }

  static detectManufacturerFromName(name) {
    const value = String(name || "");
    const brands = ["BLUM", "POLLY", "JINIO", "HETTICH", "BOYARD", "GTV", "KROMOSPAN", "EGGER", "KASTAMONU"];
    const found = brands.find(brand => new RegExp(`(^|[^A-ZА-Я])${brand}($|[^A-ZА-Я])`, "i").test(value));
    return found || "";
  }

  static buildPriceIndex(suppliers) {
    const index = { articles: {}, names: {}, entries: [], suppliers: [] };
    const activeSuppliers = (suppliers || this.getSupplierRegistry()).filter(item => item.active && item.priceUrl);
    activeSuppliers.forEach(item => {
      const files = this.getPriceFiles(item);
      files.forEach(file => this.addPriceFileToIndex(file, item.name, index));
    });
    return index;
  }

  static getPriceFiles(item) {
    const folderId = this.extractFolderId(item.priceUrl);
    if (folderId) {
      try {
        const files = [];
        const iterator = DriveApp.getFolderById(folderId).getFiles();
        while (iterator.hasNext()) {
          const file = iterator.next();
          if (this.isPriceFile(file)) files.push(file);
        }
        return files;
      } catch (error) {
        throw new Error("Не удалось открыть папку прайсов поставщика «" + item.name + "». Проверьте доступ.");
      }
    }
    const fileId = this.extractDriveFileId(item.priceUrl);
    if (!fileId) return []; // Веб-сайт обрабатывается поиском по его каталогу.
    try {
      const file = DriveApp.getFileById(fileId);
      if (!this.isPriceFile(file)) throw new Error("Файл должен быть Google Sheets, XLS или XLSX.");
      return [file];
    } catch (error) {
      throw new Error("Не удалось открыть прайс поставщика «" + item.name + "». Проверьте доступ и формат файла.");
    }
  }

  static isPriceFile(file) {
    return /\.(xls|xlsx)$/i.test(file.getName()) || file.getMimeType() === MimeType.GOOGLE_SHEETS;
  }

  static addPriceFileToIndex(file, supplier, index) {
    if (!index.suppliers.includes(supplier)) index.suppliers.push(supplier);
    const spreadsheetId = this.toNativeSpreadsheet(file);
    try {
      const entries = this.extractPriceEntries(SpreadsheetApp.openById(spreadsheetId), supplier);
      entries.forEach(entry => {
        const article = this.normalizeArticle(entry.article);
        const name = this.normalizeName(entry.name);
        if (article && !index.articles[article]) index.articles[article] = entry;
        if (name && !index.names[name]) index.names[name] = entry;
        index.entries.push(entry);
      });
    } finally {
      if (spreadsheetId !== file.getId()) Drive.Files.trash(spreadsheetId);
    }
  }

  static getSupplierRegistry() {
    const sheet = this.ensureSupplierRegistrySheet();
    return sheet.getDataRange().getValues().slice(1).map(row => ({
      name: String(row[0] || "").trim(), priceUrl: String(row[1] || "").trim(),
      active: row[2] !== false && String(row[2]).toLowerCase() !== "нет",
      note: String(row[3] || "").trim(), aliases: String(row[4] || "").trim(),
      searchTemplate: String(row[5] || "").trim(), websiteSource: this.isWebsiteSourceUrl(row[1])
    })).filter(item => item.name);
  }

  static saveSupplierRegistryItem(item) {
    const name = String(item && item.name || "").trim();
    const priceUrl = String(item && item.priceUrl || "").trim();
    let searchTemplate = String(item && item.searchTemplate || "").trim();
    if (!name) throw new Error("Укажите название поставщика.");
    if (priceUrl && !/^https?:\/\//i.test(priceUrl)) throw new Error("Укажите корректную ссылку на сайт, папку или файл.");
    if (searchTemplate && (!/^https?:\/\//i.test(searchTemplate) || searchTemplate.indexOf("{query}") < 0)) {
      throw new Error("Шаблон поиска должен быть ссылкой с подстановкой {query}.");
    }
    if (!searchTemplate && this.isWebsiteSourceUrl(priceUrl)) searchTemplate = this.discoverSearchTemplate(priceUrl);
    const sheet = this.ensureSupplierRegistrySheet();
    const values = sheet.getDataRange().getValues();
    const row = values.findIndex((value, index) => index > 0 && this.normalize(value[0]) === this.normalize(name));
    const record = [[name, priceUrl, item.active !== false, String(item.note || "").trim(), String(item.aliases || "").trim(), searchTemplate]];
    if (row > 0) sheet.getRange(row + 1, 1, 1, 6).setValues(record);
    else sheet.getRange(sheet.getLastRow() + 1, 1, 1, 6).setValues(record);
    return this.getSupplierRegistry();
  }

  static deleteSupplierRegistryItem(name) {
    const sheet = this.ensureSupplierRegistrySheet();
    const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getDisplayValues();
    const index = values.findIndex(row => this.normalize(row[0]) === this.normalize(name));
    if (index < 0) throw new Error("Поставщик не найден.");
    sheet.deleteRow(index + 2);
    return this.getSupplierRegistry();
  }

  static getConsumableCandidates() {
    const routes = this.getMaterialRoutesIndex();
    const items = this.getSnapshot().orders;
    const unique = {};
    items.forEach(item => {
      if (item.manualRequest) return;
      const key = this.normalizeArticle(item.article) || this.normalizeName(item.name);
      if (item.supplier || !key) return;
      const route = this.getMaterialRoute(item, routes);
      if (!unique[key]) unique[key] = { name: item.name, article: item.article || "", count: 0, consumable: Boolean(route.consumable), plate: Boolean(route.plate) };
      unique[key].count++;
    });
    return { candidates: Object.keys(unique).map(key => unique[key]).sort((a, b) =>
      Number(Boolean(a.consumable)) - Number(Boolean(b.consumable)) || a.name.localeCompare(b.name, "ru")),
      consumables: this.getConsumables() };
  }

  static saveConsumables(names) {
    if (!Array.isArray(names) || !names.length) throw new Error("Отметьте хотя бы один расходник.");
    const sheet = this.ensureConsumablesSheet();
    const current = this.getConsumablesIndex();
    const now = new Date();
    const rows = names.map(name => String(name || "").trim()).filter(Boolean).filter(name => !current[this.normalizeName(name)])
      .map(name => [name, "", this.normalizeName(name), "", "Да", "", now]);
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    const supply = this.ensureSheet();
    const values = supply.getDataRange().getValues();
    const selected = new Set(names.map(name => this.normalizeName(name)));
    values.slice(1).forEach((row, index) => {
      if (String(row[26] || "") === "Ручная заявка") return;
      if (!selected.has(this.normalizeName(row[5]))) return;
      const rowNumber = index + 2;
      supply.getRange(rowNumber, 9).setValue("Расходники");
      supply.getRange(rowNumber, 10).clearContent();
      supply.getRange(rowNumber, 15).setValue("Расходник");
      supply.getRange(rowNumber, 17).clearContent();
      supply.getRange(rowNumber, 18).setValue("Расходники");
      supply.getRange(rowNumber, 14).setValue(now);
    });
    return { added: rows.length, snapshot: this.getSnapshot() };
  }

  static saveMaterialRoutes(routes) {
    const sheet = this.ensureConsumablesSheet(); const now = new Date(); const existing = this.getConsumables(); const byKey = {};
    existing.forEach((item, index) => {
      if (item.normalizedArticle) byKey["article:" + item.normalizedArticle] = index + 2;
      if (item.normalizedName) byKey["name:" + item.normalizedName] = index + 2;
    });
    (routes || []).forEach(route => {
      const name = String(route.name || "").trim(); const article = String(route.article || "").trim();
      const normalizedName = this.normalizeName(name); const normalizedArticle = this.normalizeArticle(article);
      if (!normalizedName && !normalizedArticle) return;
      const key = normalizedArticle ? "article:" + normalizedArticle : "name:" + normalizedName;
      const row = [name, article, normalizedName, normalizedArticle, route.consumable ? "Да" : "", route.plate ? "Да" : "", now];
      if (byKey[key]) sheet.getRange(byKey[key], 1, 1, row.length).setValues([row]);
      else { sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]); byKey[key] = sheet.getLastRow(); }
    });
    return { added: routes.length, snapshot: this.getSnapshot() };
  }

  static rememberManualConsumableRoutes(materials) {
    const rows = (materials || []).filter(material => {
      const article = this.normalizeArticle(material.article);
      const name = this.normalizeName(material.name);
      return material.consumableManual && (article || name);
    });
    if (!rows.length) return;
    const sheet = this.ensureConsumablesSheet();
    const existing = this.getConsumables();
    const byArticle = {};
    const byName = {};
    existing.forEach((item, index) => {
      if (item.normalizedArticle) byArticle[item.normalizedArticle] = index + 2;
      if (item.normalizedName) byName[item.normalizedName] = index + 2;
    });
    const now = new Date();
    rows.forEach(material => {
      const article = String(material.article || "").trim();
      const name = String(material.name || "").trim();
      const normalizedArticle = this.normalizeArticle(article);
      const normalizedName = this.normalizeName(name);
      const rowNumber = normalizedArticle ? byArticle[normalizedArticle] : byName[normalizedName];
      const record = [[name, article, normalizedName, normalizedArticle, "Да", "", now]];
      if (rowNumber) sheet.getRange(rowNumber, 1, 1, record[0].length).setValues(record);
      else {
        const next = sheet.getLastRow() + 1;
        sheet.getRange(next, 1, 1, record[0].length).setValues(record);
        if (normalizedArticle) byArticle[normalizedArticle] = next;
        if (normalizedName) byName[normalizedName] = next;
      }
    });
  }

  static searchUnmatchedArticles(supplierName) {
    const sheet = this.ensureSheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const values = sheet.getDataRange().getValues();
      const selectedSupplier = String(supplierName || "").trim();
      const supplierRegistry = this.getSupplierRegistry();
      const searchSupplier = this.getSearchSupplier(selectedSupplier, supplierRegistry);
      if (values.length < 2) return { checked: 0, matched: 0, supplier: selectedSupplier };
      const searchRegistry = searchSupplier ? [searchSupplier] : supplierRegistry;
      const priceIndex = this.buildPriceIndex(searchRegistry);
      const supplierAliases = this.getSupplierAliasIndex(searchRegistry);
      const materialMappings = this.getMaterialMappingsIndex();
      const itktMatches = {};
      let checked = 0;
      let matched = 0;
      const now = new Date();

      values.slice(1).forEach(row => {
        if (String(row[26] || "") === "Ручная заявка") return;
        const supplier = String(row[8] || "").trim();
        const isConsumable = row[17] === "Расходники" || supplier === "Расходники";
        // Ручной выбор всегда имеет приоритет над автоматическим поиском.
        if (supplier || row[11] === true || isConsumable) return;
        checked++;
        const material = {
          article: String(row[4] || "").trim(), name: String(row[5] || "").trim(),
          unit: String(row[7] || "").trim(), manufacturer: String(row[15] || "").trim()
        };
        const known = this.matchKnownMaterial(material, materialMappings);
        const knownMatch = !searchSupplier || (known && this.sameSupplier(known.supplier, searchSupplier.name)) ? known : null;
        const mappedSupplier = this.matchSupplierByManufacturer(material, supplierAliases);
        const siteMatch = searchSupplier || knownMatch || mappedSupplier ? null : (this.isBlumMaterial(material) ? (this.matchItktBlum(material, itktMatches) || this.matchBlumRus(material)) :
          (this.isDuslarMaterial(material) ? this.matchDuslar(material) :
            (this.isSheetMaterial(material) ? this.matchExpoTorg(material) : null)));
        const configuredSiteMatch = knownMatch || mappedSupplier || siteMatch ? null :
          this.matchConfiguredSite(material, searchRegistry, itktMatches);
        const match = knownMatch || mappedSupplier ? { supplier: "", price: "", method: "", productUrl: "" } :
          (siteMatch || configuredSiteMatch || this.matchSupplier(material, priceIndex));
        const foundSupplier = knownMatch ? knownMatch.supplier : (mappedSupplier || match.supplier);
        if (!foundSupplier) return;
        row[8] = foundSupplier;
        if (Number.isFinite(knownMatch && knownMatch.price)) row[9] = knownMatch.price;
        else if (Number.isFinite(match.price)) row[9] = match.price;
        row[13] = now;
        row[14] = knownMatch ? knownMatch.method : (mappedSupplier ? "По названию поставщика" : match.method);
        if (knownMatch && knownMatch.productUrl) row[16] = knownMatch.productUrl;
        else if (match.productUrl) row[16] = match.productUrl;
        matched++;
      });
      if (matched) sheet.getRange(2, 1, values.length - 1, this.HEADERS.length).setValues(values.slice(1));
      return { checked: checked, matched: matched, supplier: selectedSupplier };
    } finally {
      lock.releaseLock();
    }
  }

  static getConsumables() {
    const sheet = this.ensureConsumablesSheet();
    return sheet.getDataRange().getDisplayValues().slice(1).filter(row => row[0] || row[1]).map(row => ({
      name: row[0], article: row[1], normalizedName: row[2] || this.normalizeName(row[0]),
      normalizedArticle: row[3] || this.normalizeArticle(row[1]), consumable: row[4] !== "",
      plate: row[5] === "Да" || row[5] === "true"
    }));
  }

  static getConsumablesIndex() {
    return this.getConsumables().reduce((index, item) => {
      if (!item.consumable) return index;
      if (item.normalizedArticle) index[item.normalizedArticle] = true;
      if (item.normalizedName) index[item.normalizedName] = true;
      return index;
    }, {});
  }

  static getPlateRoutesIndex() {
    return this.getConsumables().reduce((index, item) => {
      if (!item.plate) return index;
      if (item.normalizedArticle) index[item.normalizedArticle] = true;
      if (item.normalizedName) index[item.normalizedName] = true;
      return index;
    }, {});
  }
  static getMaterialRoutesIndex() {
    return this.getConsumables().reduce((index, item) => {
      const route = { consumable: item.consumable !== false, plate: Boolean(item.plate) };
      if (item.normalizedArticle) index["article:" + item.normalizedArticle] = route;
      if (item.normalizedName) index["name:" + item.normalizedName] = route;
      return index;
    }, {});
  }
  static getMaterialRoute(material, index) {
    const routes = index || this.getMaterialRoutesIndex();
    const article = this.normalizeArticle(material && material.article);
    const name = this.normalizeName(material && material.name);
    return (article && routes["article:" + article]) || (name && routes["name:" + name]) || {};
  }
  static isEdgeMaterial(name) { return /кромк/i.test(String(name || "")); }

  static ensureMaterialMappingsSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.MATERIAL_MAPPINGS);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.MATERIAL_MAPPINGS);
    const headers = ["Артикул БАЗИС", "Нормализованный артикул", "Последнее наименование", "Поставщик", "Цена", "Способ сопоставления", "Ссылка поставщика", "Обновлено"];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
      sheet.getRange(1, 1, 1, headers.length).createFilter();
    }
    return sheet;
  }

  static getMaterialMappingsIndex() {
    const sheet = this.ensureMaterialMappingsSheet();
    return sheet.getDataRange().getValues().slice(1).reduce((index, row) => {
      const key = String(row[1] || this.normalizeArticle(row[0]) || "");
      if (key && row[3]) index[key] = {
        supplier: String(row[3]), price: Number.isFinite(row[4]) ? row[4] : "",
        method: "Справочник сопоставлений" + (row[5] ? " · " + row[5] : ""), productUrl: String(row[6] || "")
      };
      return index;
    }, {});
  }

  static matchKnownMaterial(material, index) {
    const key = this.normalizeArticle(material.article);
    return key && index[key] ? index[key] : null;
  }

  static rememberConfirmedXmlMaterialMappings(materials) {
    const rows = (materials || []).filter(material => {
      const article = this.normalizeArticle(material.article);
      return article && material.supplier && material.supplier !== "Расходники" && material.category !== "Расходники";
    });
    if (!rows.length) return;
    const sheet = this.ensureMaterialMappingsSheet();
    const existing = sheet.getDataRange().getValues();
    const byKey = {};
    existing.slice(1).forEach((row, index) => { if (row[1]) byKey[String(row[1])] = index + 2; });
    const now = new Date();
    rows.forEach(material => {
      const key = this.normalizeArticle(material.article);
      const record = [[String(material.article || "").trim(), key, String(material.name || "").trim(), String(material.supplier || "").trim(),
        Number.isFinite(material.price) ? material.price : "", "XML: подтверждено вручную", String(material.productUrl || "").trim(), now]];
      if (byKey[key]) sheet.getRange(byKey[key], 1, 1, record[0].length).setValues(record);
      else {
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, record[0].length).setValues(record);
        byKey[key] = sheet.getLastRow();
      }
    });
  }

  static ensureConsumablesSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CONSUMABLES);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.CONSUMABLES);
    const headers = ["Наименование", "Артикул БАЗИС", "Нормализованное наименование", "Нормализованный артикул", "Расходник", "Плитный материал", "Добавлено"];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
      sheet.getRange(1, 1, 1, headers.length).createFilter();
    } else if (sheet.getLastColumn() < headers.length || sheet.getRange(1, 2).getDisplayValue() !== "Артикул БАЗИС") {
      const rows = sheet.getDataRange().getValues().slice(1);
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows.map(row => [row[0], "", row[1] || this.normalizeName(row[0]), "", row[2], row[3], row[4]]));
    }
    return sheet;
  }

  static ensureSupplierRegistrySheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.SUPPLIERS);
    const headers = ["Поставщик", "Источник поиска", "Активен", "Примечание", "Названия в спецификации", "Шаблон поиска сайта"];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
      sheet.getRange("C2:C").insertCheckboxes();
      sheet.getRange(1, 1, 1, headers.length).createFilter();
    } else if (sheet.getLastColumn() < headers.length || sheet.getRange(1, 2).getDisplayValue() !== "Источник поиска" || sheet.getRange(1, headers.length).getDisplayValue() !== "Шаблон поиска сайта") {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
    }
    return sheet;
  }

  static getSupplierAliasIndex(suppliers) {
    const index = {};
    (suppliers || this.getSupplierRegistry()).forEach(item => {
      [item.name].concat(String(item.aliases || "").split(/[;,\n]/)).forEach(alias => {
        const key = this.normalize(alias);
        if (key) index[key] = item.name;
      });
    });
    return index;
  }

  static matchSupplierByManufacturer(material, aliases) {
    const source = this.normalize(material.manufacturer);
    return source && aliases[source] ? aliases[source] : "";
  }

  static sameSupplier(left, right) {
    return this.normalize(left) === this.normalize(right);
  }

  static getSearchSupplier(supplierName, registry) {
    if (!supplierName) return null;
    if (supplierName === "Расходники" || supplierName === "__supply_other__") {
      throw new Error("Для поиска выберите конкретного поставщика или фильтр «Все».");
    }
    const supplier = (registry || []).find(item => item.active && this.sameSupplier(item.name, supplierName));
    if (!supplier) throw new Error("Выбранный поставщик не найден среди активных поставщиков.");
    const source = String(supplier.priceUrl || "").trim();
    if (!source) {
      throw new Error("У поставщика «" + supplier.name + "» нет активного источника поиска. Укажите прайс, папку прайсов или сайт.");
    }
    if (!this.isWebsiteSourceUrl(source)) return supplier;
    const configured = this.ensureSupplierSearchTemplate(supplier);
    if (!configured.searchTemplate) {
      throw new Error("Для сайта поставщика «" + supplier.name + "» не удалось определить поиск. Укажите шаблон URL с {query} в справочнике поставщиков.");
    }
    return configured;
  }

  static toNativeSpreadsheet(file) {
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) return file.getId();
    return Drive.Files.copy({ title: "TEMP_PRICE_" + Date.now() }, file.getId(), { convert: true }).id;
  }

  static extractPriceEntries(workbook, supplier) {
    const result = [];
    workbook.getSheets().forEach(sheet => {
      const values = sheet.getDataRange().getDisplayValues();
      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        const normalized = row.map(value => this.normalize(value));
        const article = normalized.findIndex(value => ["артикул", "код", "код товара"].includes(value));
        const name = normalized.findIndex(value => ["наименование", "название", "материал"].includes(value));
        const price = normalized.findIndex(value => ["цена", "розничная цена", "оптовая цена", "стоимость"].includes(value));
        if (name < 0 || (article < 0 && price < 0)) continue;
        for (let i = r + 1; i < values.length; i++) {
          if (!values[i][name]) continue;
          result.push({ supplier: supplier, article: article >= 0 ? values[i][article] : "", name: values[i][name], price: price >= 0 ? this.toNumber(values[i][price]) : "" });
        }
        break;
      }
    });
    return result;
  }

  static matchSupplier(material, index) {
    const article = this.normalizeArticle(material.article);
    if (article && index.articles[article]) return this.matchResult(index.articles[article], "Артикул");
    const name = this.normalizeName(material.name);
    if (name && index.names[name]) return this.matchResult(index.names[name], "Наименование");
    const sourceTokens = new Set(this.materialTokens(material.name));
    let best = null;
    index.entries.forEach(entry => {
      const targetTokens = this.materialTokens(entry.name);
      if (!targetTokens.length) return;
      const common = targetTokens.filter(token => sourceTokens.has(token)).length;
      const score = common / Math.max(sourceTokens.size, targetTokens.length);
      if (common >= 2 && (!best || score > best.score)) best = { entry: entry, score: score };
    });
    return best && best.score >= 0.65 ? this.matchResult(best.entry, "По смыслу") : { supplier: "", price: "", method: "" };
  }

  static matchConfiguredSite(material, suppliers, itktMatches) {
    const sources = (suppliers || []).filter(item => item.active && this.isWebsiteSourceUrl(item.priceUrl));
    for (let i = 0; i < sources.length; i++) {
      const source = this.ensureSupplierSearchTemplate(sources[i]);
      let result = this.matchGenericSupplierSite(material, source);
      if (!result) {
        if (/itkt\.ru/i.test(source.priceUrl)) result = this.matchItktBlum(material, itktMatches);
        else if (/duslar\.biz/i.test(source.priceUrl)) result = this.matchDuslar(material);
        else if (/expo-torg\.ru/i.test(source.priceUrl)) result = this.matchExpoTorg(material);
        else if (/blum-rus\.ru/i.test(source.priceUrl)) result = this.matchBlumRus(material);
      }
      if (result) return Object.assign({}, result, { supplier: source.name, method: source.name + " · " + result.method });
    }
    return null;
  }

  static isWebsiteSourceUrl(url) {
    const value = String(url || "").trim();
    return /^https?:\/\//i.test(value) && !this.extractFolderId(value) && !this.extractDriveFileId(value);
  }

  static ensureSupplierSearchTemplate(supplier) {
    if (!supplier || !this.isWebsiteSourceUrl(supplier.priceUrl) || supplier.searchTemplate) return supplier;
    const searchTemplate = this.discoverSearchTemplate(supplier.priceUrl);
    if (!searchTemplate) return supplier;
    this.saveSupplierSearchTemplate(supplier.name, searchTemplate);
    return Object.assign({}, supplier, { searchTemplate: searchTemplate });
  }

  static saveSupplierSearchTemplate(name, searchTemplate) {
    const sheet = this.ensureSupplierRegistrySheet();
    const names = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getDisplayValues();
    const row = names.findIndex(item => this.sameSupplier(item[0], name));
    if (row >= 0) sheet.getRange(row + 2, 6).setValue(searchTemplate);
  }

  static discoverSearchTemplate(sourceUrl) {
    try {
      const response = UrlFetchApp.fetch(sourceUrl, { muteHttpExceptions: true, followRedirects: true });
      if (response.getResponseCode() !== 200) return "";
      const html = response.getContentText();
      const links = String(html).match(/<link\b[^>]*>/gi) || [];
      for (let i = 0; i < links.length; i++) {
        const rel = this.htmlAttribute(links[i], "rel");
        const href = this.htmlAttribute(links[i], "href");
        if (!/\bsearch\b/i.test(rel) || !href) continue;
        const descriptor = UrlFetchApp.fetch(this.toAbsoluteUrl(href, sourceUrl), { muteHttpExceptions: true, followRedirects: true });
        if (descriptor.getResponseCode() !== 200) continue;
        const template = String(descriptor.getContentText()).match(/<Url\b[^>]*\btemplate=["']([^"']+)["']/i);
        const normalized = template && this.normalizeSearchTemplate(template[1], sourceUrl);
        if (normalized) return normalized;
      }
      const forms = String(html).match(/<form\b[^>]*>[\s\S]*?<\/form>/gi) || [];
      for (let i = 0; i < forms.length; i++) {
        const input = forms[i].match(/<input\b[^>]*\bname=["'](q|query|search|keyword)["'][^>]*>/i);
        if (!input) continue;
        const action = this.htmlAttribute(forms[i].match(/<form\b[^>]*>/i)[0], "action") || sourceUrl;
        const base = this.toAbsoluteUrl(action, sourceUrl);
        const separator = base.indexOf("?") >= 0 ? "&" : "?";
        return this.normalizeSearchTemplate(base + separator + input[1] + "={query}", sourceUrl);
      }
    } catch (error) {
      return "";
    }
    return "";
  }

  static normalizeSearchTemplate(template, sourceUrl) {
    const value = this.toAbsoluteUrl(String(template || "").replace(/&amp;/g, "&"), sourceUrl)
      .replace(/\{searchTerms(?:\?[^}]*)?\}/gi, "{query}");
    return /^https?:\/\//i.test(value) && value.indexOf("{query}") >= 0 ? value : "";
  }

  static htmlAttribute(tag, name) {
    const match = String(tag || "").match(new RegExp("\\b" + name + "\\s*=\\s*([\"'])(.*?)\\1", "i"));
    return match ? String(match[2] || "").replace(/&amp;/g, "&") : "";
  }

  static toAbsoluteUrl(url, baseUrl) {
    const value = String(url || "").trim();
    if (/^https?:\/\//i.test(value)) return value;
    const origin = String(baseUrl || "").match(/^https?:\/\/[^/]+/i);
    if (!origin) return value;
    if (value.indexOf("//") === 0) return "https:" + value;
    if (value.indexOf("/") === 0) return origin[0] + value;
    return origin[0] + "/" + value.replace(/^\.\//, "");
  }

  static matchGenericSupplierSite(material, supplier) {
    if (!supplier || !supplier.searchTemplate) return null;
    const query = String(material.article || material.name || "").trim();
    if (!query) return null;
    const cache = CacheService.getScriptCache();
    const key = this.normalizeArticle(material.article) || this.normalizeName(material.name);
    const cacheKey = "supplier-site:" + this.compact(supplier.name).slice(0, 36) + ":" + key;
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    try {
      const url = supplier.searchTemplate.replace(/\{query\}/g, encodeURIComponent(query));
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      if (response.getResponseCode() !== 200) return null;
      const result = this.parseGenericSupplierSearch(response.getContentText(), material, url);
      if (result) cache.put(cacheKey, JSON.stringify(result), 21600);
      return result;
    } catch (error) {
      return null;
    }
  }

  static parseGenericSupplierSearch(html, material, searchUrl) {
    const expectedArticle = this.normalizeArticle(material.article);
    const sourceTokens = new Set(this.materialTokens(material.name));
    const products = this.extractGenericSupplierProducts(html, searchUrl).map(product => {
      const article = this.normalizeArticle(product.article);
      const targetTokens = this.materialTokens(product.name);
      const common = targetTokens.filter(token => sourceTokens.has(token)).length;
      return Object.assign({}, product, {
        articleMatch: Boolean(expectedArticle && article === expectedArticle),
        common: common, score: common / Math.max(sourceTokens.size || 1, targetTokens.length || 1)
      });
    });
    const exact = products.filter(item => item.articleMatch);
    if (exact.length === 1) return { price: "", method: "артикул", productUrl: exact[0].productUrl };
    if (expectedArticle || !sourceTokens.size) return null;
    const suitable = products.filter(item => item.common >= 2 && item.score >= 0.65).sort((left, right) => right.score - left.score || right.common - left.common);
    if (!suitable.length || (suitable[1] && suitable[0].score - suitable[1].score < 0.08)) return null;
    return { price: "", method: "наименование", productUrl: suitable[0].productUrl };
  }

  static extractGenericSupplierProducts(html, searchUrl) {
    const products = [];
    const push = item => {
      if (!item || !item.name || !item.productUrl) return;
      const duplicate = products.some(value => value.productUrl === item.productUrl && this.normalizeArticle(value.article) === this.normalizeArticle(item.article));
      if (!duplicate) products.push(item);
    };
    const scripts = String(html || "").match(/<script\b[^>]*type=["']application\/ld\+json["'][^>]*>[\s\S]*?<\/script>/gi) || [];
    scripts.forEach(script => {
      const source = String(script).replace(/^<script[^>]*>|<\/script>$/gi, "").trim();
      try { this.collectJsonLdProducts(JSON.parse(source), searchUrl, push); } catch (error) { /* Некорректный JSON-LD пропускаем. */ }
    });
    const cards = String(html || "").matchAll(/article_block[^>]*data-value=["']([^"']+)["'][\s\S]{0,2400}?item-title[\s\S]{0,600}?<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi);
    for (const card of cards) push({ article: this.stripHtml(card[1]), name: this.stripHtml(card[3]), productUrl: this.toAbsoluteUrl(card[2], searchUrl) });
    return products;
  }

  static collectJsonLdProducts(value, searchUrl, push) {
    if (Array.isArray(value)) return value.forEach(item => this.collectJsonLdProducts(item, searchUrl, push));
    if (!value || typeof value !== "object") return;
    const type = Array.isArray(value["@type"]) ? value["@type"].join(" ") : String(value["@type"] || "");
    if (/\bProduct\b/i.test(type) && value.name) {
      const identifier = typeof value.identifier === "object" ? (value.identifier.value || value.identifier["@value"] || "") : value.identifier;
      push({ article: value.sku || value.mpn || value.productID || identifier || "", name: value.name, productUrl: this.toAbsoluteUrl(value.url || value["@id"] || "", searchUrl) });
    }
    if (value["@graph"]) this.collectJsonLdProducts(value["@graph"], searchUrl, push);
    if (value.itemListElement) this.collectJsonLdProducts(value.itemListElement, searchUrl, push);
    if (value.item) this.collectJsonLdProducts(value.item, searchUrl, push);
  }

  static isBlumMaterial(material) {
    return /\bblum\b|\bблюм\b/i.test(String(material.manufacturer || ""));
  }

  static isDuslarMaterial(material) {
    return /\bduslar\b|\bдуслар\b/i.test(String(material.manufacturer || ""));
  }

  static isSheetMaterial(material) {
    return /плитн(ый|ые)? материал|экспо[\s-]?торг|expo/i.test(String(material.manufacturer || ""));
  }

  static matchExpoTorg(material) {
    const cache = CacheService.getScriptCache();
    const key = this.normalizeArticle(material.article) || this.normalizeName(material.name);
    if (!key) return null;
    const cacheKey = "expo:" + key;
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    try {
      const query = String(material.article || material.name || "").trim();
      const url = "https://expo-torg.ru/search/index.php?q=" + encodeURIComponent(query);
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      if (response.getResponseCode() !== 200) return null;
      const result = this.parseExpoTorgSearch(response.getContentText(), material);
      if (result) cache.put(cacheKey, JSON.stringify(result), 21600);
      return result;
    } catch (error) {
      return null;
    }
  }

  static parseExpoTorgSearch(html, material) {
    const article = this.normalizeArticle(material.article);
    const sourceTokens = new Set(this.materialTokens(material.name));
    const blocks = String(html || "").split(/<div class="search_element"/i).slice(1);
    const candidates = blocks.map(block => {
      const match = block.match(/<li><a href="([^"]+)">([\s\S]*?)<\/a><\/li>/i);
      if (!match) return null;
      const name = this.stripHtml(match[2]);
      const targetTokens = this.materialTokens(name);
      const common = targetTokens.filter(token => sourceTokens.has(token)).length;
      const score = common / Math.max(sourceTokens.size, targetTokens.length);
      return { name: name, productUrl: "https://expo-torg.ru" + match[1], articleMatch: article && this.normalizeArticle(name).includes(article), common: common, score: score };
    }).filter(Boolean);
    if (!candidates.length) return null;
    const suitable = candidates.filter(item => (item.articleMatch && item.common >= 1) || (item.common >= 2 && item.score >= 0.65))
      .sort((a, b) => (Number(b.articleMatch) - Number(a.articleMatch)) || b.score - a.score || b.common - a.common);
    if (!suitable.length) return null;
    const best = suitable[0];
    const runnerUp = suitable[1];
    if (runnerUp && best.articleMatch === runnerUp.articleMatch && best.score - runnerUp.score < 0.08) return null;
    return {
      supplier: "Экспо-Торг", price: "",
      method: best.articleMatch ? "Экспо-Торг · артикул" : "Экспо-Торг · по наименованию",
      productUrl: best.productUrl
    };
  }

  static matchDuslar(material) {
    const article = String(material.article || "").trim();
    if (!article) return null;
    const key = this.normalizeArticle(article);
    const cache = CacheService.getScriptCache();
    const cacheKey = "duslar:" + key;
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    try {
      const url = "https://www.duslar.biz/search/?q=" + encodeURIComponent(article);
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      if (response.getResponseCode() !== 200) return null;
      const result = this.parseDuslarSearch(response.getContentText(), article);
      if (result) cache.put(cacheKey, JSON.stringify(result), 21600);
      return result;
    } catch (error) {
      return null;
    }
  }

  static parseDuslarSearch(html, article) {
    const expected = this.normalizeArticle(article);
    const blocks = String(html || "").split(/<div class="item product\b/i).slice(1);
    const matches = blocks.map(block => {
      const compact = this.normalizeArticle(block);
      if (!compact.includes(expected)) return null;
      const nameMatch = block.match(/<a href="([^"]+)" class="name">[\s\S]*?class="middle">([\s\S]*?)<\/span>/i);
      const priceMatch = block.match(/<a class="price">\s*([^<]+)/i);
      if (!nameMatch) return null;
      const price = priceMatch ? this.toNumber(priceMatch[1]) : "";
      return {
        supplier: "Duslar", price: Number.isFinite(price) ? price : "",
        method: "Duslar · точный артикул", productUrl: "https://www.duslar.biz" + nameMatch[1]
      };
    }).filter(Boolean);
    return matches.length === 1 ? matches[0] : null;
  }

  static matchItktBlum(material, memo) {
    const article = String(material.article || "").trim();
    if (!article) return null;
    const key = this.normalizeArticle(article);
    if (memo[key] !== undefined) return memo[key];
    const cache = CacheService.getScriptCache();
    const cacheKey = "itkt:" + key;
    const cached = cache.get(cacheKey);
    if (cached) return memo[key] = JSON.parse(cached);
    try {
      const queries = [...new Set([article, this.toLatinArticle(article)])];
      for (let i = 0; i < queries.length; i++) {
        const url = "https://www.itkt.ru/category/search/?q=" + encodeURIComponent(queries[i]);
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
        if (response.getResponseCode() !== 200) continue;
        const result = this.parseItktSearch(response.getContentText(), article, material.name, false);
        if (result) {
          cache.put(cacheKey, JSON.stringify(result), 21600);
          return memo[key] = result;
        }
      }
      // Название используем только после поиска по артикулу: это снижает
      // число обращений и не подменяет конкретную модификацию аналогом.
      const name = String(material.name || "").trim();
      if (name.length >= 5) {
        const response = UrlFetchApp.fetch("https://www.itkt.ru/category/search/?q=" + encodeURIComponent(name), { muteHttpExceptions: true, followRedirects: true });
        if (response.getResponseCode() === 200) {
          const result = this.parseItktSearch(response.getContentText(), article, name, true);
          if (result) {
            cache.put(cacheKey, JSON.stringify(result), 21600);
            return memo[key] = result;
          }
        }
      }
      return memo[key] = null;
    } catch (error) {
      // Недоступность внешнего сайта не должна останавливать импорт XLS.
      return memo[key] = null;
    }
  }

  static matchBlumRus(material) {
    const article = String(material.article || "").trim();
    const name = String(material.name || "").trim();
    if (!article && name.length < 5) return null;
    const key = this.normalizeArticle(article) || this.normalizeName(name);
    const cache = CacheService.getScriptCache();
    const cacheKey = "blumrus:" + key;
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    try {
      const kitQuery = this.makeBlumKitQuery(name);
      const queries = [...new Set([article, name, kitQuery].filter(value => String(value || "").trim()))];
      for (let i = 0; i < queries.length; i++) {
        const url = "https://blum-rus.ru/index.php?route=product/search&search=" + encodeURIComponent(queries[i]);
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
        if (response.getResponseCode() !== 200) continue;
        const result = this.parseBlumRusSearch(response.getContentText(), material);
        if (result) {
          cache.put(cacheKey, JSON.stringify(result), 21600);
          return result;
        }
      }
      return null;
    } catch (error) {
      return null;
    }
  }

  static makeBlumKitQuery(name) {
    const text = String(name || "");
    if (!/tandem/i.test(text) || !/tip[\s-]?on/i.test(text)) return "";
    const length = text.match(/\b(250|270|300|350|400|450|500|550|600)\s*мм?\b/i);
    const load = text.match(/\b(30|40|50|60)\s*кг\b/i);
    return ["TANDEM", "TIP-ON", load && load[1] + " кг", length && length[1] + " мм"].filter(Boolean).join(" ");
  }

  static parseBlumRusSearch(html, material) {
    const article = this.normalizeArticle(material.article);
    const sourceTokens = new Set(this.materialTokens(material.name));
    const isKit = /комплект|набор|tandem[\s\S]{0,40}tip[\s-]?on|tip[\s-]?on[\s\S]{0,40}tandem/i.test(String(material.name || ""));
    const blocks = String(html || "").split(/<div[^>]+class="[^"]*(?:product-thumb|product-layout|product-grid)[^"]*"/i).slice(1);
    const candidates = blocks.map(block => {
      const nameMatch = block.match(/<(?:h4|div)[^>]+class="[^"]*(?:name|caption)[^"]*"[^>]*>[\s\S]*?<a[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i) ||
        block.match(/<a[^>]+href="([^"]+)"[^>]*>([^<]*(?:TANDEM|Blum|TIP-ON)[^<]*)<\/a>/i);
      if (!nameMatch) return null;
      const candidateName = this.stripHtml(nameMatch[2]);
      const targetTokens = this.materialTokens(candidateName);
      const common = targetTokens.filter(token => sourceTokens.has(token)).length;
      const score = common / Math.max(sourceTokens.size, targetTokens.length);
      const priceMatch = block.match(/class="[^"]*price[^"]*"[^>]*>[\s\S]*?([\d\s]+[,.]\d{1,2})\s*(?:р|руб)/i);
      return {
        name: candidateName, productUrl: nameMatch[1].indexOf("http") === 0 ? nameMatch[1] : "https://blum-rus.ru/" + nameMatch[1].replace(/^\//, ""),
        price: priceMatch ? this.toNumber(priceMatch[1]) : "", common: common, score: score,
        articleMatch: article && this.normalizeArticle(block).includes(article)
      };
    }).filter(Boolean);
    const suitable = candidates.filter(item => item.articleMatch || (item.common >= (isKit ? 3 : 2) && item.score >= (isKit ? 0.45 : 0.6)))
      .sort((a, b) => Number(b.articleMatch) - Number(a.articleMatch) || b.score - a.score || b.common - a.common);
    if (!suitable.length) return null;
    const best = suitable[0];
    const next = suitable[1];
    if (next && best.articleMatch === next.articleMatch && best.score - next.score < 0.08) return null;
    const wholeKit = /комплект|набор/i.test(best.name);
    return {
      supplier: "Blum-rus", price: wholeKit || !isKit ? (Number.isFinite(best.price) ? best.price : "") : "",
      method: wholeKit ? "Blum-rus · комплект" : (isKit ? "Blum-rus · часть комплекта — состав проверить" : "Blum-rus · артикул/наименование"),
      productUrl: best.productUrl
    };
  }

  static parseItktSearch(html, article, name, allowSemantic) {
    const expected = this.normalizeArticle(article);
    const blocks = String(html || "").split(/<div class="catalog-product\b/i).slice(1);
    const products = blocks.map(block => {
      const articleMatch = block.match(/class="artnumber">\s*Арт\.\s*([^<]+?)\s*<\/div>/i);
      const nameMatch = block.match(/class="name">[\s\S]*?<a[^>]*>([\s\S]*?)<\/a>/i);
      const priceMatch = block.match(/class="price[^>]*>[\s\S]*?<span[^>]*>\s*([^<]+?)\s*<\/span>/i);
      const linkMatch = block.match(/<a href="([^"]+)"[\s\S]*?detail_link/i);
      const price = priceMatch ? this.toNumber(priceMatch[1]) : "";
      return {
        article: articleMatch ? this.normalizeArticle(articleMatch[1]) : "",
        name: nameMatch ? this.stripHtml(nameMatch[1]) : "",
        supplier: "Интеркомплект (ITKT)", price: Number.isFinite(price) ? price : "",
        productUrl: linkMatch ? "https://www.itkt.ru" + linkMatch[1] : ""
      };
    }).filter(product => product.article || product.name);
    const exact = products.find(product => product.article === expected);
    if (exact) return this.itktMatchResult(exact, "ITKT · точный артикул");

    const partial = products.filter(product => product.article && expected &&
      (product.article.includes(expected) || expected.includes(product.article)));
    if (partial.length === 1) return this.itktMatchResult(partial[0], "ITKT · частичный артикул");

    if (allowSemantic) {
      const sourceTokens = new Set(this.materialTokens(name));
      const scored = products.map(product => {
        const targetTokens = this.materialTokens(product.name);
        const common = targetTokens.filter(token => sourceTokens.has(token)).length;
        return { product: product, score: common / Math.max(sourceTokens.size, targetTokens.length), common: common };
      }).filter(item => item.common >= 2 && item.score >= 0.7)
        .sort((a, b) => b.score - a.score || b.common - a.common);
      if (scored.length === 1 || (scored[0] && scored[0].score - scored[1].score >= 0.15)) {
        return this.itktMatchResult(scored[0].product, "ITKT · по наименованию");
      }
    }
    return null;
  }

  static itktMatchResult(product, method) {
    return { supplier: product.supplier, price: product.price, method: method, productUrl: product.productUrl };
  }

  static stripHtml(value) {
    return String(value || "").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/\s+/g, " ").trim();
  }

  static materialTokens(value) {
    const synonyms = { "саморезы": "саморез", "винты": "винт", "петли": "петля", "направляющие": "направляющая", "ручки": "ручка", "ножки": "ножка", "полкодержатели": "полкодержатель", "заглушки": "заглушка" };
    return this.tokens(value).map(token => synonyms[token] || token).filter(token => token.length > 1 && !["для", "под", "комплект", "мебельный", "мебельная", "цвет"].includes(token));
  }

  static matchResult(entry, method) {
    return { supplier: entry.supplier, price: Number.isFinite(entry.price) ? entry.price : "", method: method };
  }

  static changeStatus(rowId, direction) {
    const sheet = this.ensureSheet();
    const row = this.findRow(sheet, rowId);
    const current = String(sheet.getRange(row, 11).getValue() || "Новый");
    const index = Math.max(0, this.STATUSES.indexOf(current));
    const next = Math.max(0, Math.min(this.STATUSES.length - 1, index + (direction === "back" ? -1 : 1)));
    sheet.getRange(row, 11).setValue(this.STATUSES[next]);
    sheet.getRange(row, 14).setValue(new Date());
    return { rowId: rowId, status: this.STATUSES[next] };
  }

  static setStatus(rowId, status) {
    const next = String(status || "").trim();
    if (!this.STATUSES.includes(next)) throw new Error("Недопустимый статус закупки.");
    const sheet = this.ensureSheet();
    const row = this.findRow(sheet, rowId);
    sheet.getRange(row, 11).setValue(next);
    sheet.getRange(row, 14).setValue(new Date());
    return { rowId: rowId, status: next };
  }

  static assignSupplier(rowId, supplier) {
    const sheet = this.ensureSheet();
    const row = this.findRow(sheet, rowId);
    const selectedSupplier = String(supplier || "").trim();
    sheet.getRange(row, 9).setValue(selectedSupplier);
    sheet.getRange(row, 12).setValue(true);
    sheet.getRange(row, 14).setValue(new Date());
    sheet.getRange(row, 15).setValue("Вручную");
    return { rowId: rowId, supplier: supplier };
  }

  /** Одноразовая миграция: удаляет ранее накопленные автоматические правила. */
  static resetMaterialMappingsToManualOnly() {
    const sheet = this.ensureMaterialMappingsSheet();
    const removed = Math.max(0, sheet.getLastRow() - 1);
    if (removed) sheet.getRange(2, 1, removed, sheet.getLastColumn()).clearContent();
    return { removed: removed };
  }

  static deleteMaterial(rowId) {
    const sheet = this.ensureSheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const row = this.findRow(sheet, rowId);
      const name = sheet.getRange(row, 6).getDisplayValue();
      sheet.deleteRow(row);
      return { rowId: rowId, name: name };
    } finally {
      lock.releaseLock();
    }
  }

  static clearOrderMaterials(orderKey) {
    const order = this.getMaterialOrderDirectory().find(item => item.key === String(orderKey || ""));
    if (!order) throw new Error("Материалы выбранного изделия уже не найдены в очереди снабжения.");
    const sheet = this.ensureSheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    let removedSupply = 0;
    try {
      const rows = sheet.getDataRange().getValues();
      const orderRows = rows.slice(1).filter(row => [row[1], row[2], row[3]].join("|") === order.key && String(row[26] || "") !== "Ручная заявка");
      const orderedIds = new Set(SupplierOrderService.getOrderedMaterialIds());
      if (orderRows.some(row => orderedIds.has(String(row[0])))) {
        throw new Error("Нельзя очистить материалы: часть позиций уже включена в заказ поставщику. Сначала исключите их из заказа поставщику.");
      }
      removedSupply = orderRows.length;
      const remaining = rows.slice(1).filter(row => [row[1], row[2], row[3]].join("|") !== order.key || String(row[26] || "") === "Ручная заявка");
      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, this.HEADERS.length).clearContent();
      if (remaining.length) sheet.getRange(2, 1, remaining.length, this.HEADERS.length).setValues(remaining);
    } finally {
      lock.releaseLock();
    }
    const plate = PlateMaterialService.clearOrderMaterials(order.key);
    return { orderKey: order.key, removedSupply: removedSupply, removedPlate: plate.removed };
  }

  static findRow(sheet, rowId) {
    const finder = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1)
      .createTextFinder(String(rowId)).matchEntireCell(true).findNext();
    if (!finder) throw new Error("Позиция снабжения не найдена.");
    return finder.getRow();
  }

  static getSetting(name) {
    const sheet = SheetService.getSheet(CONFIG.SHEETS.SETTINGS);
    if (!sheet || sheet.getLastRow() < 1) return "";
    const values = sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), 2).getDisplayValues();
    const row = values.find(item => this.normalize(item[0]) === this.normalize(name));
    return row ? String(row[1] || "").trim() : "";
  }

  static extractFolderId(url) {
    const match = String(url || "").match(/folders\/([\w-]+)/);
    return match ? match[1] : "";
  }

  static extractDriveFileId(url) {
    const value = String(url || "");
    const byPath = value.match(/\/d\/([\w-]+)/);
    const byQuery = value.match(/[?&]id=([\w-]+)/);
    return (byPath || byQuery || [])[1] || "";
  }

  static buildRowId(orderKey, material) {
    const identity = this.normalizeArticle(material.article) || (this.normalizeName(material.name) + "|" + this.normalize(material.unit));
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, orderKey + "|" + identity, Utilities.Charset.UTF_8);
    return bytes.slice(0, 12).map(value => (value + 256).toString(16).slice(-2)).join("");
  }

  static normalize(value) {
    return String(value === null || value === undefined ? "" : value).toLowerCase().replace(/ё/g, "е").replace(/[_.,;:()\[\]{}]+/g, " ").replace(/\s+/g, " ").trim();
  }

  static compact(value) { return this.normalize(value).replace(/[^a-zа-я0-9]/g, ""); }
  static normalizeArticle(value) { return this.toLatinArticle(this.compact(value)); }
  static toLatinArticle(value) {
    const similar = { "а": "a", "в": "b", "с": "c", "е": "e", "н": "h", "і": "i", "к": "k", "м": "m", "о": "o", "р": "p", "т": "t", "у": "y", "х": "x" };
    return String(value || "").toLowerCase().replace(/[авсенікмортостанух]/g, symbol => similar[symbol] || symbol);
  }
  static normalizeName(value) { return this.normalize(value).replace(/\b(шт|компл|м п|м\/п)\b/g, "").trim(); }
  static tokens(value) { return this.normalize(value).split(" ").filter(Boolean); }
  static normalizeXmlSourceQuantity(rawValue, materialName, allowEmpty) {
    const text = String(rawValue === null || rawValue === undefined ? "" : rawValue).trim();
    const numericText = text.replace(/\s/g, "").replace(",", ".");
    const value = this.toNumber(text);
    if (!text && allowEmpty) {
      return { value: 1, warning: "В позиции «" + materialName + "» исходное количество не указано; установлено 1. Проверьте строку перед формированием заказа." };
    }
    if (!text || !/^\d+(?:\.\d+)?$/.test(numericText) || !Number.isFinite(value) || value < 0) {
      throw new Error("В XML-позиции «" + materialName + "» количество должно быть положительным числом.");
    }
    if (value === 0) {
      return { value: 1, warning: "В позиции «" + materialName + "» исходное количество 0 заменено на 1." };
    }
    return { value: value, warning: "" };
  }
  static toNumber(value) {
    const text = String(value === null || value === undefined ? "" : value).replace(/\s/g, "").replace(",", ".");
    const match = text.match(/-?\d+(?:\.\d+)?/);
    const number = match ? Number(match[0]) : NaN;
    return Number.isFinite(number) ? number : NaN;
  }
}
