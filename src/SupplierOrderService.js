/**
 * Заказы поставщикам. Заказ хранит неизменяемый состав строк снабжения,
 * чтобы дальнейшее изменение спецификации не меняло уже отправленную заявку.
 */
class SupplierOrderService {
  static get ORDER_STATUSES() {
    return ["Сформирован", "Счёт получен", "Выгружен", "Оплачен", "Товар получен"];
  }

  static get MATERIAL_STATUS_BY_ORDER_STATUS() {
    return {
      "Сформирован": "Заказ поставщику",
      "Счёт получен": "Счёт получен",
      "Выгружен": "Счёт получен",
      "Оплачен": "Оплачено",
      "Товар получен": "Товар получен"
    };
  }

  static deliveryDate(value) {
    if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const date = String(value || "").trim();
    if (!date) return "";
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date.slice(0, 4) < "1900" ||
        new Date(date + "T12:00:00Z").toISOString().slice(0, 10) !== date) throw new Error("Укажите корректную дату поставки.");
    return date;
  }

  static canSetDeliveryDate(status) {
    return ["Счёт получен", "Счет получен", "Выгружен", "Оплачен", "Товар получен"].includes(status);
  }

  static saveDeliveryDates(orderId, payload) {
    const data = payload || {};
    const date = this.deliveryDate(data.date);
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const sheet = this.ensureOrdersSheet();
      const orders = sheet.getDataRange().getValues();
      const index = orders.findIndex((row, i) => i > 0 && String(row[0]) === String(orderId));
      if (index < 1) throw new Error("Заказ поставщику не найден.");
      if (!this.canSetDeliveryDate(String(orders[index][3]))) throw new Error("Дата поставки доступна после получения счёта.");
      if (data.lineIds !== undefined) {
        const ids = [...new Set((Array.isArray(data.lineIds) ? data.lineIds : []).map(String))];
        const linesSheet = this.ensureLinesSheet();
        const lines = linesSheet.getDataRange().getValues().map((row, i) => ({ row, number: i + 1 }))
          .filter(item => item.number > 1 && String(item.row[1]) === String(orderId) && ids.includes(String(item.row[0])));
        if (!ids.length || lines.length !== ids.length) throw new Error("Строки заказа изменились. Обновите состав.");
        linesSheet.getRangeList(lines.map(item => "M" + item.number)).setNumberFormat("@").setValue(date);
      } else {
        sheet.getRange(index + 1, 10).setNumberFormat("@").setValue(date);
      }
      sheet.getRange(index + 1, 6).setValue(new Date());
      return { orderId: String(orderId) };
    } finally { lock.releaseLock(); }
  }

  static articleProblem(value) {
    const article = String(value || "").replace(/\u00a0/g, " ").trim();
    if (!article) return "Не заполнен артикул";
    if (/[;\r\n\t"]/.test(article) || article.startsWith("=")) return "Артикул содержит недопустимые символы";
    return "";
  }

  static exportLineIssues(orders, lines) {
    const byId = new Map(orders.map(item => [String(item.row[0]), item.row]));
    return lines.map(row => {
      const order = byId.get(String(row[1]));
      const articleError = this.articleProblem(row[6]);
      const quantity = typeof row[8] === "number" ? row[8] : Number(String(row[8] || "").trim().replace(",", "."));
      const reasons = [articleError, !Number.isFinite(quantity) || quantity <= 0 ? "Количество должно быть положительным числом" : ""].filter(Boolean);
      return reasons.length ? { orderId: String(row[1]), orderNumber: String(order[1]), lineId: String(row[0]),
        name: String(row[7] || ""), product: [row[3], row[4], row[5]].map(value => String(value || "")).join(" · "),
        article: String(row[6] || ""), reason: reasons.join("; "), articleError: !!articleError } : null;
    }).filter(Boolean);
  }

  static setLineArticle(orderId, lineIds, article, expectedArticle) {
    const value = String(article || "").replace(/\u00a0/g, " ").trim();
    const problem = this.articleProblem(value);
    if (problem) throw new Error(problem + ". Исправьте артикул и повторите сохранение.");
    const ids = [...new Set((Array.isArray(lineIds) ? lineIds : []).map(String))];
    if (!ids.length) throw new Error("Не выбраны строки заказа.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const ordersSheet = this.ensureOrdersSheet();
      const orderRows = ordersSheet.getDataRange().getValues();
      const orderIndex = orderRows.findIndex((row, i) => i > 0 && String(row[0]) === String(orderId));
      if (orderIndex < 1) throw new Error("Заказ поставщику не найден.");
      const lineSheet = this.ensureLinesSheet();
      const rows = lineSheet.getDataRange().getValues().map((row, i) => ({ row: row, number: i + 1 }))
        .filter(item => item.number > 1 && String(item.row[1]) === String(orderId) && ids.includes(String(item.row[0])));
      if (rows.length !== ids.length || rows.some(item => String(item.row[6] || "") !== String(expectedArticle || ""))) {
        throw new Error("Строки или артикул уже изменились. Обновите состав заказа и повторите исправление.");
      }
      const ranges = lineSheet.getRangeList(rows.map(item => "G" + item.number));
      ranges.setNumberFormat("@");
      ranges.setValue(value);
      ordersSheet.getRange(orderIndex + 1, 6).setValue(new Date());
      return { orderId: String(orderId), lineIds: ids, article: value };
    } finally { lock.releaseLock(); }
  }

  static collectMdmExport(orderIds, allowExported = false) {
    const ids = [...new Set((Array.isArray(orderIds) ? orderIds : []).map(String).filter(Boolean))];
    if (!ids.length) throw new Error("Выберите хотя бы один заказ МДМ.");
    const sheet = this.ensureOrdersSheet();
    const all = sheet.getDataRange().getValues();
    const orders = ids.map(id => {
      const index = all.findIndex((row, i) => i > 0 && String(row[0]) === id);
      if (index < 1) throw new Error("Заказ не найден. Обновите список.");
      const row = all[index];
      if (String(row[2] || "").trim().toUpperCase() !== "МДМ" ||
          !(["Счёт получен", "Счет получен"].includes(String(row[3])) || (allowExported && row[3] === "Выгружен"))) {
        throw new Error("Для выгрузки нужны только заказы МДМ в статусе «Счёт получен». Обновите список.");
      }
      return { row: row, rowNumber: index + 1 };
    });
    const idSet = new Set(ids);
    const lines = this.ensureLinesSheet().getDataRange().getValues().slice(1).filter(row => idSet.has(String(row[1])));
    const issues = this.exportLineIssues(orders, lines);
    if (issues.length) {
      const error = new Error("Найдены проблемные позиции: " + issues.map(issue => issue.orderNumber + ": " + issue.name + " — " + issue.reason).join("; "));
      error.issues = issues;
      throw error;
    }
    const grouped = new Map();
    ids.forEach(id => {
      if (!lines.some(row => String(row[1]) === id)) throw new Error("В одном из заказов нет позиций.");
    });
    lines.forEach(row => {
      const article = String(row[6] || "").replace(/\u00a0/g, " ").trim();
      const rawQuantity = typeof row[8] === "number" ? row[8] : Number(String(row[8] || "").trim().replace(",", "."));
      if (!article || /[;\r\n\t"]/.test(article)) throw new Error("Пустой или некорректный артикул в заказе. Исправьте состав перед выгрузкой.");
      if (!Number.isFinite(rawQuantity) || rawQuantity <= 0) throw new Error("Количество должно быть положительным числом: " + article);
      const previous = grouped.get(article);
      const quantity = Number(((previous ? previous.quantity : 0) + rawQuantity).toFixed(9));
      if (!Number.isFinite(quantity) || quantity <= 0 || quantity > Number.MAX_SAFE_INTEGER) throw new Error("Некорректное суммарное количество: " + article);
      grouped.set(article, { article: article, quantity: quantity });
    });
    if (grouped.size > 500) throw new Error("МДМ принимает не более 500 артикулов в файле. Уменьшите выбор заказов.");
    const csv = [...grouped.values()].map(item => item.article + ";" + String(item.quantity).replace(".", ",")).join("\r\n");
    // Сверяем весь состав и реквизиты заказов перед сменой статуса после сохранения.
    const source = JSON.stringify({ orders: orders.map(item => [item.row[0], item.row[1], item.row[2], item.row[8]]), lines: lines });
    const digest = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8));
    return { orders: orders, csv: csv, articleCount: grouped.size, digest: digest, ids: ids };
  }

  static prepareMdmExport(orderIds) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const data = this.collectMdmExport(orderIds);
      const blob = Utilities.newBlob("").setDataFromString(data.csv, "windows-1251");
      if (blob.getDataAsString("windows-1251") !== data.csv) {
        throw new Error("В артикулах есть символы, которые нельзя сохранить в Windows-1251. Замените необычные символы в составе заказа.");
      }
      const token = Utilities.getUuid();
      CacheService.getUserCache().put("mdm-export-" + token, JSON.stringify({ ids: data.ids, digest: data.digest }), 1800);
      return { token: token, csv: data.csv, base64: Utilities.base64Encode(blob.getBytes()),
        mimeType: "text/csv;charset=windows-1251", encoding: "windows-1251", articleCount: data.articleCount, orderCount: data.ids.length,
        filename: "MDM_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") + ".csv" };
    } catch (error) {
      if (error.issues) return { issues: error.issues };
      throw error;
    } finally { lock.releaseLock(); }
  }

  static completeMdmExport(token) {
    const key = "mdm-export-" + String(token || "");
    const cache = CacheService.getUserCache();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const saved = cache.get(key);
      if (!saved) throw new Error("Подготовка выгрузки истекла. Статусы не изменены; сформируйте файл заново.");
      const record = JSON.parse(saved);
      const data = this.collectMdmExport(record.ids, true);
      if (data.digest !== record.digest) throw new Error("Состав или реквизиты заказа изменились после формирования файла. Статусы не изменены; повторите выгрузку.");
      const sheet = this.ensureOrdersSheet();
      sheet.getRangeList(data.orders.map(item => "D" + item.rowNumber)).setValue("Выгружен");
      sheet.getRangeList(data.orders.map(item => "F" + item.rowNumber)).setValue(new Date());
      // Состав не меняется. Материалы остаются в статусе «Счёт получен».
      return { orderCount: data.orders.length, status: "Выгружен" };
    } finally { lock.releaseLock(); }
  }

  static get ORDER_HEADERS() {
    return ["ID", "Номер", "Поставщик", "Статус", "Создано", "Обновлено", "Сумма", "Позиций", "Номер счёта", "Общая дата поставки"];
  }

  static get LINE_HEADERS() {
    return ["ID", "ID заказа поставщика", "ID материала", "Договор", "Клиент", "Изделие", "Артикул", "Наименование", "Количество", "Единица", "Цена", "Сумма", "Индивидуальная дата поставки"];
  }

  static ensureOrdersSheet() {
    return this.ensureSheet(CONFIG.SHEETS.SUPPLIER_ORDERS, this.ORDER_HEADERS, "#155e75");
  }

  static ensureLinesSheet() {
    return this.ensureSheet(CONFIG.SHEETS.SUPPLIER_ORDER_LINES, this.LINE_HEADERS, "#155e75");
  }

  static ensureSheet(name, headers, color) {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = spreadsheet.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
    } else if (sheet.getLastColumn() < headers.length || sheet.getRange(1, headers.length).getDisplayValue() !== headers[headers.length - 1]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
    }
    return sheet;
  }

  static getOrders() {
    const ordersSheet = SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIER_ORDERS);
    if (!ordersSheet || ordersSheet.getLastRow() < 2) return [];
    const rows = ordersSheet.getDataRange().getValues().slice(1).filter(row => row[0]);
    const linesSheet = SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIER_ORDER_LINES);
    const productsByOrderId = {};
    const articleIssuesByOrderId = {};
    const datesByOrderId = {};
    if (linesSheet && linesSheet.getLastRow() >= 2) {
      linesSheet.getDataRange().getValues().slice(1).forEach(line => {
        const orderId = String(line[1] || "").trim();
        if (orderId && this.articleProblem(line[6])) articleIssuesByOrderId[orderId] = (articleIssuesByOrderId[orderId] || 0) + 1;
        if (!datesByOrderId[orderId]) datesByOrderId[orderId] = [];
        datesByOrderId[orderId].push(this.deliveryDate(line[12]));
        const product = [line[3], line[4], line[5]].map(value => String(value || "").trim()).filter(Boolean).join(" · ");
        if (!orderId || !product) return;
        if (!productsByOrderId[orderId]) productsByOrderId[orderId] = new Set();
        productsByOrderId[orderId].add(product);
      });
    }
    return rows.map(row => ({
      id: String(row[0]), number: String(row[1]), supplier: String(row[2]), status: String(row[3] || "Сформирован"),
      createdAt: row[4], updatedAt: row[5], total: row[6], lineCount: Number(row[7] || 0),
      invoiceNumbers: String(row[8] || ""),
      defaultDeliveryDate: this.deliveryDate(row[9]),
      deliveryDate: (datesByOrderId[String(row[0])] || []).map(date => date || this.deliveryDate(row[9])).filter(Boolean).sort().pop() || "",
      deliveryDateMissingCount: (datesByOrderId[String(row[0])] || []).filter(date => !date && !this.deliveryDate(row[9])).length,
      articleIssueCount: articleIssuesByOrderId[String(row[0])] || 0,
      products: [...(productsByOrderId[String(row[0])] || [])].sort((left, right) => left.localeCompare(right, "ru"))
    })).sort((left, right) => new Date(right.createdAt || 0) - new Date(left.createdAt || 0));
  }

  static getSnapshot() {
    this.reconcileMaterialStatuses();
    const orders = this.getOrders();
    return JSON.parse(JSON.stringify({
      section: "supplier_orders",
      orders: orders,
      dictionaries: {
        suppliers: [...new Set(
          SupplyService.getSupplierRegistry().filter(item => item.active).map(item => item.name)
            .concat(orders.map(order => order.supplier).filter(Boolean))
        )]
          .sort((left, right) => left.localeCompare(right, "ru")),
        supplierOrderStatuses: this.ORDER_STATUSES,
        supplierOrders: orders
      }
    }));
  }

  static reconcileMaterialStatuses() {
    const cache = CacheService.getScriptCache();
    const cacheKey = "supplier-order-material-status-sync-v1";
    if (cache.get(cacheKey)) return { updated: 0 };
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      if (cache.get(cacheKey)) return { updated: 0 };
      const orders = this.getOrders();
      const statusesByOrderId = orders.reduce((result, order) => {
        result[order.id] = this.MATERIAL_STATUS_BY_ORDER_STATUS[order.status] || "Заказ поставщику";
        return result;
      }, {});
      const lineSheet = SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIER_ORDER_LINES);
      if (!Object.keys(statusesByOrderId).length || !lineSheet || lineSheet.getLastRow() < 2) {
        cache.put(cacheKey, "1", 21600);
        return { updated: 0 };
      }
      const desiredByMaterialId = {};
      lineSheet.getDataRange().getValues().slice(1).forEach(row => {
        const status = statusesByOrderId[String(row[1] || "")];
        const materialId = String(row[2] || "");
        if (status && materialId) desiredByMaterialId[materialId] = status;
      });
      const supplySheet = SupplyService.ensureSheet();
      const supplyRows = supplySheet.getDataRange().getValues();
      const rowsByStatus = {};
      supplyRows.slice(1).forEach((row, index) => {
        const nextStatus = desiredByMaterialId[String(row[0] || "")];
        if (!nextStatus || String(row[10] || "Новый") === nextStatus) return;
        if (!rowsByStatus[nextStatus]) rowsByStatus[nextStatus] = [];
        rowsByStatus[nextStatus].push(index + 2);
      });
      const updatedRows = Object.values(rowsByStatus).reduce((total, rows) => total + rows.length, 0);
      Object.keys(rowsByStatus).forEach(status => {
        supplySheet.getRangeList(rowsByStatus[status].map(row => "K" + row)).setValue(status);
      });
      if (updatedRows) supplySheet.getRangeList(Object.values(rowsByStatus).flat().map(row => "N" + row)).setValue(new Date());
      cache.put(cacheKey, "1", 21600);
      return { updated: updatedRows };
    } finally {
      lock.releaseLock();
    }
  }

  static setStatus(orderId, status) {
    const id = String(orderId || "").trim();
    const nextStatus = String(status || "").trim();
    if (!this.ORDER_STATUSES.includes(nextStatus)) throw new Error("Недопустимый статус заказа поставщику.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const ordersSheet = this.ensureOrdersSheet();
      const orderRows = ordersSheet.getDataRange().getValues();
      const orderIndex = orderRows.slice(1).findIndex(row => String(row[0] || "") === id);
      if (orderIndex < 0) throw new Error("Заказ поставщику не найден.");
      const rowNumber = orderIndex + 2;
      const lineSheet = this.ensureLinesSheet();
      const materialIds = new Set(lineSheet.getDataRange().getValues().slice(1)
        .filter(row => String(row[1] || "") === id)
        .map(row => String(row[2] || "")).filter(Boolean));
      const supplySheet = SupplyService.ensureSheet();
      const sourceRows = supplySheet.getDataRange().getValues().slice(1)
        .map((row, index) => ({ id: String(row[0] || ""), rowNumber: index + 2 }))
        .filter(item => materialIds.has(item.id));
      const now = new Date();
      ordersSheet.getRange(rowNumber, 4).setValue(nextStatus);
      ordersSheet.getRange(rowNumber, 6).setValue(now);
      this.syncSourceMaterialStatuses(id, nextStatus, now, sourceRows);
      // Не возвращаем объект заказа целиком: в нём есть даты Google Sheets,
      // которые нельзя передавать через google.script.run. Актуальный заказ
      // уже приходит отдельным JSON-безопасным snapshot в точке входа RPC.
      return { id: id, status: nextStatus, materialCount: sourceRows.length };
    } finally {
      lock.releaseLock();
    }
  }

  static syncSourceMaterialStatuses(orderId, status, updatedAt, sourceRows) {
    const rows = sourceRows || (() => {
      const lineSheet = this.ensureLinesSheet();
      const materialIds = new Set(lineSheet.getDataRange().getValues().slice(1)
        .filter(row => String(row[1] || "") === String(orderId || ""))
        .map(row => String(row[2] || "")).filter(Boolean));
      const supplySheet = SupplyService.ensureSheet();
      return supplySheet.getDataRange().getValues().slice(1)
        .map((row, index) => ({ id: String(row[0] || ""), rowNumber: index + 2 }))
        .filter(item => materialIds.has(item.id));
    })();
    if (!rows.length) return 0;
    const supplySheet = SupplyService.ensureSheet();
    const materialStatus = this.MATERIAL_STATUS_BY_ORDER_STATUS[status];
    if (!materialStatus) return 0;
    supplySheet.getRangeList(rows.map(item => "K" + item.rowNumber)).setValue(materialStatus);
    supplySheet.getRangeList(rows.map(item => "N" + item.rowNumber)).setValue(updatedAt || new Date());
    return rows.length;
  }

  static saveInvoiceNumbers(orderId, invoiceNumbers) {
    const id = String(orderId || "").trim();
    const value = String(invoiceNumbers || "").replace(/\r\n?/g, "\n").trim();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const sheet = this.ensureOrdersSheet();
      const rows = sheet.getDataRange().getValues();
      const index = rows.slice(1).findIndex(row => String(row[0] || "") === id);
      if (index < 0) throw new Error("Заказ поставщику не найден.");
      const rowNumber = index + 2;
      const currentStatus = String(rows[rowNumber - 1][3] || "Сформирован");
      const now = new Date();
      sheet.getRange(rowNumber, 9).setValue(value);
      sheet.getRange(rowNumber, 6).setValue(now);
      let status = currentStatus;
      if (value && currentStatus === "Сформирован") {
        status = "Счёт получен";
        sheet.getRange(rowNumber, 4).setValue(status);
        this.syncSourceMaterialStatuses(id, status, now);
      }
      return { id: id, invoiceNumbers: value, status: status };
    } finally {
      lock.releaseLock();
    }
  }

  static appendInvoiceNumber(orderId, invoiceNumber) {
    const value = String(invoiceNumber || "").trim();
    const order = this.getOrders().find(item => item.id === String(orderId || "").trim());
    if (!order) throw new Error("Заказ поставщику не найден.");
    const existing = String(order.invoiceNumbers || "").split(/\r?\n/).map(item => item.trim()).filter(Boolean);
    if (value && !existing.includes(value)) existing.push(value);
    return this.saveInvoiceNumbers(order.id, existing.join("\n"));
  }

  static markInvoiceReceived(orderId) {
    const order = this.getOrders().find(item => item.id === String(orderId || "").trim());
    if (!order) throw new Error("Заказ поставщику не найден.");
    return order.status === "Сформирован" ? this.setStatus(order.id, "Счёт получен") : { id: order.id, status: order.status };
  }

  static getOrderedMaterialIds() {
    const sheet = SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIER_ORDER_LINES);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getDisplayValues().map(row => String(row[0] || "")).filter(Boolean);
  }

  static getCandidates(supplier, productKeys) {
    const selectedSupplier = String(supplier || "").trim();
    if (!selectedSupplier || selectedSupplier === "Расходники") throw new Error("Выберите поставщика для заказа.");
    const selectedProducts = new Set((productKeys || []).map(String));
    const supplySheet = SupplyService.ensureSheet();
    const ordered = new Set(this.getOrderedMaterialIds());
    const candidates = supplySheet.getDataRange().getValues().slice(1).filter(row =>
      String(row[8] || "").trim() === selectedSupplier && String(row[10] || "Новый") === "Новый" &&
      String(row[17] || "") !== "Расходники" && String(row[26] || "") !== "Ручная заявка" && !ordered.has(String(row[0]))
    ).map(row => {
      const productKey = [row[1], row[2], row[3]].join("|");
      const orderAmount = this.getSupplyOrderAmount(row); const quantity = orderAmount.quantity; const price = row[9] === "" || row[9] === null ? "" : Number(row[9]);
      return { rowId: String(row[0]), contract: String(row[1]), client: String(row[2]), product: String(row[3]), productKey: productKey,
        article: String(row[4] || ""), name: String(row[5] || ""), quantity: quantity, unit: orderAmount.unit,
        price: Number.isFinite(price) ? price : "", total: Number.isFinite(price) ? quantity * price : "" };
    }).filter(item => !selectedProducts.size || selectedProducts.has(item.productKey));
    const products = [...new Map(candidates.map(item => [item.productKey, { key: item.productKey, contract: item.contract, client: item.client, product: item.product }])).values()]
      .sort((left, right) => (left.contract + left.product).localeCompare(right.contract + right.product, "ru"));
    return { supplier: selectedSupplier, products: products, candidates: candidates };
  }

  static createFromSupplier(supplier, rowIds) {
    const selectedSupplier = String(supplier || "").trim();
    if (!selectedSupplier || selectedSupplier === "Расходники") throw new Error("Выберите поставщика для заказа.");
    const selectedIds = [...new Set((rowIds || []).map(String).filter(Boolean))];
    if (!selectedIds.length) throw new Error("Выберите хотя бы одну позицию для заказа.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const supplySheet = SupplyService.ensureSheet();
      const supplyRows = supplySheet.getDataRange().getValues();
      const lineSheet = this.ensureLinesSheet();
      const existingLines = lineSheet.getDataRange().getValues().slice(1);
      const alreadyOrdered = new Set(existingLines.map(row => String(row[2] || "")).filter(Boolean));
      const existingOrderIdsBySelectedRows = selectedIds.map(id => {
        const line = existingLines.find(item => String(item[2] || "") === id);
        return line ? String(line[1] || "") : "";
      }).filter(Boolean);
      const existingOrderIds = new Set(existingOrderIdsBySelectedRows);
      // Apps Script does not provide a transaction across three sheets. If a
      // past attempt created the order lines but failed when updating statuses,
      // a repeated click completes that attempt instead of creating a duplicate.
      if (existingOrderIdsBySelectedRows.length === selectedIds.length && existingOrderIds.size === 1) {
        const existingOrderId = [...existingOrderIds][0];
        const existingOrder = this.getOrders().find(order => order.id === existingOrderId);
        if (existingOrder && existingOrder.supplier === selectedSupplier) {
          const sourceRows = supplyRows.slice(1).map((row, index) => ({ id: String(row[0]), rowNumber: index + 2 }))
            .filter(item => selectedIds.includes(item.id));
          supplySheet.getRangeList(sourceRows.map(item => "K" + item.rowNumber)).setValue("Заказ поставщику");
          supplySheet.getRangeList(sourceRows.map(item => "N" + item.rowNumber)).setValue(new Date());
          return { id: existingOrder.id, number: existingOrder.number, supplier: existingOrder.supplier, lineCount: existingOrder.lineCount };
        }
      }
      const candidates = supplyRows.slice(1).map((row, index) => ({ row: row, rowNumber: index + 2 }))
        .filter(item => String(item.row[8] || "").trim() === selectedSupplier)
        .filter(item => String(item.row[10] || "Новый") === "Новый")
        .filter(item => String(item.row[17] || "") !== "Расходники")
        .filter(item => String(item.row[26] || "") !== "Ручная заявка")
        .filter(item => !alreadyOrdered.has(String(item.row[0])));
      const byId = new Map(candidates.map(item => [String(item.row[0]), item]));
      const selected = selectedIds.map(id => byId.get(id));
      if (selected.some(item => !item)) throw new Error("Часть позиций изменилась, уже заказана или больше не относится к поставщику. Обновите подбор.");

      const ordersSheet = this.ensureOrdersSheet();
      const now = new Date();
      const orderId = Utilities.getUuid();
      const number = this.nextNumber(ordersSheet, now);
      const hasUnknownPrice = selected.some(item => item.row[9] === "" || item.row[9] === null || !Number.isFinite(Number(item.row[9])));
      const total = selected.reduce((sum, item) => sum + this.getSupplyOrderAmount(item.row).quantity * Number(item.row[9] || 0), 0);
      const lineRows = selected.map(item => {
        const row = item.row;
        const orderAmount = this.getSupplyOrderAmount(row); const quantity = orderAmount.quantity;
        const price = row[9] === "" || row[9] === null ? NaN : Number(row[9]);
        return [Utilities.getUuid(), orderId, String(row[0]), row[1], row[2], row[3], row[4], row[5], quantity, orderAmount.unit, Number.isFinite(price) ? price : "", Number.isFinite(price) ? quantity * price : "", ""];
      });
      ordersSheet.appendRow([orderId, number, selectedSupplier, this.ORDER_STATUSES[0], now, now, hasUnknownPrice ? "" : total, lineRows.length, ""]);
      lineSheet.getRange(lineSheet.getLastRow() + 1, 1, lineRows.length, this.LINE_HEADERS.length).setValues(lineRows);
      supplySheet.getRangeList(selected.map(item => "K" + item.rowNumber)).setValue("Заказ поставщику");
      supplySheet.getRangeList(selected.map(item => "N" + item.rowNumber)).setValue(now);
      return { id: orderId, number: number, supplier: selectedSupplier, lineCount: lineRows.length };
    } finally {
      lock.releaseLock();
    }
  }

  static addMaterials(orderId, rowIds) {
    const id = String(orderId || "").trim();
    const selectedIds = [...new Set((rowIds || []).map(String).filter(Boolean))];
    if (!selectedIds.length) throw new Error("Выберите хотя бы одну позицию для добавления.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const ordersSheet = this.ensureOrdersSheet();
      const orderRows = ordersSheet.getDataRange().getValues();
      const orderIndex = orderRows.slice(1).findIndex(row => String(row[0] || "") === id);
      if (orderIndex < 0) throw new Error("Заказ поставщику не найден.");
      const orderRow = orderRows[orderIndex + 1];
      if (String(orderRow[3] || "Сформирован") !== "Сформирован") throw new Error("Состав можно менять только у сформированного заказа.");
      const supplier = String(orderRow[2] || "").trim();
      const supplySheet = SupplyService.ensureSheet();
      const supplyRows = supplySheet.getDataRange().getValues();
      const lineSheet = this.ensureLinesSheet();
      const ordered = new Set(lineSheet.getDataRange().getValues().slice(1).map(row => String(row[2] || "")).filter(Boolean));
      const candidates = supplyRows.slice(1).map((row, index) => ({ row: row, rowNumber: index + 2 }))
        .filter(item => String(item.row[8] || "").trim() === supplier)
        .filter(item => String(item.row[10] || "Новый") === "Новый")
        .filter(item => String(item.row[17] || "") !== "Расходники")
        .filter(item => !ordered.has(String(item.row[0] || "")));
      const byId = new Map(candidates.map(item => [String(item.row[0]), item]));
      const selected = selectedIds.map(materialId => byId.get(materialId));
      if (selected.some(item => !item)) throw new Error("Часть позиций уже заказана, изменила статус или относится к другому поставщику. Обновите подбор.");
      const lineRows = selected.map(item => {
        const row = item.row;
        const orderAmount = this.getSupplyOrderAmount(row); const quantity = orderAmount.quantity;
        const price = row[9] === "" || row[9] === null ? NaN : Number(row[9]);
        return [Utilities.getUuid(), id, String(row[0]), row[1], row[2], row[3], row[4], row[5], quantity, orderAmount.unit, Number.isFinite(price) ? price : "", Number.isFinite(price) ? quantity * price : "", ""];
      });
      lineSheet.getRange(lineSheet.getLastRow() + 1, 1, lineRows.length, this.LINE_HEADERS.length).setValues(lineRows);
      const now = new Date();
      supplySheet.getRangeList(selected.map(item => "K" + item.rowNumber)).setValue("Заказ поставщику");
      supplySheet.getRangeList(selected.map(item => "N" + item.rowNumber)).setValue(now);
      return this.updateOrderSummary(id, ordersSheet, lineSheet, orderIndex + 2, now);
    } finally {
      lock.releaseLock();
    }
  }

  static getSupplyOrderAmount(row) {
    const purchaseQuantity = Number(row[23]);
    const purchaseUnit = String(row[24] || "").trim();
    if (Number.isFinite(purchaseQuantity) && purchaseQuantity > 0 && purchaseUnit) return { quantity: purchaseQuantity, unit: purchaseUnit };
    return { quantity: Number(row[6] || 0), unit: String(row[7] || "") };
  }

  static updateLine(orderId, lineId, quantity, unit) {
    const id = String(orderId || "").trim();
    const targetLineId = String(lineId || "").trim();
    const parsedQuantity = Number(String(quantity || "").replace(",", "."));
    const nextUnit = String(unit || "").trim();
    if (!Number.isFinite(parsedQuantity) || parsedQuantity <= 0) throw new Error("Количество должно быть больше нуля.");
    if (!nextUnit) throw new Error("Укажите единицу измерения.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const ordersSheet = this.ensureOrdersSheet();
      const orderRows = ordersSheet.getDataRange().getValues();
      const orderIndex = orderRows.slice(1).findIndex(row => String(row[0] || "") === id);
      if (orderIndex < 0) throw new Error("Заказ поставщику не найден.");
      if (String(orderRows[orderIndex + 1][3] || "Сформирован") !== "Сформирован") throw new Error("Состав можно менять только у сформированного заказа.");
      const lineSheet = this.ensureLinesSheet();
      const lineRows = lineSheet.getDataRange().getValues();
      const lineIndex = lineRows.slice(1).findIndex(row => String(row[0] || "") === targetLineId && String(row[1] || "") === id);
      if (lineIndex < 0) throw new Error("Позиция заказа не найдена.");
      const rowNumber = lineIndex + 2;
      const price = lineRows[rowNumber - 1][10];
      const numericPrice = price === "" || price === null ? NaN : Number(price);
      lineSheet.getRange(rowNumber, 9, 1, 4).setValues([[parsedQuantity, nextUnit, Number.isFinite(numericPrice) ? numericPrice : "", Number.isFinite(numericPrice) ? parsedQuantity * numericPrice : ""]]);
      return this.updateOrderSummary(id, ordersSheet, lineSheet, orderIndex + 2, new Date());
    } finally {
      lock.releaseLock();
    }
  }

  static removeLines(orderId, lineIds) {
    const id = String(orderId || "").trim();
    const targetLineIds = [...new Set((lineIds || []).map(String).filter(Boolean))];
    if (!targetLineIds.length) throw new Error("Выберите позицию для удаления.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const ordersSheet = this.ensureOrdersSheet();
      const orderRows = ordersSheet.getDataRange().getValues();
      const orderIndex = orderRows.slice(1).findIndex(row => String(row[0] || "") === id);
      if (orderIndex < 0) throw new Error("Заказ поставщику не найден.");
      if (String(orderRows[orderIndex + 1][3] || "Сформирован") !== "Сформирован") throw new Error("Состав можно менять только у сформированного заказа.");
      const lineSheet = this.ensureLinesSheet();
      const lineRows = lineSheet.getDataRange().getValues();
      const selected = lineRows.slice(1).map((row, index) => ({ row: row, rowNumber: index + 2 }))
        .filter(item => String(item.row[1] || "") === id && targetLineIds.includes(String(item.row[0] || "")));
      if (selected.length !== targetLineIds.length) throw new Error("Часть позиций заказа не найдена. Обновите состав.");
      const removedMaterialIds = new Set(selected.map(item => String(item.row[2] || "")).filter(Boolean));
      selected.map(item => item.rowNumber).sort((left, right) => right - left).forEach(rowNumber => lineSheet.deleteRow(rowNumber));
      const remainingMaterialIds = new Set(lineSheet.getDataRange().getValues().slice(1).map(row => String(row[2] || "")).filter(Boolean));
      const returnMaterialIds = [...removedMaterialIds].filter(materialId => !remainingMaterialIds.has(materialId));
      if (returnMaterialIds.length) {
        const supplySheet = SupplyService.ensureSheet();
        const sourceRows = supplySheet.getDataRange().getValues().slice(1).map((row, index) => ({ id: String(row[0] || ""), rowNumber: index + 2 }))
          .filter(item => returnMaterialIds.includes(item.id));
        if (sourceRows.length) {
          supplySheet.getRangeList(sourceRows.map(item => "K" + item.rowNumber)).setValue("Новый");
          supplySheet.getRangeList(sourceRows.map(item => "N" + item.rowNumber)).setValue(new Date());
        }
      }
      return this.updateOrderSummary(id, ordersSheet, lineSheet, orderIndex + 2, new Date());
    } finally {
      lock.releaseLock();
    }
  }

  /** Переносит строки одного изделия в новый заказ того же поставщика. */
  static splitLinesToNewOrder(orderId, lineIds) {
    const id = String(orderId || "").trim();
    const targetLineIds = [...new Set((lineIds || []).map(String).filter(Boolean))];
    if (!targetLineIds.length) throw new Error("Выберите позиции для отдельного заказа.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const ordersSheet = this.ensureOrdersSheet();
      const orderRows = ordersSheet.getDataRange().getValues();
      const orderIndex = orderRows.slice(1).findIndex(row => String(row[0] || "") === id);
      if (orderIndex < 0) throw new Error("Заказ поставщику не найден.");
      const sourceOrder = orderRows[orderIndex + 1];
      if (String(sourceOrder[3] || "Сформирован") !== "Сформирован") {
        throw new Error("Состав можно менять только у сформированного заказа.");
      }
      const lineSheet = this.ensureLinesSheet();
      const lineRows = lineSheet.getDataRange().getValues();
      const selected = lineRows.slice(1).map((row, index) => ({ row: row, rowNumber: index + 2 }))
        .filter(item => String(item.row[1] || "") === id && targetLineIds.includes(String(item.row[0] || "")));
      if (selected.length !== targetLineIds.length) throw new Error("Часть позиций заказа не найдена. Обновите состав.");
      const productKeys = new Set(selected.map(item => [item.row[3], item.row[4], item.row[5]].map(value => String(value || "")).join("\u0001")));
      if (productKeys.size !== 1) {
        throw new Error("В отдельный заказ можно вынести позиции только одного изделия. Отметьте строки одного договора, клиента и изделия.");
      }

      const now = new Date();
      const newOrderId = Utilities.getUuid();
      const number = this.nextNumber(ordersSheet, now);
      const hasUnknownPrice = selected.some(item => item.row[10] === "" || item.row[10] === null || !Number.isFinite(Number(item.row[10])));
      const total = selected.reduce((sum, item) => sum + Number(item.row[8] || 0) * Number(item.row[10] || 0), 0);
      ordersSheet.appendRow([newOrderId, number, sourceOrder[2], this.ORDER_STATUSES[0], now, now, hasUnknownPrice ? "" : total, selected.length]);
      lineSheet.getRangeList(selected.map(item => "B" + item.rowNumber)).setValue(newOrderId);
      const source = this.updateOrderSummary(id, ordersSheet, lineSheet, orderIndex + 2, now);
      return {
        source: source,
        newOrder: { id: newOrderId, number: number, supplier: String(sourceOrder[2] || ""), lineCount: selected.length }
      };
    } finally {
      lock.releaseLock();
    }
  }

  static updateOrderSummary(orderId, ordersSheet, lineSheet, orderRowNumber, updatedAt) {
    const lines = lineSheet.getDataRange().getValues().slice(1).filter(row => String(row[1] || "") === orderId);
    if (!lines.length) {
      ordersSheet.deleteRow(orderRowNumber);
      return { id: orderId, deleted: true, lineCount: 0 };
    }
    const hasUnknownPrice = lines.some(row => row[10] === "" || row[10] === null || !Number.isFinite(Number(row[10])));
    const total = lines.reduce((sum, row) => sum + Number(row[8] || 0) * Number(row[10] || 0), 0);
    ordersSheet.getRange(orderRowNumber, 6, 1, 3).setValues([[updatedAt, hasUnknownPrice ? "" : total, lines.length]]);
    return { id: orderId, deleted: false, lineCount: lines.length, total: hasUnknownPrice ? "" : total };
  }

  static getOrderDetails(orderId) {
    const id = String(orderId || "").trim();
    const order = this.getOrders().find(item => item.id === id);
    if (!order) throw new Error("Заказ поставщику не найден.");
    const lines = this.ensureLinesSheet().getDataRange().getValues().slice(1)
      .filter(row => String(row[1]) === id)
      .map(row => ({ lineId: String(row[0]), materialId: String(row[2]), contract: row[3], client: row[4], product: row[5], article: row[6], name: row[7], quantity: Number(row[8] || 0), unit: row[9], price: row[10], total: row[11], deliveryDateOverride: this.deliveryDate(row[12]), deliveryDate: this.deliveryDate(row[12]) || order.defaultDeliveryDate }));
    const grouped = {};
    lines.forEach(line => {
      const key = [line.article, line.name, line.unit, line.price, line.deliveryDateOverride].join("|");
      if (!grouped[key]) grouped[key] = Object.assign({}, line, { sourceCount: 0, lineIds: [] });
      else grouped[key].quantity += line.quantity;
      grouped[key].sourceCount++;
      grouped[key].lineIds.push(line.lineId);
    });
    // google.script.run accepts only plain JSON-like values.  Date values read
    // from a sheet may otherwise make the whole RPC result arrive as null in
    // the browser, even though the lines themselves were found correctly.
    return JSON.parse(JSON.stringify({
      order: order,
      lines: Object.keys(grouped)
        .map(key => grouped[key])
        .sort((left, right) => String(left.name).localeCompare(String(right.name), "ru")),
      detailedLines: lines.sort((left, right) => (String(left.contract) + String(left.product) + String(left.name)).localeCompare(String(right.contract) + String(right.product) + String(right.name), "ru"))
    }));
  }

  /**
   * Формирует временную Google-таблицу, отдаёт её как XLSX и обязательно
   * удаляет временный файл. В Drive выгрузка не сохраняется.
   */
  static exportExcel(orderId) {
    const data = this.getOrderDetails(orderId);
    const order = data.order;
    const detailedLines = data.detailedLines || [];
    if (!detailedLines.length) {
      throw new Error("В заказе поставщику нет позиций для выгрузки в Excel.");
    }

    const groups = {};
    detailedLines.forEach(line => {
      const groupKey = [line.contract, line.client, line.product].map(value => String(value || "")).join("\u0001");
      if (!groups[groupKey]) {
        groups[groupKey] = {
          contract: String(line.contract || ""),
          client: String(line.client || ""),
          product: String(line.product || ""),
          materials: {}
        };
      }
      const materialKey = [line.article, line.name, line.unit].map(value => String(value || "")).join("\u0001");
      if (!groups[groupKey].materials[materialKey]) {
        groups[groupKey].materials[materialKey] = {
          article: String(line.article || ""),
          name: String(line.name || ""),
          quantity: 0,
          unit: String(line.unit || "")
        };
      }
      groups[groupKey].materials[materialKey].quantity += Number(line.quantity || 0);
    });

    const safePart = value => String(value || "").trim()
      .replace(/[\\/:*?"<>|]+/g, "_").replace(/\s+/g, " ").slice(0, 80) || "Заказ";
    const fileName = safePart(order.number) + "_" + safePart(order.supplier) + ".xlsx";
    const spreadsheet = SpreadsheetApp.create("Временный экспорт " + safePart(order.number));
    const spreadsheetId = spreadsheet.getId();
    const sheet = spreadsheet.getSheets()[0];
    const xlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    try {
      sheet.setName("Материалы");
      sheet.getRange(1, 1, 1, 4).merge();
      sheet.getRange(1, 1).setValue("Заказ поставщику " + order.number + " · " + order.supplier)
        .setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14)
        .setHorizontalAlignment("center");
      let row = 3;
      Object.keys(groups).map(key => groups[key]).sort((left, right) =>
        (left.contract + left.client + left.product).localeCompare(right.contract + right.client + right.product, "ru")
      ).forEach(group => {
        [["Договор", group.contract], ["Клиент", group.client], ["Изделие", group.product]].forEach(pair => {
          sheet.getRange(row, 1).setValue(pair[0]).setFontWeight("bold").setBackground("#e0f2fe");
          sheet.getRange(row, 2, 1, 3).merge().setValue(pair[1]).setBackground("#f8fafc");
          row++;
        });
        sheet.getRange(row, 1, 1, 4).setValues([["Артикул", "Наименование", "Количество", "Единица измерения"]])
          .setBackground("#0f766e").setFontColor("#ffffff").setFontWeight("bold");
        row++;
        const materials = Object.keys(group.materials).map(key => group.materials[key]).sort((left, right) =>
          (left.name + left.article).localeCompare(right.name + right.article, "ru")
        );
        const values = materials.map(item => [item.article, item.name, item.quantity, item.unit]);
        sheet.getRange(row, 1, values.length, 4).setValues(values).setVerticalAlignment("top").setWrap(true)
          .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(row, 3, values.length, 1).setNumberFormat("0.###");
        row += values.length + 2;
      });
      sheet.autoResizeColumns(1, 4);
      sheet.setColumnWidth(2, 420);
      SpreadsheetApp.flush();

      const response = UrlFetchApp.fetch(
        "https://www.googleapis.com/drive/v3/files/" + spreadsheetId + "/export?mimeType=" + encodeURIComponent(xlsxMimeType),
        { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
      );
      if (response.getResponseCode() !== 200) {
        throw new Error("Не удалось сформировать Excel-файл: " + response.getContentText().slice(0, 300));
      }
      return {
        filename: fileName,
        mimeType: xlsxMimeType,
        base64: Utilities.base64Encode(response.getBlob().getBytes())
      };
    } finally {
      try {
        Drive.Files.trash(spreadsheetId);
      } catch (cleanupError) {
        console.warn("Не удалось удалить временную таблицу экспорта: " + cleanupError.message);
      }
    }
  }

  static nextNumber(sheet, date) {
    const prefix = "ЗП-" + Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd") + "-";
    const numbers = sheet.getDataRange().getDisplayValues().slice(1).map(row => String(row[1] || ""))
      .filter(value => value.indexOf(prefix) === 0).map(value => Number(value.slice(prefix.length)) || 0);
    return prefix + String((numbers.length ? Math.max.apply(null, numbers) : 0) + 1).padStart(3, "0");
  }
}
