/** Счета по технике. Счёт хранится отдельно от позиции и может связывать несколько позиций. */
class EquipmentInvoiceService {
  static get INVOICE_HEADERS() {
    return ["ID", "Номер счёта", "Дата счёта", "Поставщик", "Сумма", "ID файла Drive", "Имя файла", "Mime type", "Статус оплаты", "Дата оплаты", "Создано"];
  }

  static get LINE_HEADERS() {
    return ["ID", "ID счёта", "ID позиции техники", "Договор", "Клиент", "Наименование техники", "Модель", "Количество", "Цена за единицу, ₽", "Сумма, ₽"];
  }

  static get PAYMENT_STATUSES() { return ["Счёт", "Оплачен"]; }

  static ensureSheet(name, headers) {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = spreadsheet.insertSheet(name);
    if (!sheet.getLastRow()) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    else {
      const actual = sheet.getRange(1, 1, 1, Math.max(headers.length, sheet.getLastColumn())).getDisplayValues()[0];
      if (headers.some((header, index) => actual[index] !== header)) {
        throw new Error("Лист «" + name + "» имеет несовместимые заголовки.");
      }
    }
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#0f766e").setFontColor("#ffffff").setFontWeight("bold");
    return sheet;
  }

  static ensureInvoicesSheet() { return this.ensureSheet(CONFIG.SHEETS.EQUIPMENT_INVOICES, this.INVOICE_HEADERS); }
  static ensureLinesSheet() { return this.ensureSheet(CONFIG.SHEETS.EQUIPMENT_INVOICE_LINES, this.LINE_HEADERS); }

  static getSnapshot() {
    const invoices = this.getInvoices();
    const suppliers = [...new Set(SupplyService.getSupplierRegistry().filter(item => item.active).map(item => item.name).concat(invoices.map(item => item.supplier).filter(Boolean)))]
      .sort((a, b) => String(a).localeCompare(String(b), "ru"));
    return JSON.parse(JSON.stringify({
      section: "equipment_invoices", orders: invoices,
      dictionaries: { suppliers: suppliers, equipmentInvoiceStatuses: this.PAYMENT_STATUSES }
    }));
  }

  static getInvoices() {
    const sheet = this.ensureInvoicesSheet();
    if (sheet.getLastRow() < 2) return [];
    const lineIndex = this.getLinesByInvoice();
    return sheet.getDataRange().getValues().slice(1).filter(row => row[0]).map(row => {
      const id = String(row[0]);
      const lines = lineIndex[id] || [];
      return this.invoiceFromRow(row, lines);
    }).sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
  }

  static invoiceFromRow(row, lines) {
    return {
      id: String(row[0] || ""), number: String(row[1] || ""), date: this.formatDate(row[2]), supplier: String(row[3] || ""),
      amount: Number(row[4] || 0), fileId: String(row[5] || ""), fileName: String(row[6] || ""), mimeType: String(row[7] || ""),
      paymentStatus: String(row[8] || "Счёт"), paidAt: this.formatDate(row[9]), createdAt: this.formatDateTime(row[10]), lines: lines || []
    };
  }

  static getLinesByInvoice() {
    const sheet = this.ensureLinesSheet();
    if (sheet.getLastRow() < 2) return {};
    return sheet.getDataRange().getValues().slice(1).filter(row => row[0]).reduce((index, row) => {
      const line = { id: String(row[0]), invoiceId: String(row[1]), equipmentId: String(row[2]), contract: String(row[3] || ""), client: String(row[4] || ""), name: String(row[5] || ""), model: String(row[6] || ""), quantity: Number(row[7] || 0), unitPrice: Number(row[8] || 0), total: Number(row[9] || 0) };
      (index[line.invoiceId] = index[line.invoiceId] || []).push(line);
      return index;
    }, {});
  }

  static getInvoiceDetail(id) {
    const invoice = this.getInvoices().find(item => item.id === String(id || ""));
    if (!invoice) throw new Error("Счёт техники не найден.");
    return invoice;
  }

  static getEquipmentInvoices(equipmentId) {
    const target = String(equipmentId || "");
    return this.getInvoices().filter(invoice => invoice.lines.some(line => line.equipmentId === target));
  }

  static preview(file) {
    this.assertFile(file);
    return { fileName: String(file.name), mimeType: String(file.mimeType || ""), candidates: EquipmentService.getSnapshot().orders };
  }

  static save(payload) {
    const value = payload || {};
    this.assertFile(value.file);
    const number = String(value.number || "").trim();
    const supplier = String(value.supplier || "").trim();
    const amount = this.parseNumber(value.amount);
    const status = String(value.paymentStatus || "Счёт").trim();
    if (!number || !supplier) throw new Error("Укажите номер счёта и поставщика.");
    if (!Number.isFinite(amount) || amount < 0) throw new Error("Укажите корректную сумму счёта.");
    if (!this.PAYMENT_STATUSES.includes(status)) throw new Error("Недопустимый статус оплаты счёта.");
    const lines = (value.lines || []).map(line => ({ equipmentId: String(line.equipmentId || "").trim(), quantity: this.parseNumber(line.quantity), unitPrice: this.parseNumber(line.unitPrice) }))
      .filter(line => line.equipmentId);
    if (!lines.length) throw new Error("Отметьте хотя бы одну позицию техники.");
    if (new Set(lines.map(line => line.equipmentId)).size !== lines.length) throw new Error("Одна позиция техники не может повторяться в одном счёте.");
    lines.forEach(line => {
      if (!Number.isInteger(line.quantity) || line.quantity <= 0) throw new Error("Количество в счёте должно быть целым положительным числом.");
      if (!Number.isFinite(line.unitPrice) || line.unitPrice < 0) throw new Error("Укажите цену за единицу для каждой позиции.");
    });
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      this.ensureSupplier(supplier);
      const equipmentSheet = EquipmentService.ensureSheet();
      const equipment = lines.map(line => {
        const row = EquipmentService.findRow(equipmentSheet, line.equipmentId);
        return { line: line, row: row, item: EquipmentService.fromRow(equipmentSheet.getRange(row, 1, 1, EquipmentService.HEADERS.length).getValues()[0]) };
      });
      const invoiceSheet = this.ensureInvoicesSheet();
      const lineSheet = this.ensureLinesSheet();
      const folder = this.ensureInvoiceFolder(number);
      const file = value.file;
      const driveFile = folder.createFile(Utilities.newBlob(Utilities.base64Decode(String(file.base64)), file.mimeType || "application/octet-stream", this.safeName(file.name)));
      const invoiceId = Utilities.getUuid();
      const now = new Date();
      const paidAt = status === "Оплачен" ? (this.toDate(value.paidAt) || now) : "";
      invoiceSheet.appendRow([invoiceId, number, this.toDate(value.date), supplier, amount, driveFile.getId(), driveFile.getName(), driveFile.getMimeType(), status, paidAt, now]);
      lineSheet.getRange(lineSheet.getLastRow() + 1, 1, equipment.length, this.LINE_HEADERS.length).setValues(equipment.map(entry => [
        Utilities.getUuid(), invoiceId, entry.item.id, entry.item.contract, entry.item.client, entry.item.name, entry.item.model,
        entry.line.quantity, entry.line.unitPrice, entry.line.quantity * entry.line.unitPrice
      ]));
      equipment.forEach(entry => {
        equipmentSheet.getRange(entry.row, 7).setValue(entry.line.unitPrice);
        equipmentSheet.getRange(entry.row, 9).setValue(supplier);
        equipmentSheet.getRange(entry.row, 10).setValue(status === "Оплачен" ? "Оплачен" : "Счет");
        equipmentSheet.getRange(entry.row, 14).setValue(now);
      });
      return { id: invoiceId };
    } finally { lock.releaseLock(); }
  }

  static markPaid(id, paidAt) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const invoiceSheet = this.ensureInvoicesSheet();
      const row = this.findInvoiceRow(invoiceSheet, id);
      const invoice = invoiceSheet.getRange(row, 1, 1, this.INVOICE_HEADERS.length).getValues()[0];
      invoiceSheet.getRange(row, 9, 1, 2).setValues([["Оплачен", this.toDate(paidAt) || new Date()]]);
      const equipmentSheet = EquipmentService.ensureSheet();
      (this.getLinesByInvoice()[String(id)] || []).forEach(line => {
        const equipmentRow = EquipmentService.findRow(equipmentSheet, line.equipmentId);
        equipmentSheet.getRange(equipmentRow, 10).setValue("Оплачен");
        equipmentSheet.getRange(equipmentRow, 14).setValue(new Date());
      });
      return { id: String(id), status: "Оплачен", supplier: String(invoice[3] || "") };
    } finally { lock.releaseLock(); }
  }

  static ensureSupplier(name) {
    const exists = SupplyService.getSupplierRegistry().some(item => String(item.name).toLowerCase() === String(name).toLowerCase());
    if (!exists) SupplyService.saveSupplierRegistryItem({ name: name, active: true, note: "Добавлен из счёта техники" });
  }

  static ensureInvoiceFolder(number) {
    const root = DriveApp.getFolderById(GOOGLE_DRIVE_ORDERS_ROOT_ID);
    const roots = root.getFoldersByName("Счета техники");
    const invoicesRoot = roots.hasNext() ? roots.next() : root.createFolder("Счета техники");
    const safe = this.safeName(number || "Счёт");
    const folders = invoicesRoot.getFoldersByName(safe);
    return folders.hasNext() ? folders.next() : invoicesRoot.createFolder(safe);
  }

  static assertFile(file) {
    if (!file || !file.name || !file.base64) throw new Error("Прикрепите файл счёта.");
    if (!/\.(pdf|xlsx?|xls)$/i.test(String(file.name))) throw new Error("Поддерживаются счета в PDF, XLS и XLSX.");
    if (Utilities.base64Decode(String(file.base64)).length > 10 * 1024 * 1024) throw new Error("Размер счёта не должен превышать 10 МБ.");
  }
  static findInvoiceRow(sheet, id) { const found = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).createTextFinder(String(id || "")).matchEntireCell(true).findNext(); if (!found) throw new Error("Счёт техники не найден."); return found.getRow(); }
  static parseNumber(value) { const normalized = String(value === null || value === undefined ? "" : value).replace(/\s/g, "").replace(",", "."); return normalized === "" ? NaN : Number(normalized); }
  static toDate(value) { if (!value) return ""; if (value instanceof Date) return value; const raw = String(value); const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/) || raw.match(/^(\d{2})\.(\d{2})\.(\d{4})$/); if (!match) return ""; return new Date(match[1].length === 4 ? raw + "T00:00:00" : match[3] + "-" + match[2] + "-" + match[1] + "T00:00:00"); }
  static formatDate(value) { return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(value || ""); }
  static formatDateTime(value) { return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : String(value || ""); }
  static safeName(value) { return String(value || "файл").replace(/[\\/:*?\"<>|]/g, "_").slice(0, 180); }
}
