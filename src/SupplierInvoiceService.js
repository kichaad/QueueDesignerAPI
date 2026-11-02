/** Счета, приложенные к заказам поставщиков. */
class SupplierInvoiceService {
  static get INVOICE_HEADERS() {
    return ["ID", "ID заказа поставщика", "Номер счёта", "Дата счёта", "Сумма", "ID файла Drive", "Имя файла", "Mime type", "Статус сверки", "Разница", "Причина", "Загружен"];
  }

  static get LINE_HEADERS() {
    return ["ID", "ID счёта", "Артикул", "Наименование", "Количество", "Единица", "Цена", "Сумма", "Статус сверки", "Причина"];
  }

  static ensureInvoicesSheet() {
    return this.ensureSheet(CONFIG.SHEETS.SUPPLIER_INVOICES, this.INVOICE_HEADERS);
  }

  static ensureLinesSheet() {
    return this.ensureSheet(CONFIG.SHEETS.SUPPLIER_INVOICE_LINES, this.LINE_HEADERS);
  }

  static ensureSheet(name, headers) {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = spreadsheet.insertSheet(name);
    const existing = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0] : [];
    if (!sheet.getLastRow()) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else if (existing.join("\u0001") !== headers.join("\u0001")) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold");
    return sheet;
  }

  static getOrderSummary(orderId, orderLines) {
    const id = String(orderId || "").trim();
    const sheet = SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIER_INVOICES);
    if (!id || !sheet || sheet.getLastRow() < 2) return { invoices: [], invoiceTotal: 0, hasInvoices: false, difference: "", verification: "Нет счетов" };
    const rows = sheet.getDataRange().getValues().slice(1).filter(row => String(row[1] || "") === id);
    const invoices = rows.map(row => ({
      id: String(row[0]), number: String(row[2] || ""), date: row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[3] || ""),
      amount: Number(row[4] || 0), fileId: String(row[5] || ""), fileName: String(row[6] || ""), mimeType: String(row[7] || "application/pdf"),
      verification: String(row[8] || "Требует проверки"), difference: row[9] === "" ? "" : Number(row[9]), reason: String(row[10] || ""), uploadedAt: row[11]
    }));
    const invoiceTotal = invoices.reduce((sum, invoice) => sum + Number(invoice.amount || 0), 0);
    const orderTotal = (orderLines || []).some(line => line.price === "" || line.price === null || !Number.isFinite(Number(line.price)))
      ? "" : (orderLines || []).reduce((sum, line) => sum + Number(line.quantity || 0) * Number(line.price || 0), 0);
    const difference = orderTotal === "" ? "" : invoiceTotal - orderTotal;
    const verification = !invoices.length ? "Нет счетов" : invoices.some(invoice => invoice.verification !== "Совпадает") || (difference !== "" && Math.abs(difference) > 0.01) ? "Требует проверки" : "Совпадает";
    return JSON.parse(JSON.stringify({ invoices: invoices, invoiceTotal: invoiceTotal, hasInvoices: Boolean(invoices.length), orderTotal: orderTotal, difference: difference, verification: verification }));
  }

  static preview(orderId, file) {
    this.assertOrder(orderId);
    if (!file || !file.name || !file.base64 || !/\.pdf$/i.test(String(file.name))) throw new Error("Выберите PDF-счёт.");
    let text = String(file.clientText || "").trim();
    let source = text ? "браузер" : "";
    if (!text) {
      const extracted = PlateMaterialService.readPdfText(file);
      text = String(extracted.text || "").trim();
      source = "сервер";
    }
    const parsed = this.parseText(text);
    const verified = this.verifyLines(orderId, parsed.lines);
    return { source: source || "не распознан", textFound: Boolean(text), invoice: parsed.invoice, lines: verified.lines, verification: verified.verification, reason: verified.reason };
  }

  static parseText(text) {
    const value = String(text || "").replace(/\r/g, "").replace(/[\u00a0]/g, " ");
    const number = (value.match(/(?:сч[её]т(?:[- ]?фактура)?|счет)\s*(?:№|N|#)?\s*([A-ZА-Я0-9][A-ZА-Я0-9.\/-]{1,})/i) || [])[1] || "";
    const date = (value.match(/(?:от|дата)\s*(\d{1,2}[.]\d{1,2}[.]\d{2,4})/i) || [])[1] || "";
    const totalMatches = [...value.matchAll(/(?:всего(?:\s+к\s+оплате)?|итого(?:\s+к\s+оплате)?|к\s+оплате)[^\d]{0,30}([\d\s]+(?:[.,]\d{1,2})?)/gi)];
    const amount = totalMatches.length ? this.parseAmount(totalMatches[totalMatches.length - 1][1]) : "";
    const lines = value.split(/\n+/).map(line => line.replace(/\s+/g, " ").trim()).filter(Boolean).map(line => this.parseLine(line)).filter(Boolean);
    return { invoice: { number: number, date: this.normalizeDate(date), amount: amount }, lines: lines };
  }

  static parseLine(line) {
    if (!/^\d{1,3}[.)]\s+/.test(line) || /(итого|всего|к оплате|наименован)/i.test(line)) return null;
    const values = line.replace(/^\d{1,3}[.)]\s+/, "").split(/\s{2,}|\t+/).filter(Boolean);
    if (values.length < 2) return null;
    const trailing = line.match(/(\d+(?:[.,]\d+)?)\s*(шт|м\/п|м2|м²|м|компл\.?|уп\.?|кг|л)\b.*?(\d+(?:[.,]\d+)?)(?:\s+\d+(?:[.,]\d+)?)?\s*$/i);
    const quantity = trailing ? Number(String(trailing[1]).replace(",", ".")) : "";
    const unit = trailing ? trailing[2] : "";
    const sum = trailing ? this.parseAmount(trailing[3]) : "";
    const article = (values[0].match(/[A-ZА-Я0-9][A-ZА-Я0-9.\/-]{2,}/i) || [])[0] || "";
    return { article: article, name: values.slice(article ? 1 : 0).join(" ").replace(/\s+\d+(?:[.,]\d+)?\s*(?:шт|м\/п|м2|м²|м|компл\.?|уп\.?|кг|л).*$/i, "").trim() || line, quantity: quantity, unit: unit, price: "", total: sum };
  }

  static save(orderId, payload) {
    const order = this.assertOrder(orderId);
    const file = payload && payload.file;
    if (!file || !file.name || !file.base64) throw new Error("Файл счёта не получен.");
    const amount = this.parseAmount(payload.amount);
    if (!Number.isFinite(amount) || amount < 0) throw new Error("Укажите корректную сумму счёта.");
    const lines = (payload.lines || []).map(line => ({ article: String(line.article || "").trim(), name: String(line.name || "").trim(), quantity: this.parseAmount(line.quantity), unit: String(line.unit || "").trim(), price: this.parseAmount(line.price), total: this.parseAmount(line.total) }))
      .filter(line => line.article || line.name);
    const verified = this.verifyLines(order.id, lines);
    const folder = this.ensureInvoiceFolder(order.number);
    const blob = Utilities.newBlob(Utilities.base64Decode(String(file.base64)), file.mimeType || MimeType.PDF, GoogleDriveService.getSafeFileName(file.name));
    const driveFile = folder.createFile(blob);
    const invoiceSheet = this.ensureInvoicesSheet();
    const lineSheet = this.ensureLinesSheet();
    const invoiceId = Utilities.getUuid();
    const now = new Date();
    const verification = verified.verification;
    invoiceSheet.appendRow([invoiceId, order.id, String(payload.number || "").trim(), this.toSheetDate(payload.date), amount, driveFile.getId(), driveFile.getName(), driveFile.getMimeType(), verification, "", verified.reason, now]);
    if (verified.lines.length) {
      lineSheet.getRange(lineSheet.getLastRow() + 1, 1, verified.lines.length, this.LINE_HEADERS.length).setValues(verified.lines.map(line => [Utilities.getUuid(), invoiceId, line.article, line.name, line.quantity === "" ? "" : line.quantity, line.unit, line.price === "" ? "" : line.price, line.total === "" ? "" : line.total, line.matchStatus, line.reason]));
    }
    this.refreshOrderChecks(order.id);
    return this.getOrderSummary(order.id, this.getOrderLines(order.id));
  }

  static attachQuick(orderId, payload) {
    const order = this.assertOrder(orderId);
    const file = payload && payload.file;
    if (!file || !file.name || !file.base64 || !/\.pdf$/i.test(String(file.name))) throw new Error("Выберите PDF-счёт.");
    const folder = this.ensureInvoiceFolder(order.number);
    const blob = Utilities.newBlob(Utilities.base64Decode(String(file.base64)), file.mimeType || MimeType.PDF, GoogleDriveService.getSafeFileName(file.name));
    const driveFile = folder.createFile(blob);
    const invoiceSheet = this.ensureInvoicesSheet();
    const now = new Date();
    const number = String(payload.number || "").trim();
    invoiceSheet.appendRow([Utilities.getUuid(), order.id, number, "", "", driveFile.getId(), driveFile.getName(), driveFile.getMimeType(), "Не проверялся", "", "Быстрое прикрепление без сверки", now]);
    const update = number ? SupplierOrderService.appendInvoiceNumber(order.id, number) : SupplierOrderService.markInvoiceReceived(order.id);
    return { orderId: order.id, fileName: driveFile.getName(), status: update.status, invoiceNumbers: update.invoiceNumbers || order.invoiceNumbers || "" };
  }

  static delete(orderId, invoiceId) {
    const order = this.assertOrder(orderId);
    if (order.status !== "Сформирован") throw new Error("Удалить счёт можно только у сформированного заказа.");
    const sheet = this.ensureInvoicesSheet();
    const rows = sheet.getDataRange().getValues();
    const index = rows.slice(1).findIndex(row => String(row[0] || "") === String(invoiceId || "") && String(row[1] || "") === order.id);
    if (index < 0) throw new Error("Счёт не найден.");
    const row = rows[index + 1];
    const fileId = String(row[5] || "");
    if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (error) { Logger.log("Файл счёта недоступен: " + error.message); } }
    sheet.deleteRow(index + 2);
    const lineSheet = this.ensureLinesSheet();
    const lineRows = lineSheet.getDataRange().getValues();
    lineRows.slice(1).map((line, rowIndex) => ({ line: line, rowNumber: rowIndex + 2 })).filter(item => String(item.line[1] || "") === String(invoiceId || "")).map(item => item.rowNumber).sort((a, b) => b - a).forEach(rowNumber => lineSheet.deleteRow(rowNumber));
    this.refreshOrderChecks(order.id);
    return this.getOrderSummary(order.id, this.getOrderLines(order.id));
  }

  static download(orderId, invoiceId) {
    this.assertOrder(orderId);
    const sheet = this.ensureInvoicesSheet();
    const row = sheet.getDataRange().getValues().slice(1).find(item => String(item[0] || "") === String(invoiceId || "") && String(item[1] || "") === String(orderId || ""));
    if (!row) throw new Error("Счёт не найден.");
    const file = DriveApp.getFileById(String(row[5] || ""));
    if (file.getSize() > 10 * 1024 * 1024) throw new Error("Через приложение доступен файл счёта до 10 МБ.");
    const blob = file.getBlob();
    return { name: file.getName(), mimeType: blob.getContentType() || MimeType.PDF, base64: Utilities.base64Encode(blob.getBytes()) };
  }

  static refreshOrderChecks(orderId) {
    const summary = this.getOrderSummary(orderId, this.getOrderLines(orderId));
    const sheet = this.ensureInvoicesSheet();
    if (sheet.getLastRow() < 2) return summary;
    const rows = sheet.getDataRange().getValues();
    rows.slice(1).map((row, index) => ({ row: row, rowNumber: index + 2 })).filter(item => String(item.row[1] || "") === String(orderId || "")).forEach(item => {
      const lineIssue = String(item.row[10] || "");
      const status = lineIssue || summary.verification === "Совпадает" ? (lineIssue ? "Требует проверки" : "Совпадает") : "Требует проверки";
      sheet.getRange(item.rowNumber, 9, 1, 2).setValues([[status, summary.difference]]);
    });
    return summary;
  }

  static verifyLines(orderId, lines) {
    const orderLines = this.getOrderLines(orderId);
    const normalized = new Map();
    orderLines.forEach(line => {
      const article = this.normalize(line.article);
      const nameKey = this.normalize(line.name) + "|" + this.normalize(line.unit);
      if (article) normalized.set("article:" + article, line);
      normalized.set("name:" + nameKey, line);
    });
    let hasIssue = !lines.length;
    const verifiedLines = lines.map(line => {
      const article = this.normalize(line.article);
      const byArticle = article ? normalized.get("article:" + article) : null;
      const byName = normalized.get("name:" + this.normalize(line.name) + "|" + this.normalize(line.unit));
      const match = byArticle || byName;
      let reason = "";
      if (!match) reason = "Позиция не найдена в заказе";
      else if (line.quantity !== "" && Number(line.quantity) > Number(match.quantity || 0) + 0.0001) reason = "Количество превышает заказ";
      if (reason) hasIssue = true;
      return Object.assign({}, line, { matchStatus: reason ? "Не совпадает" : "Совпадает", reason: reason });
    });
    return { lines: verifiedLines, verification: hasIssue ? "Требует проверки" : "Совпадает", reason: hasIssue ? (verifiedLines.filter(line => line.reason).map(line => line.reason).join("; ") || "Не удалось распознать строки счёта") : "" };
  }

  static getOrderLines(orderId) {
    const sheet = SupplierOrderService.ensureLinesSheet();
    if (sheet.getLastRow() < 2) return [];
    return sheet.getDataRange().getValues().slice(1).filter(row => String(row[1] || "") === String(orderId || "")).map(row => ({ article: row[6], name: row[7], quantity: Number(row[8] || 0), unit: row[9], price: row[10] }));
  }

  static assertOrder(orderId) {
    const order = SupplierOrderService.getOrders().find(item => item.id === String(orderId || "").trim());
    if (!order) throw new Error("Заказ поставщику не найден.");
    return order;
  }

  static ensureInvoiceFolder(orderNumber) {
    const root = DriveApp.getFolderById(GOOGLE_DRIVE_ORDERS_ROOT_ID);
    const roots = root.getFoldersByName("Счета поставщиков");
    const invoicesRoot = roots.hasNext() ? roots.next() : root.createFolder("Счета поставщиков");
    const name = GoogleDriveService.getSafeFileName(String(orderNumber || "Заказ"));
    const matches = invoicesRoot.getFoldersByName(name);
    return matches.hasNext() ? matches.next() : invoicesRoot.createFolder(name);
  }

  static normalize(value) { return String(value || "").toLowerCase().replace(/[^a-zа-я0-9]+/gi, ""); }
  static parseAmount(value) { const normalized = String(value ?? "").replace(/\s/g, "").replace(",", "."); return normalized === "" ? "" : Number(normalized); }
  static normalizeDate(value) { const match = String(value || "").match(/(\d{1,2})[.](\d{1,2})[.](\d{2,4})/); if (!match) return ""; return `${match[3].length === 2 ? "20" + match[3] : match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`; }
  static toSheetDate(value) { const date = this.normalizeDate(value); return date ? new Date(date + "T00:00:00") : ""; }
}
