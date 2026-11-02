/** Регистрация заказов на столешницы и коммерческих предложений. */
class TabletopService {
  static get HEADERS() {
    return ["ID", "Договор", "Клиент", "Наименование столешницы", "Менеджер", "Поставщик", "Статус", "Ссылка поставщика", "Стоимость КП, ₽", "Папка КП", "Создано", "Обновлено"];
  }

  static get STATUSES() { return ["Заказать", "Счет", "Оплачен", "Доставлен"]; }

  static getSnapshot() {
    const sheet = this.ensureSheet();
    const rows = sheet.getDataRange().getValues();
    const items = rows.slice(1).filter(row => row[0]).map(row => this.fromRow(row));
    const activeSuppliers = SupplyService.getSupplierRegistry().filter(item => item.active)
      .map(item => String(item.name || "").trim()).filter(Boolean);
    return JSON.parse(JSON.stringify({
      section: "tabletops",
      orders: items,
      dictionaries: {
        clients: this.getClients(),
        contracts: [...new Set(items.map(item => item.contract).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru")),
        suppliers: [...new Set(activeSuppliers.concat(items.map(item => item.supplier).filter(Boolean)))].sort((a, b) => a.localeCompare(b, "ru")),
        activeSuppliers: [...new Set(activeSuppliers)].sort((a, b) => a.localeCompare(b, "ru")),
        tabletopStatuses: this.STATUSES
      }
    }));
  }

  static ensureSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TABLETOPS);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.TABLETOPS);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, this.HEADERS.length).setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
      [250, 120, 180, 280, 150, 180, 120, 260, 150, 260, 150, 150].forEach((width, index) => sheet.setColumnWidth(index + 1, width));
      sheet.getRange(2, 9, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat("#,##0.00 \"₽\"");
      sheet.getRange(2, 11, Math.max(1, sheet.getMaxRows() - 1), 2).setNumberFormat("dd.MM.yyyy HH:mm");
      sheet.getRange(1, 1, sheet.getMaxRows(), this.HEADERS.length).createFilter();
    } else {
      const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), this.HEADERS.length)).getDisplayValues()[0];
      if (this.HEADERS.some((header, index) => headers[index] !== header)) {
        throw new Error("Лист «" + CONFIG.SHEETS.TABLETOPS + "» уже существует, но его заголовки не соответствуют структуре раздела «Столешницы».");
      }
    }
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(this.STATUSES, true).setAllowInvalid(false).build();
    sheet.getRange(2, 7, Math.max(1, sheet.getMaxRows() - 1), 1).setDataValidation(rule);
    return sheet;
  }

  static create(data) {
    const lock = LockService.getDocumentLock(); lock.waitLock(30000);
    try {
      const item = this.validate(data, null);
      const now = new Date(); const id = Utilities.getUuid(); const sheet = this.ensureSheet();
      sheet.appendRow(this.toRow(Object.assign(item, { id: id, proposalFolderUrl: "", createdAt: now, updatedAt: now })));
      return { id: id };
    } finally { lock.releaseLock(); }
  }

  static update(id, data) {
    const lock = LockService.getDocumentLock(); lock.waitLock(30000);
    try {
      const sheet = this.ensureSheet(); const row = this.findRow(sheet, id);
      const current = this.fromRow(sheet.getRange(row, 1, 1, this.HEADERS.length).getValues()[0]);
      const item = this.validate(data, current);
      sheet.getRange(row, 1, 1, this.HEADERS.length).setValues([this.toRow(Object.assign(item, { id: current.id, proposalFolderUrl: current.proposalFolderUrl, createdAt: current.createdAt, updatedAt: new Date() }))]);
      return { id: current.id };
    } finally { lock.releaseLock(); }
  }

  static setStatus(id, status) {
    const value = String(status || "").trim(); if (!this.STATUSES.includes(value)) throw new Error("Недопустимый статус столешницы.");
    return this.updateQuick(id, 7, value, "status");
  }

  static setSupplier(id, supplier) {
    const value = String(supplier || "").trim(); const current = this.getItem(id); this.assertSupplier(value, current.supplier);
    return this.updateQuick(id, 6, value, "supplier");
  }

  static updateQuick(id, column, value, field) {
    const lock = LockService.getDocumentLock(); lock.waitLock(30000);
    try { const sheet = this.ensureSheet(); const row = this.findRow(sheet, id); sheet.getRange(row, column).setValue(value); sheet.getRange(row, 12).setValue(new Date()); return { id: String(id), [field]: value }; }
    finally { lock.releaseLock(); }
  }

  static validate(data, current) {
    const source = data || {};
    const item = { contract: String(source.contract || "").trim(), client: String(source.client || "").trim(), name: String(source.name || "").trim(), manager: String(source.manager || "").trim(), supplier: String(source.supplier || "").trim(), status: String(source.status || "Заказать").trim(), supplierUrl: String(source.supplierUrl || "").trim(), quoteCost: source.quoteCost === "" || source.quoteCost === null || source.quoteCost === undefined ? "" : Number(source.quoteCost) };
    [["Клиент", item.client], ["Договор", item.contract], ["Наименование столешницы", item.name], ["Менеджер", item.manager]].forEach(([label, value]) => { if (!value) throw new Error("Заполните поле: " + label); });
    if (!this.STATUSES.includes(item.status)) throw new Error("Недопустимый статус столешницы.");
    if (item.quoteCost !== "" && (!Number.isFinite(item.quoteCost) || item.quoteCost < 0)) throw new Error("Стоимость КП должна быть неотрицательным числом.");
    if (item.supplierUrl && !/^https?:\/\/\S+$/i.test(item.supplierUrl)) throw new Error("Ссылка поставщика должна начинаться с http:// или https://.");
    this.assertSupplier(item.supplier, current ? current.supplier : "");
    return item;
  }

  static assertSupplier(supplier, currentSupplier) {
    if (!supplier || supplier === String(currentSupplier || "").trim()) return;
    const active = SupplyService.getSupplierRegistry().filter(item => item.active).map(item => String(item.name || "").trim());
    if (!active.includes(supplier)) throw new Error("Поставщик не найден среди активных поставщиков.");
  }

  static getClients() {
    const rows = SheetService.objectsFromRows(SheetService.getData(CONFIG.SHEETS.ORDERS));
    return [...new Set(rows.map(row => String(row["Клиент"] || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru"));
  }

  static getItem(id) { const sheet = this.ensureSheet(); return this.fromRow(sheet.getRange(this.findRow(sheet, id), 1, 1, this.HEADERS.length).getValues()[0]); }
  static getProposalFiles(id) { const item = this.getItem(id); if (!item.proposalFolderUrl) return []; try { return this.listFiles(DriveApp.getFolderById(this.extractFolderId(item.proposalFolderUrl))); } catch (error) { return []; } }

  static uploadProposalFile(id, file) {
    if (!file || !file.name || !file.base64) throw new Error("Не удалось получить файл КП.");
    if (!/\.(pdf|docx?|xlsx?|jpe?g|png|webp)$/i.test(String(file.name))) throw new Error("Поддерживаются КП в PDF, Word, Excel или формате изображения.");
    const bytes = Utilities.base64Decode(String(file.base64)); if (bytes.length > 10 * 1024 * 1024) throw new Error("Размер файла КП не должен превышать 10 МБ.");
    const lock = LockService.getDocumentLock(); lock.waitLock(30000);
    try { const folder = this.ensureProposalFolder(this.getItem(id)); folder.createFile(Utilities.newBlob(bytes, file.mimeType || "application/octet-stream", this.safeName(file.name))); return this.listFiles(folder); }
    finally { lock.releaseLock(); }
  }

  static deleteProposalFile(id, fileId) {
    const lock = LockService.getDocumentLock(); lock.waitLock(30000);
    try {
      const item = this.getItem(id); if (!item.proposalFolderUrl) throw new Error("Папка КП не найдена.");
      const folder = DriveApp.getFolderById(this.extractFolderId(item.proposalFolderUrl)); const files = folder.getFiles();
      while (files.hasNext()) { const file = files.next(); if (file.getId() === String(fileId || "")) { file.setTrashed(true); return this.listFiles(folder); } }
      throw new Error("Файл КП не найден.");
    } finally { lock.releaseLock(); }
  }

  static ensureProposalFolder(item) {
    const stored = this.extractFolderId(item.proposalFolderUrl); if (stored) { try { return DriveApp.getFolderById(stored); } catch (error) {} }
    const root = DriveApp.getFolderById(GOOGLE_DRIVE_ORDERS_ROOT_ID);
    const category = this.getOrCreateFolder(root, "Столешницы"); const pair = this.getOrCreateFolder(category, this.safeName(item.client + " — " + item.contract));
    const rowFolder = this.getOrCreateFolder(pair, this.safeName(item.id)); const proposal = this.getOrCreateFolder(rowFolder, "КП");
    const sheet = this.ensureSheet(); const row = this.findRow(sheet, item.id); sheet.getRange(row, 10).setValue(proposal.getUrl()); sheet.getRange(row, 12).setValue(new Date());
    return proposal;
  }

  static getOrCreateFolder(parent, name) { const folders = parent.getFoldersByName(name); return folders.hasNext() ? folders.next() : parent.createFolder(name); }
  static listFiles(folder) { const result = []; const files = folder.getFiles(); while (files.hasNext()) { const file = files.next(); result.push({ id: file.getId(), name: file.getName(), url: file.getUrl(), mimeType: file.getMimeType() }); } return result.sort((a, b) => a.name.localeCompare(b.name, "ru")); }
  static findRow(sheet, id) { const found = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).createTextFinder(String(id || "")).matchEntireCell(true).findNext(); if (!found) throw new Error("Заказ на столешницу не найден."); return found.getRow(); }
  static toRow(item) { return [item.id, item.contract, item.client, item.name, item.manager, item.supplier, item.status, item.supplierUrl, item.quoteCost, item.proposalFolderUrl, item.createdAt, item.updatedAt]; }
  static fromRow(row) { return { id: String(row[0] || ""), contract: String(row[1] || ""), client: String(row[2] || ""), name: String(row[3] || ""), manager: String(row[4] || ""), supplier: String(row[5] || ""), status: String(row[6] || "Заказать"), supplierUrl: String(row[7] || ""), quoteCost: row[8] === "" ? "" : Number(row[8]), proposalFolderUrl: String(row[9] || ""), createdAt: row[10] || "", updatedAt: row[11] || "" }; }
  static extractFolderId(url) { const match = String(url || "").match(/folders\/([\w-]+)/); return match ? match[1] : ""; }
  static safeName(value) { return String(value || "Столешницы").replace(/[\\/\u0000]/g, "_").replace(/\s+/g, " ").trim().slice(0, 180) || "Столешницы"; }
}
