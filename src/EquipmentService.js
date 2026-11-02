/**
 * Закупка техники: реестр позиций и коммерческие предложения в Google Drive.
 */
class EquipmentService {
  static get HEADERS() {
    return [
      "ID", "Договор", "Клиент", "Наименование техники", "Модель", "Количество",
      "Цена за единицу, ₽", "Менеджер", "Поставщик", "Статус",
      "Ссылка поставщика", "Папка КП", "Создано", "Обновлено"
    ];
  }

  static get STATUSES() {
    return ["Заказать", "Счет", "Оплачен", "Доставлен", "Нет в наличии"];
  }

  static getSnapshot() {
    const sheet = this.ensureSheet();
    const rows = sheet.getDataRange().getValues();
    const items = rows.slice(1).filter(row => row[0]).map(row => ({
      id: String(row[0]),
      contract: String(row[1] || ""),
      client: String(row[2] || ""),
      name: String(row[3] || ""),
      model: String(row[4] || ""),
      quantity: Number(row[5]) || 1,
      unitPrice: row[6] === "" ? "" : Number(row[6]),
      manager: String(row[7] || ""),
      supplier: String(row[8] || ""),
      status: String(row[9] || "Заказать"),
      supplierUrl: String(row[10] || ""),
      proposalFolderUrl: String(row[11] || ""),
      createdAt: row[12] || "",
      updatedAt: row[13] || ""
    }));
    const pairs = this.getOrderPairs();
    const managers = this.getManagers();
    const activeSuppliers = SupplyService.getSupplierRegistry()
      .filter(item => item.active)
      .map(item => String(item.name || "").trim())
      .filter(Boolean);
    const suppliers = [...new Set(activeSuppliers.concat(items.map(item => item.supplier).filter(Boolean)))]
      .sort((a, b) => a.localeCompare(b, "ru"));

    return JSON.parse(JSON.stringify({
      section: "equipment",
      orders: items,
      dictionaries: {
        clients: this.getClients(),
        contracts: this.getContracts(),
        customerContracts: pairs,
        managers: managers,
        suppliers: suppliers,
        activeSuppliers: [...new Set(activeSuppliers)].sort((a, b) => a.localeCompare(b, "ru")),
        equipmentStatuses: this.STATUSES
      }
    }));
  }

  static ensureSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.EQUIPMENT);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.EQUIPMENT);

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, this.HEADERS.length)
        .setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
      const widths = [250, 120, 180, 240, 160, 90, 140, 150, 180, 120, 260, 260, 150, 150];
      widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
      sheet.getRange(2, 6, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat("0");
      sheet.getRange(2, 7, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat("#,##0.00 \"₽\"");
      sheet.getRange(2, 13, Math.max(1, sheet.getMaxRows() - 1), 2).setNumberFormat("dd.MM.yyyy HH:mm");
      sheet.getRange(1, 1, sheet.getMaxRows(), this.HEADERS.length).createFilter();
    } else {
      const actualHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), this.HEADERS.length)).getDisplayValues()[0];
      const mismatch = this.HEADERS.some((header, index) => actualHeaders[index] !== header);
      if (mismatch) {
        throw new Error("Лист «" + CONFIG.SHEETS.EQUIPMENT + "» уже существует, но его заголовки не соответствуют структуре раздела «Техника».");
      }
    }

    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(this.STATUSES, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 10, Math.max(1, sheet.getMaxRows() - 1), 1).setDataValidation(statusRule);
    return sheet;
  }

  static create(data) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const normalized = this.validate(data, null);
      const sheet = this.ensureSheet();
      const id = Utilities.getUuid();
      const now = new Date();
      sheet.appendRow(this.toRow(Object.assign({}, normalized, {
        id: id,
        proposalFolderUrl: "",
        createdAt: now,
        updatedAt: now
      })));
      return { id: id };
    } finally {
      lock.releaseLock();
    }
  }

  static update(id, data) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const sheet = this.ensureSheet();
      const rowNumber = this.findRow(sheet, id);
      const current = this.fromRow(sheet.getRange(rowNumber, 1, 1, this.HEADERS.length).getValues()[0]);
      const normalized = this.validate(data, current);
      const updated = Object.assign({}, normalized, {
        id: current.id,
        proposalFolderUrl: current.proposalFolderUrl,
        createdAt: current.createdAt,
        updatedAt: new Date()
      });
      sheet.getRange(rowNumber, 1, 1, this.HEADERS.length).setValues([this.toRow(updated)]);
      return { id: current.id };
    } finally {
      lock.releaseLock();
    }
  }

  static setStatus(id, status) {
    const value = String(status || "").trim();
    if (!this.STATUSES.includes(value)) throw new Error("Недопустимый статус техники.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const sheet = this.ensureSheet();
      const rowNumber = this.findRow(sheet, id);
      sheet.getRange(rowNumber, 10).setValue(value);
      sheet.getRange(rowNumber, 14).setValue(new Date());
      return { id: String(id), status: value };
    } finally {
      lock.releaseLock();
    }
  }

  static setSupplier(id, supplier) {
    const value = String(supplier || "").trim();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const current = this.getItem(id);
      this.assertSupplier(value, current.supplier);
      const sheet = this.ensureSheet();
      const rowNumber = this.findRow(sheet, id);
      sheet.getRange(rowNumber, 9).setValue(value);
      sheet.getRange(rowNumber, 14).setValue(new Date());
      return { id: String(id), supplier: value };
    } finally {
      lock.releaseLock();
    }
  }

  static validate(data, current) {
    const value = data || {};
    const result = {
      contract: String(value.contract || "").trim(),
      client: String(value.client || "").trim(),
      name: String(value.name || "").trim(),
      model: String(value.model || "").trim(),
      quantity: Number(value.quantity),
      unitPrice: value.unitPrice === "" || value.unitPrice === null || value.unitPrice === undefined
        ? "" : Number(value.unitPrice),
      manager: String(value.manager || "").trim(),
      supplier: String(value.supplier || "").trim(),
      status: String(value.status || "Заказать").trim(),
      supplierUrl: String(value.supplierUrl || "").trim()
    };
    [["Клиент", result.client], ["Договор", result.contract], ["Наименование техники", result.name],
      ["Модель", result.model], ["Менеджер", result.manager]].forEach(item => {
      if (!item[1]) throw new Error("Заполните поле: " + item[0]);
    });
    if (!Number.isInteger(result.quantity) || result.quantity <= 0) {
      throw new Error("Количество должно быть целым положительным числом.");
    }
    if (result.unitPrice !== "" && (!Number.isFinite(result.unitPrice) || result.unitPrice < 0)) {
      throw new Error("Цена за единицу должна быть неотрицательным числом.");
    }
    if (!this.STATUSES.includes(result.status)) throw new Error("Недопустимый статус техники.");
    if (result.supplierUrl && !/^https?:\/\/\S+$/i.test(result.supplierUrl)) {
      throw new Error("Ссылка поставщика должна начинаться с http:// или https://.");
    }
    if (!this.getManagers().includes(result.manager)) throw new Error("Менеджер не найден в справочнике.");
    this.assertSupplier(result.supplier, current ? current.supplier : "");
    return result;
  }

  static assertSupplier(supplier, currentSupplier) {
    if (!supplier) return;
    if (supplier === String(currentSupplier || "").trim()) return;
    const active = SupplyService.getSupplierRegistry()
      .filter(item => item.active)
      .map(item => String(item.name || "").trim());
    if (!active.includes(supplier)) throw new Error("Поставщик не найден среди активных поставщиков.");
  }

  static getOrderPairs() {
    const sheet = this.ensureSheet();
    const rows = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, this.HEADERS.length)
      .getValues().map(row => this.fromRow(row));
    const seen = {};
    return rows.map(row => ({
      contract: String(row.contract || "").trim(),
      client: String(row.client || "").trim()
    })).filter(pair => {
      const key = pair.client + "|" + pair.contract;
      if (!pair.client || !pair.contract || seen[key]) return false;
      seen[key] = true;
      return true;
    }).sort((a, b) => a.client.localeCompare(b.client, "ru") || a.contract.localeCompare(b.contract, "ru"));
  }

  static getClients() {
    const rows = SheetService.objectsFromRows(SheetService.getData(CONFIG.SHEETS.ORDERS));
    return [...new Set(rows.map(row => String(row["Клиент"] || "").trim()).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, "ru"));
  }

  static getContracts() {
    const sheet = this.ensureSheet();
    const values = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
    return [...new Set(values.map(value => String(value || "").trim()).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, "ru"));
  }

  static getManagers() {
    const dictionaries = DictionaryService.getAll();
    return [...new Set([].concat(dictionaries["Менеджер"] || [], dictionaries["Менеджеры"] || [])
      .map(value => String(value || "").trim()).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, "ru"));
  }

  static getItem(id) {
    const sheet = this.ensureSheet();
    return this.fromRow(sheet.getRange(this.findRow(sheet, id), 1, 1, this.HEADERS.length).getValues()[0]);
  }

  static getProposalFiles(id) {
    const item = this.getItem(id);
    if (!item.proposalFolderUrl) return [];
    let folder;
    try {
      folder = DriveApp.getFolderById(this.extractFolderId(item.proposalFolderUrl));
    } catch (error) {
      return [];
    }
    return this.listFiles(folder);
  }

  static uploadProposalFile(id, file) {
    if (!file || !file.name || !file.base64) throw new Error("Не удалось получить файл КП.");
    if (!/\.(pdf|docx?|xlsx?|jpe?g|png|webp)$/i.test(String(file.name))) {
      throw new Error("Поддерживаются КП в PDF, Word, Excel или формате изображения.");
    }
    const bytes = Utilities.base64Decode(String(file.base64));
    if (bytes.length > 10 * 1024 * 1024) throw new Error("Размер файла КП не должен превышать 10 МБ.");
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const item = this.getItem(id);
      const folder = this.ensureProposalFolder(item);
      folder.createFile(Utilities.newBlob(bytes, file.mimeType || "application/octet-stream", this.safeName(file.name)));
      return this.listFiles(folder);
    } finally {
      lock.releaseLock();
    }
  }

  static deleteProposalFile(id, fileId) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const item = this.getItem(id);
      if (!item.proposalFolderUrl) throw new Error("Папка КП не найдена.");
      const folder = DriveApp.getFolderById(this.extractFolderId(item.proposalFolderUrl));
      const iterator = folder.getFiles();
      while (iterator.hasNext()) {
        const file = iterator.next();
        if (file.getId() === String(fileId || "")) {
          file.setTrashed(true);
          return this.listFiles(folder);
        }
      }
      throw new Error("Файл КП не найден.");
    } finally {
      lock.releaseLock();
    }
  }

  static ensureProposalFolder(item) {
    const storedId = this.extractFolderId(item.proposalFolderUrl);
    if (storedId) {
      try { return DriveApp.getFolderById(storedId); } catch (error) { /* recreate below */ }
    }
    const root = DriveApp.getFolderById(GOOGLE_DRIVE_ORDERS_ROOT_ID);
    const equipmentRoot = this.getOrCreateFolder(root, "Закупка техники");
    const pairFolder = this.getOrCreateFolder(equipmentRoot, this.safeName(item.client + " — " + item.contract));
    const itemFolder = this.getOrCreateFolder(pairFolder, this.safeName(item.id));
    const proposalFolder = this.getOrCreateFolder(itemFolder, "КП");
    const sheet = this.ensureSheet();
    const rowNumber = this.findRow(sheet, item.id);
    sheet.getRange(rowNumber, 12).setValue(proposalFolder.getUrl());
    sheet.getRange(rowNumber, 14).setValue(new Date());
    return proposalFolder;
  }

  static getOrCreateFolder(parent, name) {
    const matches = parent.getFoldersByName(name);
    return matches.hasNext() ? matches.next() : parent.createFolder(name);
  }

  static listFiles(folder) {
    const files = [];
    const iterator = folder.getFiles();
    while (iterator.hasNext()) {
      const file = iterator.next();
      files.push({ id: file.getId(), name: file.getName(), url: file.getUrl(), mimeType: file.getMimeType() });
    }
    return files.sort((a, b) => a.name.localeCompare(b.name, "ru"));
  }

  static findRow(sheet, id) {
    const count = Math.max(1, sheet.getLastRow() - 1);
    const found = sheet.getRange(2, 1, count, 1).createTextFinder(String(id || "")).matchEntireCell(true).findNext();
    if (!found) throw new Error("Позиция техники не найдена.");
    return found.getRow();
  }

  static toRow(item) {
    return [item.id, item.contract, item.client, item.name, item.model, item.quantity, item.unitPrice,
      item.manager, item.supplier, item.status, item.supplierUrl, item.proposalFolderUrl,
      item.createdAt, item.updatedAt];
  }

  static fromRow(row) {
    return {
      id: String(row[0] || ""), contract: String(row[1] || ""), client: String(row[2] || ""),
      name: String(row[3] || ""), model: String(row[4] || ""), quantity: Number(row[5]) || 1,
      unitPrice: row[6] === "" ? "" : Number(row[6]), manager: String(row[7] || ""),
      supplier: String(row[8] || ""), status: String(row[9] || "Заказать"),
      supplierUrl: String(row[10] || ""), proposalFolderUrl: String(row[11] || ""),
      createdAt: row[12] || "", updatedAt: row[13] || ""
    };
  }

  static extractFolderId(url) {
    const match = String(url || "").match(/folders\/([\w-]+)/);
    return match ? match[1] : "";
  }

  static safeName(value) {
    return String(value || "Техника").replace(/[\\/\u0000]/g, "_").replace(/\s+/g, " ").trim().slice(0, 180) || "Техника";
  }
}
