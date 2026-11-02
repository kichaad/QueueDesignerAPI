/**
 * Очередь листовых материалов. Источник — PDF-карты раскроя БАЗИС.
 */
class PlateMaterialService {
  static get HEADERS() {
    return [
      "ID", "Договор", "Клиент", "Изделие", "Категория", "Артикул", "Наименование", "Цвет/декор", "Размер листа",
      "Толщина, мм", "Единица", "Расчётное количество листов", "Количество листов к заказу", "Листов вручную", "Поставщик", "Цена", "Статус закупки",
      "Файл источника", "Обновлено"
    ];
  }

  static get STATUSES() { return SupplyService.STATUSES; }

  static ensureSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PLATE_MATERIALS);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.PLATE_MATERIALS);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, this.HEADERS.length).setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
      sheet.getRange(1, 1, 1, this.HEADERS.length).createFilter();
    } else {
      const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), this.HEADERS.length)).getValues()[0];
      if (currentHeaders.join("|") !== this.HEADERS.join("|")) this.migrateSheetColumns(sheet, currentHeaders);
    }
    return sheet;
  }

  static migrateSheetColumns(sheet, oldHeaders) {
    const oldWidth = Math.max(oldHeaders.length, sheet.getLastColumn());
    const oldRows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, oldWidth).getValues() : [];
    const projected = oldRows.map(row => this.HEADERS.map(header => {
      const index = oldHeaders.indexOf(header);
      if (index >= 0) return row[index];
      return header === "Единица" ? "лист" : "";
    }));
    sheet.clearContents();
    sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]);
    if (projected.length) sheet.getRange(2, 1, projected.length, this.HEADERS.length).setValues(projected);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, this.HEADERS.length).setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold");
  }

  static getSnapshot() {
    SheetService.ensureOrderDeadlineColumn();
    const values = this.ensureSheet().getDataRange().getValues();
    const deadlines = SupplyService.getOrderDirectory().reduce((result, order) => {
      result[order.key] = order.deadline || "";
      return result;
    }, {});
    const items = values.slice(1).filter(row => row[0]).map(row => ({
      rowId: String(row[0]), contract: row[1], client: row[2], product: row[3], category: row[4], article: row[5],
      deadline: deadlines[[row[1], row[2], row[3]].join("|")] || "",
      name: row[6], decor: row[7], sheetSize: row[8], thickness: row[9], unit: row[10] || "лист", calculatedSheets: row[11], orderSheets: row[12],
      sheetsManual: row[13] === true, supplier: row[14], price: row[15], purchaseStatus: row[16] || "Новый",
      sourceFile: row[17], updatedAt: row[18]
    }));
    return JSON.parse(JSON.stringify({ section: "plate", orders: items, dictionaries: { plateCategories: ["ЛДСП", "ЛМДФ", "ЛХДФ", "МДФ", "Кромка"], purchaseStatuses: this.STATUSES } }));
  }

  static mergeSpecificationMaterials(materials, sourceName, order) {
    const sheet = this.ensureSheet(); const rows = sheet.getDataRange().getValues(); const now = new Date(); const additions = [];
    (materials || []).forEach(material => {
      const catalog = PlateCatalogService.match(material); const category = material.plateCategory || catalog.category || (SupplyService.isEdgeMaterial(material.name) ? "Кромка" : "");
      if (!category) return;
      const name = String(material.name || "").trim(); const decor = String(material.plateDecor || catalog.decor || "").trim(); const thickness = material.plateThickness || catalog.thickness || this.extractThickness(name);
      const exists = rows.slice(1).some(row => [row[1], row[2], row[3]].join("|") === order.key && String(row[4]) === category &&
        (SupplyService.normalizeName(row[6]) === SupplyService.normalizeName(name) || (decor && SupplyService.normalizeName(row[7]).indexOf(SupplyService.normalizeName(decor)) >= 0)));
      if (exists) return;
      const record = { category: category, article: String(material.article || ""), name: name, decor: decor, thickness: thickness, sheetSize: "", unit: String(material.unit || (category === "Кромка" ? "м" : "лист")), calculatedSheets: "", orderSheets: Number(material.quantity) || 0 };
      additions.push([this.buildRowId(order.key, record), order.contract, order.client, order.product, record.category, record.article, record.name, record.decor, record.sheetSize, record.thickness, record.unit, record.calculatedSheets, record.orderSheets, false, "", "", "Новый", sourceName, now]);
    });
    if (additions.length) sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, this.HEADERS.length).setValues(additions);
    return { added: additions.length };
  }

  static clearOrderMaterials(orderKey) {
    const sheet = this.ensureSheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const rows = sheet.getDataRange().getValues();
      const removed = rows.slice(1).filter(row => [row[1], row[2], row[3]].join("|") === String(orderKey || "")).length;
      if (!removed) return { removed: 0 };
      const remaining = rows.slice(1).filter(row => [row[1], row[2], row[3]].join("|") !== String(orderKey || ""));
      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, this.HEADERS.length).clearContent();
      if (remaining.length) sheet.getRange(2, 1, remaining.length, this.HEADERS.length).setValues(remaining);
      return { removed: removed };
    } finally {
      lock.releaseLock();
    }
  }

  static previewFiles(files, orderKey) {
    const order = orderKey ? SupplyService.getOrderDirectory().find(item => item.key === String(orderKey || "")) : null;
    if (orderKey && !order) throw new Error("Выбранный заказ больше не найден в листе «Изделия».");
    if (!Array.isArray(files) || !files.length) throw new Error("Выберите папку с PDF-картами раскроя.");
    const parsed = [];
    const problems = [];
    const diagnostics = [];
    const texts = [];
    files.forEach(file => {
      if (!file || !/\.pdf$/i.test(String(file.name || ""))) return;
      const status = { file: String(file.name || "PDF"), version: String(file.clientVersion || CONFIG.APP_VERSION), received: true, browserText: false, drivePdf: false, driveDocument: false, textExtracted: false, materials: 0, message: "" };
      try {
        let text = String(file.clientText || "").trim();
        if (text) {
          status.browserText = true;
          status.textExtracted = true;
          status.message = "Текст извлечён в браузере.";
        } else {
          const result = this.readPdfText(file);
          text = result.text;
          status.drivePdf = result.drivePdf;
          status.driveDocument = result.driveDocument;
          status.textExtracted = Boolean(text);
          status.message = `${file.clientReadError ? `Браузерный разбор недоступен: ${file.clientReadError}. ` : ""}${result.message}`;
        }
        texts.push(text);
        const materials = this.extractPlateMaterials(text, file.name);
        status.materials = materials.length;
        if (materials.length) parsed.push.apply(parsed, materials);
        else {
          status.message = "Текст получен, но листовые материалы не распознаны.";
          problems.push({ file: file.name, reason: status.message });
        }
      } catch (error) {
        status.message = String(error.message || error);
        problems.push({ file: status.file, reason: status.message });
      }
      diagnostics.push(status);
    });
    if (!parsed.length) {
      const details = diagnostics.map(item => `${item.file}: ${item.message || "материалы не найдены"}`).join("\n");
      throw new Error(`Не удалось обработать PDF-карты.\n${details}`);
    }
    return { order: order, materials: this.aggregateMaterials(parsed), problems: problems, diagnostics: diagnostics, orderCandidates: this.getImportOrders(texts.join("\n")) };
  }

  static getImportOrders(projectText) {
    const orders = SupplyService.getOrderDirectory();
    const source = SupplyService.compact(projectText);
    const matchedContracts = [...new Set(orders.map(order => order.contract).filter(contract => {
      const key = SupplyService.compact(contract);
      return key && source.indexOf(key) >= 0;
    }))];
    const visibleOrders = matchedContracts.length ? orders.filter(order => matchedContracts.includes(order.contract)) : orders;
    const contracts = [...new Set(visibleOrders.map(order => order.contract))].sort((a, b) => String(a).localeCompare(String(b), "ru"));
    return { contracts: contracts, orders: visibleOrders, detectedContract: matchedContracts.length === 1 ? matchedContracts[0] : "" };
  }

  static readPdfText(file) {
    const bytes = Utilities.base64Decode(String(file.base64 || ""));
    if (!bytes.length || bytes.length > 8 * 1024 * 1024) throw new Error("PDF пуст или превышает 8 МБ.");
    const blob = Utilities.newBlob(bytes, file.mimeType || MimeType.PDF, file.name);
    let pdfId = "";
    let documentId = "";
    try {
      try {
        // 1. Сохраняем исходник именно как PDF — это важно для Drive API.
        pdfId = Drive.Files.insert({ title: "TEMP_PLATE_PDF_" + Date.now(), mimeType: MimeType.PDF }, blob).id;
      } catch (error) {
        throw new Error(`Этап «загрузка PDF в Drive»: ${String(error.message || error)}`);
      }
      try {
        // 2. Только после загрузки создаём отдельный Google Docs без OCR.
        documentId = Drive.Files.copy({ title: "TEMP_PLATE_DOC_" + Date.now(), mimeType: MimeType.GOOGLE_DOCS }, pdfId, { convert: true }).id;
      } catch (error) {
        throw new Error(`Этап «конвертация PDF в Google Docs»: ${String(error.message || error)}`);
      }
      try {
        const text = DocumentApp.openById(documentId).getBody().getText();
        if (!String(text || "").trim()) throw new Error("PDF не содержит распознаваемого текста.");
        return { text: text, drivePdf: true, driveDocument: true, message: "Текст извлечён через двухшаговый Google Drive." };
      } catch (error) {
        throw new Error(`Этап «чтение текста Google Docs»: ${String(error.message || error)}`);
      }
    } catch (error) {
      const message = `Ошибка конвертации Google Drive: ${String(error.message || error)}`;
      console.error(JSON.stringify({ event: "plate_pdf_read_failed", file: file.name, message: message }));
      throw new Error(message);
    } finally {
      if (documentId) {
        try { Drive.Files.trash(documentId); }
        catch (cleanupError) { console.warn(`Не удалось удалить временный PDF-документ ${documentId}: ${cleanupError}`); }
      }
      if (pdfId) {
        try { Drive.Files.trash(pdfId); }
        catch (cleanupError) { console.warn(`Не удалось удалить временный PDF ${pdfId}: ${cleanupError}`); }
      }
    }
  }

  static extractPlateMaterials(text, sourceFile) {
    // Конвертер Drive иногда разрывает аббревиатуры на отдельные буквы
    // ("Л Д С П"). Склеиваем их до разбиения текста на строки.
    const sourceText = this.normalizeOcrText(text);
    const lines = sourceText.replace(/\r/g, "").split(/\n+/).map(line => line.replace(/\s+/g, " ").trim()).filter(Boolean);
    const sourceCategory = this.detectCategory(sourceFile);
    const result = [];
    for (let index = 0; index < lines.length; index++) {
      const category = this.detectCategory(lines[index]);
      if (!category) continue;
      // В выгрузке БАЗИС параметры листа и «Количество плит материала»
      // находятся перед строкой «Материал», а «всего листов» — после неё.
      const block = lines.slice(Math.max(0, index - 8), Math.min(lines.length, index + 4)).join(" ");
      const name = this.extractMaterialName(sourceText, category) || lines[index];
      const sheets = this.extractSheetCount(block);
      result.push({
        category: category,
        article: this.extractArticle(name + " " + sourceText),
        name: name,
        decor: this.extractDecor(name, category),
        sheetSize: this.extractSheetSize(block),
        thickness: this.extractThickness(name), unit: "лист",
        calculatedSheets: sheets,
        orderSheets: sheets,
        sourceFile: sourceFile,
        needsReview: !Number.isFinite(sheets)
      });
    }
    // У карт БАЗИС имя файла уже содержит категорию. Это резервный путь для
    // PDF, у которых Google Drive распознал таблицу с переносами строк.
    if (!result.length && sourceCategory) {
      const materialIndex = lines.findIndex(line => /материал|артикул/i.test(line));
      const name = this.extractMaterialName(sourceText, sourceCategory) || (materialIndex >= 0
        ? lines.slice(Math.max(0, materialIndex - 1), Math.min(lines.length, materialIndex + 3)).join(" ")
        : String(sourceFile || "").replace(/\.pdf$/i, ""));
      result.push({
        category: sourceCategory,
        article: this.extractArticle(sourceText),
        name: name,
        decor: this.extractDecor(name, sourceCategory),
        sheetSize: this.extractSheetSize(sourceText),
        thickness: this.extractThickness(name), unit: "лист",
        calculatedSheets: this.extractSheetCount(sourceText),
        orderSheets: this.extractSheetCount(sourceText),
        sourceFile: sourceFile,
        needsReview: !Number.isFinite(this.extractSheetCount(sourceText))
      });
    }
    // Один материал повторяется на каждой карте раскроя. Внутри одного PDF
    // берём итог «всего листов», а не складываем повторяющиеся страницы.
    const unique = {};
    result.forEach(material => {
      const key = [material.category, SupplyService.normalizeArticle(material.article), SupplyService.normalizeName(material.decor), material.thickness, material.sheetSize].join("|");
      if (!unique[key] || (Number.isFinite(material.calculatedSheets) && material.calculatedSheets > unique[key].calculatedSheets)) unique[key] = material;
    });
    return Object.keys(unique).map(key => unique[key]);
  }

  static detectCategory(value) {
    const text = this.normalizeOcrText(value).toUpperCase();
    if (/(^|[^А-ЯA-Z])ЛДСП($|[^А-ЯA-Z])/.test(text)) return "ЛДСП";
    if (/(^|[^А-ЯA-Z])ЛМДФ($|[^А-ЯA-Z])|(^|[^А-ЯA-Z])МДФ($|[^А-ЯA-Z])/.test(text)) return "МДФ";
    if (/(^|[^А-ЯA-Z])ЛХДФ($|[^А-ЯA-Z])|(^|[^А-ЯA-Z])ХДФ($|[^А-ЯA-Z])/.test(text)) return "ХДФ";
    return "";
  }

  static normalizeOcrText(value) {
    return String(value || "")
      .replace(/Л\s*Д\s*С\s*П/gi, "ЛДСП")
      .replace(/Л\s*М\s*Д\s*Ф/gi, "ЛМДФ")
      .replace(/Л\s*Х\s*Д\s*Ф/gi, "ЛХДФ")
      .replace(/М\s*Д\s*Ф/gi, "МДФ")
      .replace(/Х\s*Д\s*Ф/gi, "ХДФ")
      .replace(/А\s*Р\s*Т\s*И\s*К\s*У\s*Л/gi, "Артикул")
      .replace(/М\s*А\s*Т\s*Е\s*Р\s*И\s*А\s*Л/gi, "Материал")
      // PDF.js иногда склеивает соседние текстовые блоки: «МатериалМДФ».
      .replace(/Материал(?=(?:ЛДСП|ЛМДФ|МДФ|ЛХДФ|ХДФ))/gi, "Материал ")
      .replace(/К\s*О\s*Л\s*И\s*Ч\s*Е\s*С\s*Т\s*В\s*О/gi, "Количество");
  }

  static extractArticle(text) {
    const value = String(text || "");
    const explicit = value.match(/артикул\s*[:№]?\s*(\d{4,8})/i);
    if (explicit) return explicit[1];
    const match = value.match(/\b([A-ZА-Я]\d{3,4}(?:\s*[A-ZА-Я]{1,3}\d{0,2})?|\d{4,8})\b/i);
    return match ? match[1].replace(/\s+/g, " ") : "";
  }

  static extractMaterialName(text, category) {
    const categoryPattern = category === "МДФ" ? "(?:ЛМДФ|МДФ)" : category === "ХДФ" ? "(?:ЛХДФ|ХДФ)" : "ЛДСП";
    const match = this.normalizeOcrText(text).match(new RegExp(`Материал\\s+(${categoryPattern}[^\\n]{0,280}?)(?=\\s*\\(?\\s*Артикул(?:\\s|\\)|$)|\\s+Заказ(?:\\s|$)|\\s+Лист\\s*\\d|$)`, "i"));
    return match ? match[1].replace(/\\s+/g, " ").trim() : "";
  }

  static extractThickness(text) {
    const match = String(text || "").match(/(\d{1,2})\s*мм/i);
    return match ? Number(match[1]) : "";
  }

  static extractSheetSize(text) {
    const value = String(text || "");
    const full = value.match(/\b(\d{1,2}\s*[xх*]\s*\d{3,4}\s*[xх*]\s*\d{3,4})\b/i);
    if (full) return full[1].replace(/\s+/g, "");
    const dimensions = [...value.matchAll(/(\d{3,4}\s*[xх*]\s*\d{3,4})/ig)]
      .map(match => match[1].replace(/\s+/g, ""))
      .filter(item => item.split(/[xх*]/i).every(size => Number(size) >= 1000));
    return dimensions.length ? dimensions[0] : "";
  }

  static extractSheetCount(text) {
    const value = String(text || "");
    // В карте есть десятки количеств деталей. Берём только явно подписанное
    // итоговое число листов, а не наибольшее число деталей в раскрое.
    const total = [...value.matchAll(/всего\s+лист(?:ов|а)?\s*[:=]?\s*(\d+(?:[,.]\d+)?)/ig)];
    if (total.length) return Math.max.apply(null, total.map(match => Number(String(match[1]).replace(",", "."))));
    const plates = [...value.matchAll(/кол(?:ичество|[- ]?во)?\s+плит(?:ы)?\s+материала\s*[:=]?\s*(\d+(?:[,.]\d+)?)/ig)];
    if (plates.length) return Math.max.apply(null, plates.map(match => Number(String(match[1]).replace(",", "."))));
    return NaN;
  }

  static extractDecor(text, category) {
    let value = String(text || "");
    const materialWords = [...value.matchAll(/материал/ig)];
    if (materialWords.length) value = value.slice(materialWords[materialWords.length - 1].index);
    value = value.split(/\(?\s*артикул|заказ|лист\s*\d+/i)[0];
    value = value.replace(new RegExp(category === "МДФ" ? "Л?МДФ" : category === "ХДФ" ? "Л?ХДФ" : category, "ig"), "");
    value = value.replace(/материал/ig, "").replace(/(?:egger|kastamonu|кроношпан|kronospan)/ig, "");
    // Коды декора (U216, ST9, 101 PE) — часть цвета и нужны для заказа.
    value = value.replace(/\d{1,2}\s*мм/ig, "").replace(/\d{1,2}\s*[xх*]\s*\d{3,4}\s*[xх*]\s*\d{3,4}/ig, "");
    // Это не декор, а технические пометки к МДФ.
    value = value.replace(/\(?\s*(?:голый|2\s*ст)\s*\)?/ig, "");
    value = value.replace(/(?:кол(?:ичество|[- ]?во)?\s*(?:лист(?:ов|а)?)?|лист(?:ов|а)?|шт\.?)\s*[:=]?\s*\d+(?:[,.]\d+)?/ig, "");
    return value.replace(/\s+/g, " ").replace(/^[-:;, ]+|[-:;, ]+$/g, "").slice(0, 160);
  }

  static aggregateMaterials(materials) {
    const grouped = {};
    materials.forEach(material => {
      const key = [material.category, SupplyService.normalizeArticle(material.article), SupplyService.normalizeName(material.decor), material.thickness, material.sheetSize].join("|");
      if (!grouped[key]) grouped[key] = Object.assign({}, material);
      else {
        const previous = grouped[key];
        if (Number.isFinite(material.calculatedSheets)) previous.calculatedSheets = (Number.isFinite(previous.calculatedSheets) ? previous.calculatedSheets : 0) + material.calculatedSheets;
        previous.orderSheets = previous.calculatedSheets;
        previous.needsReview = previous.needsReview || material.needsReview;
        previous.sourceFile += "; " + material.sourceFile;
      }
    });
    return Object.keys(grouped).map(key => grouped[key]).sort((a, b) => a.category.localeCompare(b.category, "ru") || a.decor.localeCompare(b.decor, "ru"));
  }

  static saveMaterials(materials, sourceName, orderKey) {
    const order = SupplyService.getOrderDirectory().find(item => item.key === String(orderKey || ""));
    if (!order) throw new Error("Выбранный заказ больше не найден в листе «Изделия».");
    const clean = (materials || []).map(material => ({
      category: this.detectCategory(material.category), article: String(material.article || "").trim(), name: String(material.name || "").trim(),
      decor: String(material.decor || "").trim(), sheetSize: String(material.sheetSize || "").trim(), thickness: this.extractThickness(material.thickness || material.name), unit: String(material.unit || "лист").trim() || "лист",
      calculatedSheets: Number(material.calculatedSheets), orderSheets: Number(material.orderSheets), sourceFile: String(material.sourceFile || sourceName || "").trim()
    })).filter(material => material.category && material.name && Number.isFinite(material.orderSheets) && material.orderSheets >= 0);
    if (!clean.length) throw new Error("Нет корректных строк плитного материала для импорта.");
    const sheet = this.ensureSheet();
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      const rows = sheet.getDataRange().getValues();
      const existing = {};
      rows.slice(1).forEach(row => { if ([row[1], row[2], row[3]].join("|") === order.key) existing[String(row[0])] = row; });
      const now = new Date();
      const newRows = clean.map(material => {
        const rowId = this.buildRowId(order.key, material);
        const old = existing[rowId];
        const manualSheets = Boolean(old && old[13] === true);
        return [rowId, order.contract, order.client, order.product, material.category, material.article, material.name, material.decor, material.sheetSize,
          material.thickness, material.unit, Number.isFinite(material.calculatedSheets) ? material.calculatedSheets : "", manualSheets ? old[12] : material.orderSheets, manualSheets,
          old ? old[14] : "", old ? old[15] : "", old ? (old[16] || "Новый") : "Новый", material.sourceFile, now];
      });
      const nextIds = new Set(newRows.map(row => String(row[0])));
      const others = rows.slice(1).filter(row => [row[1], row[2], row[3]].join("|") !== order.key);
      const finalRows = others.concat(newRows);
      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, this.HEADERS.length).clearContent();
      if (finalRows.length) sheet.getRange(2, 1, finalRows.length, this.HEADERS.length).setValues(finalRows);
      return { imported: newRows.length, needsReview: clean.filter(item => !Number.isFinite(item.calculatedSheets)).map(item => item.name) };
    } finally { lock.releaseLock(); }
  }

  static setOrderSheets(rowId, value) {
    const sheets = Number(value);
    if (!Number.isFinite(sheets) || sheets < 0) throw new Error("Укажите корректное количество листов.");
    const sheet = this.ensureSheet();
    const row = this.findRow(sheet, rowId);
    sheet.getRange(row, 13).setValue(sheets);
    sheet.getRange(row, 14).setValue(true);
    sheet.getRange(row, 19).setValue(new Date());
    return { rowId: rowId, orderSheets: sheets };
  }

  static setStatus(rowId, status) {
    const value = String(status || "").trim();
    if (!this.STATUSES.includes(value)) throw new Error("Недопустимый статус закупки.");
    const sheet = this.ensureSheet();
    const row = this.findRow(sheet, rowId);
    sheet.getRange(row, 17).setValue(value);
    sheet.getRange(row, 19).setValue(new Date());
    return { rowId: rowId, status: value };
  }

  static findRow(sheet, rowId) {
    const found = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).createTextFinder(String(rowId)).matchEntireCell(true).findNext();
    if (!found) throw new Error("Позиция плитного материала не найдена.");
    return found.getRow();
  }

  static buildRowId(orderKey, material) {
    const key = [orderKey, material.category, SupplyService.normalizeArticle(material.article), SupplyService.normalizeName(material.decor), material.thickness, material.sheetSize].join("|");
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, key, Utilities.Charset.UTF_8);
    return bytes.slice(0, 12).map(value => (value + 256).toString(16).slice(-2)).join("");
  }
}
