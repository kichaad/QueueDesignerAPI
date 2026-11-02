/**
 * Формирует производственную спецификацию из XML БАЗИС-Мебельщик.
 * Исходный XML и результирующая книга используются только во время запроса.
 */
class BasisSpecificationService {
  static createExcel(file) {
    if (!file || !file.name || !file.base64) throw new Error("Выберите XML-файл БАЗИС.");
    if (!/\.xml$/i.test(String(file.name))) throw new Error("Для спецификации нужен файл XML.");

    const bytes = Utilities.base64Decode(String(file.base64));
    const maxBytes = 10 * 1024 * 1024;
    if (bytes.length > maxBytes) throw new Error("XML-файл больше 10 МБ.");

    let root;
    try {
      root = XmlService.parse(Utilities.newBlob(bytes, "application/xml", String(file.name)).getDataAsString("UTF-8")).getRootElement();
    } catch (error) {
      throw new Error("Не удалось прочитать XML БАЗИС: " + error.message);
    }
    if (root.getName() !== "Проект") throw new Error("В XML не найден корневой элемент «Проект».");
    const product = root.getChild("Изделие");
    if (!product) throw new Error("В XML БАЗИС не найдено изделие.");

    const data = this.parseProduct(root, product);
    if (!data.sections.length) throw new Error("В XML изделия не найдены секции.");
    return this.createWorkbook(data);
  }

  static parseProduct(root, product) {
    const projectName = this.text(root, "Наименование") || root.getAttribute("Наименование")?.getValue() || "Проект";
    const productName = this.text(product, "Наименование") || "Изделие";
    const sections = [];
    const panels = [];
    const hardware = [];
    const productList = product.getChild("СписокЭлементов");

    (productList ? productList.getChildren() : []).forEach(sectionNode => {
      if (!this.isStructural(sectionNode)) return;
      const section = { name: this.text(sectionNode, "Наименование") || "Без секции", units: [] };
      sections.push(section);
      const sectionList = sectionNode.getChild("СписокЭлементов");
      (sectionList ? sectionList.getChildren() : []).forEach(child => {
        if (this.isAssemblyUnit(child)) {
          const unit = { id: "", name: this.text(child, "Наименование") || "Без названия", section: section.name, panels: [], hardware: [] };
          section.units.push(unit);
          this.walk(child, section.name, unit, panels, hardware);
        } else {
          this.walk(child, section.name, null, panels, hardware);
        }
      });
    });

    let unitNumber = 0;
    sections.forEach(section => section.units.forEach(unit => { unitNumber += 1; unit.id = unitNumber; }));
    return { projectName: projectName, productName: productName, sections: sections, panels: panels, hardware: hardware };
  }

  static walk(node, sectionName, unit, panels, hardware) {
    const type = this.text(node, "ТипОбъекта");
    if (type === "Панель") {
      const panel = this.readPanel(node, sectionName, unit);
      panels.push(panel);
      if (unit) unit.panels.push(panel);
    } else if (type === "Фурнитура") {
      const item = this.readHardware(node);
      hardware.push(item);
      if (unit) unit.hardware.push(item);
    }
    const children = node.getChild("СписокЭлементов");
    (children ? children.getChildren() : []).forEach(child => this.walk(child, sectionName, unit, panels, hardware));
  }

  static readPanel(node, sectionName, unit) {
    const material = this.material(node);
    return {
      position: this.text(node, "Позиция") || this.text(node, "КодДетали") || "—",
      name: this.text(node, "Наименование") || "Без наименования",
      materialType: this.materialType(material.name),
      section: sectionName,
      unit: unit ? unit.name : "ВНЕ СЕ",
      quantity: this.numberOrText(this.text(node, "Количество")),
      length: this.numberOrText(this.text(node, "Длина")),
      width: this.numberOrText(this.text(node, "Ширина")),
      thickness: this.numberOrText(this.text(node, "ОбщаяТолщина") || material.thickness),
      sourceIndex: 0
    };
  }

  static readHardware(node) {
    const material = this.material(node);
    const objectCode = this.text(node, "Код");
    return {
      article: material.code || objectCode || "—",
      name: material.name || this.text(node, "Наименование") || "Без наименования",
      unit: material.unit || "—",
      quantity: this.numberOrText(material.quantity || this.text(node, "Количество")),
      materialId: material.id || "",
      thickness: material.thickness || "",
      manufacturer: material.manufacturer || ""
    };
  }

  static material(node) {
    const material = node.getChild("ОсновнойМатериал");
    if (!material) return { present: false, name: "", code: "", unit: "", modelUnit: "", quantity: "", id: "", thickness: "", manufacturer: "" };
    return {
      present: true,
      name: this.text(material, "Наименование"),
      code: this.text(material, "Код"),
      unit: this.text(material, "ЕдИзм"),
      modelUnit: this.text(material, "ЕдИзмСМодели"),
      quantity: this.text(material, "Количество"),
      id: this.text(material, "ID"),
      thickness: this.text(material, "Толщина"),
      manufacturer: this.text(material, "Примечание")
    };
  }

  static isStructural(node) {
    return Boolean(node && node.getChild("СписокЭлементов")) && !["Панель", "Фурнитура", "Составная фурнитура"].includes(this.text(node, "ТипОбъекта"));
  }

  static isAssemblyUnit(node) {
    return this.isStructural(node) && this.text(node, "ТипОбъекта") !== "Составная фурнитура";
  }

  static materialType(name) {
    const value = String(name || "").toUpperCase();
    if (value.includes("ЛДСП")) return "ЛДСП";
    if (value.includes("МДФ")) return "МДФ";
    return "Прочие";
  }

  static sortPanels(panels) {
    const rank = { "ЛДСП": 0, "МДФ": 1, "Прочие": 2 };
    return panels.map((item, index) => Object.assign({}, item, { sourceIndex: index }))
      .sort((left, right) => (rank[left.materialType] - rank[right.materialType]) || left.sourceIndex - right.sourceIndex);
  }

  static uniqueDetailPanels(panels) {
    const seen = {};
    return panels.filter(panel => {
      // Лист «Детали» в исходном отчёте БАЗИС содержит одну строку на
      // позицию внутри сборочной единицы. Повторные экземпляры остаются в
      // отчёте СЕ, но здесь не дублируют спецификацию.
      const key = [panel.section, panel.unit, panel.position].join("\u0001");
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
  }

  static createWorkbook(data) {
    const xlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    const safePart = value => String(value || "Спецификация").replace(/[\\/:*?\"<>|]+/g, "_").replace(/\s+/g, " ").trim().slice(0, 80);
    const spreadsheet = SpreadsheetApp.create("Временная спецификация БАЗИС " + safePart(data.productName));
    const spreadsheetId = spreadsheet.getId();
    try {
      const report = spreadsheet.getSheets()[0];
      report.setName("Отчёт");
      const details = spreadsheet.insertSheet("Детали");
      const hardware = spreadsheet.insertSheet("Фурнитура");
      this.writeReport(report, data);
      this.writeDetails(details, data.panels);
      this.writeHardware(hardware, data.hardware);
      SpreadsheetApp.flush();
      const response = UrlFetchApp.fetch(
        "https://www.googleapis.com/drive/v3/files/" + spreadsheetId + "/export?mimeType=" + encodeURIComponent(xlsxMimeType),
        { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
      );
      if (response.getResponseCode() !== 200) throw new Error("Не удалось сформировать Excel-файл: " + response.getContentText().slice(0, 300));
      return {
        filename: "Спецификация_БАЗИС_" + safePart(data.productName) + "_" + stamp + ".xlsx",
        mimeType: xlsxMimeType,
        base64: Utilities.base64Encode(this.configureExcelPrint(response.getBlob()).getBytes())
      };
    } finally {
      try { Drive.Files.trash(spreadsheetId); } catch (error) { console.warn("Не удалось удалить временную спецификацию: " + error.message); }
    }
  }

  // В XLSX параметры страницы хранятся в XML каждого листа, а не в данных ячеек.
  static configureExcelPrint(blob) {
    try {
      const files = Utilities.unzip(blob.copyBlob().setContentType("application/zip"));
      let count = 0;
      const updated = files.map(file => {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(file.getName())) return file;
        count += 1;
        return Utilities.newBlob(this.setWorksheetPrintXml(file.getDataAsString("UTF-8")), "application/xml", file.getName());
      });
      if (count !== 3) throw new Error("В книге должно быть три листа; найдено: " + count);
      return Utilities.zip(updated, "specification.xlsx").setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    } catch (error) {
      throw new Error("Не удалось настроить печать Excel на A4: " + error.message);
    }
  }

  static setWorksheetPrintXml(xml) {
    const text = String(xml || "");
    const document = XmlService.parse(text);
    const root = document.getRootElement();
    if (root.getName() !== "worksheet") throw new Error("Некорректный XML листа Excel.");
    const ns = root.getNamespace();
    const properties = root.getChild("sheetPr", ns) || XmlService.createElement("sheetPr", ns);
    if (!root.getChild("sheetPr", ns)) root.addContent(0, properties);
    let setupProperties = properties.getChild("pageSetUpPr", ns);
    if (!setupProperties) {
      setupProperties = XmlService.createElement("pageSetUpPr", ns);
      const contents = properties.getAllContent();
      const extIndex = contents.findIndex(item => item.getType() === XmlService.ContentTypes.ELEMENT && item.asElement().getName() === "extLst");
      if (extIndex < 0) properties.addContent(setupProperties);
      else properties.addContent(extIndex, setupProperties);
    }
    setupProperties.setAttribute("fitToPage", "1");
    let pageSetup = root.getChild("pageSetup", ns);
    if (!pageSetup) {
      pageSetup = XmlService.createElement("pageSetup", ns);
      const following = new Set(["headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]);
      const content = root.getAllContent();
      const index = content.findIndex(item => item.getType() === XmlService.ContentTypes.ELEMENT && following.has(item.asElement().getName()));
      if (index < 0) root.addContent(pageSetup);
      else root.addContent(index, pageSetup);
    }
    pageSetup.setAttribute("paperSize", "9");
    pageSetup.setAttribute("orientation", "landscape");
    pageSetup.setAttribute("fitToWidth", "1");
    pageSetup.setAttribute("fitToHeight", "0");
    pageSetup.removeAttribute("scale");
    return XmlService.getCompactFormat().format(document);
  }

  static writeReport(sheet, data) {
    const summaryColumns = ["№ СЕ", "Секция", "Сборочная единица", "Панелей", "Позиций фурнитуры"];
    sheet.getRange(1, 1, 1, summaryColumns.length).merge().setValue("БАЗИС — " + data.projectName + " / " + data.productName);
    this.styleTitle(sheet.getRange(1, 1, 1, summaryColumns.length));
    sheet.getRange(3, 1, 1, summaryColumns.length).setValues([summaryColumns]);
    this.styleHeader(sheet.getRange(3, 1, 1, summaryColumns.length));
    const units = data.sections.reduce((result, section) => result.concat(section.units), []);
    if (units.length) {
      sheet.getRange(4, 1, units.length, summaryColumns.length).setValues(units.map(unit => [unit.id, unit.section, unit.name, unit.panels.length, unit.hardware.length]));
      this.styleRows(sheet.getRange(4, 1, units.length, summaryColumns.length));
    }
    let row = 4 + units.length + 2;
    const unitPanels = new Set(units.reduce((result, unit) => result.concat(unit.panels), []));
    const outsidePanels = data.panels.filter(panel => !unitPanels.has(panel));
    const reportGroups = units.slice();
    if (outsidePanels.length) reportGroups.push({ name: "ВНЕ СЕ", panels: outsidePanels, hardware: [], outside: true });
    reportGroups.forEach(unit => {
      const panels = this.sortPanels(unit.panels);
      const hardware = this.groupHardware(unit.hardware);
      if (!panels.length && !hardware.length) return;
      const title = unit.outside ? unit.name : unit.id + ". " + unit.name + " (" + unit.section + ")";
      sheet.getRange(row, 1, 1, 7).merge().setValue(title).setFontWeight("bold").setFontSize(16).setBackground("#dbeafe").setFontColor("#1e3a5f");
      row += 1;
      if (panels.length) {
        const columns = ["Позиция", "Наименование", "Количество", "Длина", "Ширина", "Толщина"];
        sheet.getRange(row, 1, 1, columns.length).setValues([columns]);
        this.styleHeader(sheet.getRange(row, 1, 1, columns.length));
        row += 1;
        ["ЛДСП", "МДФ", "Прочие"].forEach(type => {
          const items = panels.filter(panel => panel.materialType === type);
          if (!items.length) return;
          sheet.getRange(row, 1, 1, columns.length).merge().setValue(type).setFontWeight("bold").setBackground("#f1f5f9");
          row += 1;
          sheet.getRange(row, 1, items.length, columns.length).setValues(items.map(panel => [panel.position, panel.name, panel.quantity, panel.length, panel.width, panel.thickness]));
          this.styleRows(sheet.getRange(row, 1, items.length, columns.length));
          row += items.length;
        });
      }
      if (hardware.length) {
        if (panels.length) row += 1;
        sheet.getRange(row, 1, 1, 5).merge().setValue("Фурнитура").setFontWeight("bold").setFontSize(13).setBackground("#fbd9d7");
        row += 1;
        const columns = ["Артикул", "Наименование", "Ед. измерения", "Количество", "Примечание"];
        sheet.getRange(row, 1, 1, columns.length).setValues([columns]);
        this.styleHeader(sheet.getRange(row, 1, 1, columns.length)).setBackground("#fbd9d7");
        row += 1;
        sheet.getRange(row, 1, hardware.length, columns.length).setValues(hardware.map(item => [item.article, item.name, item.unit, item.quantity, item.manufacturer || "—"]));
        this.styleRows(sheet.getRange(row, 1, hardware.length, columns.length));
        row += hardware.length;
      }
      row += 2;
    });
    sheet.setFrozenRows(3);
    [100, 380, 110, 100, 120, 90, 90].forEach((width, index) => sheet.setColumnWidth(index + 1, width));
    sheet.autoResizeRows(1, Math.max(1, row - 1));
  }

  static writeDetails(sheet, panels) {
    const columns = ["Позиция", "Наименование", "Тип материала", "Секция", "Сборочная единица", "Количество", "Длина", "Ширина", "Толщина"];
    const rows = this.sortPanels(this.uniqueDetailPanels(panels));
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
    this.styleHeader(sheet.getRange(1, 1, 1, columns.length));
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, columns.length).setValues(rows.map(panel => [panel.position, panel.name, panel.materialType, panel.section, panel.unit, panel.quantity, panel.length, panel.width, panel.thickness]));
      this.styleRows(sheet.getRange(2, 1, rows.length, columns.length));
    }
    sheet.setFrozenRows(1);
    [14, 38, 16, 18, 30, 12, 12, 12, 12].forEach((width, index) => sheet.setColumnWidth(index + 1, width * 7));
    sheet.autoResizeRows(1, Math.max(1, rows.length + 1));
  }

  static groupHardware(hardware) {
    const grouped = new Map();
    hardware.forEach(item => {
      const key = JSON.stringify([item.article, item.name, item.unit, item.manufacturer || ""]);
      if (!grouped.has(key)) grouped.set(key, { article: item.article, name: item.name, unit: item.unit, quantity: 0, manufacturer: item.manufacturer || "" });
      grouped.get(key).quantity += Number(item.quantity) || 0;
    });
    return Array.from(grouped.values()).sort((left, right) => String(left.article).localeCompare(String(right.article), "ru"));
  }

  static writeHardware(sheet, hardware) {
    const columns = ["Артикул", "Наименование", "Ед. измерения", "Количество", "Примечание"];
    const rows = this.groupHardware(hardware);
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
    this.styleHeader(sheet.getRange(1, 1, 1, columns.length));
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, columns.length).setValues(rows.map(item => [item.article, item.name, item.unit, item.quantity, item.manufacturer || "—"]));
      this.styleRows(sheet.getRange(2, 1, rows.length, columns.length));
    } else {
      sheet.getRange(2, 1).setValue("Фурнитура не найдена в XML.");
    }
    sheet.setFrozenRows(1);
    [24, 46, 16, 14, 38].forEach((width, index) => sheet.setColumnWidth(index + 1, width * 7));
    sheet.autoResizeRows(1, Math.max(2, rows.length + 1));
  }

  static styleTitle(range) {
    range.setBackground("#155e75").setFontColor("#ffffff").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("left");
  }

  static styleHeader(range) {
    return range.setBackground("#dbeafe").setFontColor("#0f172a").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  }

  static styleRows(range) {
    range.setVerticalAlignment("top").setWrap(true).setBorder(true, true, true, true, true, true, "#d1d5db", SpreadsheetApp.BorderStyle.SOLID);
  }

  static text(element, name) {
    const child = element && element.getChild(name);
    return child ? String(child.getText() || "").trim() : "";
  }

  static numberOrText(value) {
    const text = String(value || "").trim();
    if (!text) return "—";
    const number = Number(text.replace(",", "."));
    return Number.isFinite(number) ? number : text;
  }
}
