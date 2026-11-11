/**
 * Сервис работы с Google Sheets
 */
class SheetService {

  // Cutover safety: the primary order sheet is a read-only SQL mirror after
  // explicit operator fencing. The authenticated mirror receiver may append
  // a row, but application CRUD must use the PHP/MySQL master.
  static assertProductionSheetWritable() {
    const frozen=PropertiesService.getScriptProperties().getProperty("TIMEWEB_PRODUCTION_GOOGLE_WRITES_DISABLED")==="enabled";
    const mirror=typeof TimewebProductionMirror!=="undefined" && TimewebProductionMirror._writing===true;
    if(frozen&&!mirror)throw new Error("SQL_MASTER|Изменения заказов выполняются только в PHP/MySQL.");
  }

  static getSpreadsheet() {
    if (!this._spreadsheet) {
      this._spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    }
    return this._spreadsheet;
  }

  static normalizeOrderPart(value) {
    return String(value === null || value === undefined ? "" : value)
      .trim()
      .replace(/\s+/g, " ")
      .toLowerCase();
  }

  static buildOrderLookupKey(contract, client, product) {
    return [contract, client, product].map(value => this.normalizeOrderPart(value)).join("|");
  }

  // Переносит площадь покраски из рабочего производственного листа в основную
  // строку изделия. Источник остаётся главным: изменившееся значение будет
  // синхронизировано при следующем обновлении приложения.
  static syncPaintingAreasFromWorkOrders(targetValues) {
    const sourceSheet = this.getSheet(CONFIG.SHEETS.WORK_ORDERS);
    if (!sourceSheet || sourceSheet.getLastRow() < 2) return targetValues || [];

    const targetSheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    if (!targetSheet || targetSheet.getLastRow() < 2) return targetValues || [];

    const sourceValues = sourceSheet.getDataRange().getValues();
    const sourceHeaders = sourceValues[0].map(value => String(value || "").trim());
    const sourceColumns = {
      contract: sourceHeaders.indexOf("Договор"),
      client: sourceHeaders.indexOf("Клиент"),
      product: sourceHeaders.indexOf("Изделие"),
      area: sourceHeaders.indexOf("Покраска МДФ, м²")
    };
    if (Object.values(sourceColumns).some(index => index < 0)) return targetValues || [];

    const areasByOrder = {};
    sourceValues.slice(1).forEach(row => {
      const key = this.buildOrderLookupKey(
        row[sourceColumns.contract],
        row[sourceColumns.client],
        row[sourceColumns.product]
      );
      const rawArea = row[sourceColumns.area];
      const area = typeof rawArea === "number"
        ? rawArea
        : Number(String(rawArea || "").replace(/\s/g, "").replace(",", "."));
      if (key !== "||" && Number.isFinite(area) && area >= 0) areasByOrder[key] = area;
    });

    targetValues = targetValues || targetSheet.getDataRange().getValues();
    const targetHeaders = targetValues[0].map(value => String(value || "").trim());
    const targetColumns = {
      contract: targetHeaders.indexOf("Договор"),
      client: targetHeaders.indexOf("Клиент"),
      product: targetHeaders.indexOf("Изделие"),
      stage: targetHeaders.indexOf("Этап"),
      area: targetHeaders.indexOf("Площадь покраски, м2")
    };
    if (Object.values(targetColumns).some(index => index < 0)) return targetValues;

    const productionStages = new Set([
      "Запуск", "Распил", "Распил ЛДСП", "Распил МДФ", "Кромка", "ЧПУ",
      "Присадка", "Фрезеровка", "Сборка", "В покраску", "Подготовка",
      "Грунт 1 слой", "Шлифовка 1 слой", "Грунт 2 слой",
      "Шлифовка 2 слой", "Покраска тыльной стороны", "Обклейка",
      "Покраска лицевой стороны", "Упаковка", "Упаковка Корпуса", "Упаковка Фасада",
      "СГП", "СГП Корпуса", "СГП Фасада", "Монтаж"
    ]);
    const areaValues = targetValues.slice(1).map(row => [row[targetColumns.area]]);
    let changed = 0;

    targetValues.slice(1).forEach((row, index) => {
      if (!productionStages.has(String(row[targetColumns.stage] || "").trim())) return;
      const key = this.buildOrderLookupKey(
        row[targetColumns.contract],
        row[targetColumns.client],
        row[targetColumns.product]
      );
      if (!Object.prototype.hasOwnProperty.call(areasByOrder, key)) return;
      const nextArea = areasByOrder[key];
      const currentArea = Number(row[targetColumns.area]);
      if (Number.isFinite(currentArea) && currentArea === nextArea) return;
      areaValues[index][0] = nextArea;
      row[targetColumns.area] = nextArea;
      changed += 1;
    });

    if (changed) {
      targetSheet.getRange(2, targetColumns.area + 1, areaValues.length, 1).setValues(areaValues);
    }
    return targetValues;
  }

  static ensureOrderColumns(fieldNames) {
    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const columns = this.getColumns(sheet);
    const missing = [...new Set(fieldNames || [])]
      .filter(fieldName => fieldName && !columns[fieldName]);
    if (missing.length) {
      const startColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, startColumn, 1, missing.length).setValues([missing]);
      missing.forEach((fieldName, index) => {
        columns[fieldName] = startColumn + index;
      });
    }
    return { sheet, columns };
  }

  // Поля длительности должны быть числами. Новые колонки иногда наследуют
  // формат даты от соседнего «Дедлайна», из-за чего 0,5 отображается как дата.
  static ensureOrderDurationNumberFormat(fieldNames) {
    const result = this.ensureOrderColumns(fieldNames);
    const rowCount = Math.max(result.sheet.getMaxRows() - 1, 1);
    [...new Set(fieldNames || [])].forEach(fieldName => {
      const column = result.columns[fieldName];
      if (column) result.sheet.getRange(2, column, rowCount, 1).setNumberFormat("0.##");
    });
    return result.sheet;
  }

  /**
   * «Дедлайн» определяется в карточке изделия: либо сроком изготовления от
   * даты согласования проформы, либо конкретной датой. Ранее все строки
   * использовали жёсткие 50 дней; при первой миграции сохраняем это значение.
   */
  static ensureOrderDeadlineColumn() {
    const result = this.ensureOrderColumns([
      "Дедлайн", "Способ определения дедлайна", "Срок изготовления, дней", "Дата дедлайна вручную"
    ]);
    const sheet = result.sheet;
    const columns = this.getColumns(sheet);
    const deadlineColumn = columns["Дедлайн"];
    const approvalColumn = columns["Проформа Согласована"];
    const modeColumn = columns["Способ определения дедлайна"];
    const daysColumn = columns["Срок изготовления, дней"];
    const manualDateColumn = columns["Дата дедлайна вручную"];
    if (!deadlineColumn || !approvalColumn || !modeColumn || !daysColumn || !manualDateColumn) return sheet;

    const lastRow = sheet.getLastRow();
    const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
    const keyColumns = [columns["Договор"], columns["Клиент"], columns["Изделие"]];
    let lastOrderOffset = 0;
    data.forEach((row, index) => {
      if (keyColumns.some(column => String(row[column - 1] || "").trim())) lastOrderOffset = index + 1;
    });
    const rowCount = Math.max(lastOrderOffset, 1);

    // Миграция уже созданных изделий: только пустые новые поля получают
    // прежнее поведение «50 календарных дней от согласования проформы».
    if (lastOrderOffset) {
      const modes = sheet.getRange(2, modeColumn, lastOrderOffset, 1).getValues();
      const days = sheet.getRange(2, daysColumn, lastOrderOffset, 1).getValues();
      let modeChanged = false;
      let daysChanged = false;
      data.slice(0, lastOrderOffset).forEach((row, index) => {
        const isOrder = keyColumns.some(column => String(row[column - 1] || "").trim());
        if (!isOrder) return;
        if (!String(modes[index][0] || "").trim()) {
          modes[index][0] = "От согласования проформы";
          modeChanged = true;
        }
        if (days[index][0] === "" || days[index][0] === null) {
          days[index][0] = 50;
          daysChanged = true;
        }
      });
      if (modeChanged) sheet.getRange(2, modeColumn, lastOrderOffset, 1).setValues(modes);
      if (daysChanged) sheet.getRange(2, daysColumn, lastOrderOffset, 1).setValues(days);
    }

    const modeOffset = modeColumn - deadlineColumn;
    const approvalOffset = approvalColumn - deadlineColumn;
    const daysOffset = daysColumn - deadlineColumn;
    const manualDateOffset = manualDateColumn - deadlineColumn;
    const formula = `=IF(RC[${modeOffset}]="Конкретная дата";IF(OR(RC[${manualDateOffset}]="";RC[${manualDateOffset}]=0);"—";RC[${manualDateOffset}]);IF(OR(RC[${approvalOffset}]="";RC[${approvalOffset}]=0);"—";RC[${approvalOffset}]+IF(OR(RC[${daysOffset}]="";RC[${daysOffset}]=0);50;RC[${daysOffset}])))`;
    const range = sheet.getRange(2, deadlineColumn, rowCount, 1);
    const formulas = range.getFormulas();
    const usesLegacyFormula = formulas.some(row => row[0] && row[0].indexOf("Конкретная дата") === -1);
    if (usesLegacyFormula || !formulas.some(row => row[0])) range.setFormulaR1C1(formula);
    range.setNumberFormat("dd.MM.yyyy");
    sheet.getRange(2, manualDateColumn, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("dd.MM.yyyy");
    sheet.getRange(2, daysColumn, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("0");
    const modeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["От согласования проформы", "Конкретная дата"], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, modeColumn, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(modeRule);
    return sheet;
  }

  static setOrderDeadlineFormula(sheet, rowNumber) {
    const columns = this.getColumns(sheet);
    const deadlineColumn = columns["Дедлайн"];
    const approvalColumn = columns["Проформа Согласована"];
    const modeColumn = columns["Способ определения дедлайна"];
    const daysColumn = columns["Срок изготовления, дней"];
    const manualDateColumn = columns["Дата дедлайна вручную"];
    if (!deadlineColumn || !approvalColumn || !modeColumn || !daysColumn || !manualDateColumn || rowNumber < 2) return;
    const modeOffset = modeColumn - deadlineColumn;
    const approvalOffset = approvalColumn - deadlineColumn;
    const daysOffset = daysColumn - deadlineColumn;
    const manualDateOffset = manualDateColumn - deadlineColumn;
    sheet.getRange(rowNumber, deadlineColumn)
      .setFormulaR1C1(`=IF(RC[${modeOffset}]="Конкретная дата";IF(OR(RC[${manualDateOffset}]="";RC[${manualDateOffset}]=0);"—";RC[${manualDateOffset}]);IF(OR(RC[${approvalOffset}]="";RC[${approvalOffset}]=0);"—";RC[${approvalOffset}]+IF(OR(RC[${daysOffset}]="";RC[${daysOffset}]=0);50;RC[${daysOffset}])))`)
      .setNumberFormat("dd.MM.yyyy");
  }

  static migrateLegacyOrderValues(rows) {
    if (!rows || rows.length < 2) return rows || [];
    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const headers = rows[0].map(value => String(value || "").trim());
    const stageIndex = headers.indexOf("Этап");
    const branchStageIndexes = [headers.indexOf("Этап ЛДСП"), headers.indexOf("Этап МДФ"), headers.indexOf("Этап Шпон")]
      .filter(index => index >= 0);
    const statusIndex = headers.indexOf("Статус");
    const replacements = {
      "Обработка": "В покраску",
      "Грунт": "Грунт 1 слой",
      "Белая шлифовка": "Шлифовка 1 слой",
      "Шлифовка": "Шлифовка 1 слой",
      "Покраска": "Покраска лицевой стороны"
    };
    let stageChanged = false;
    const changedBranchStageIndexes = new Set();
    let statusChanged = false;
    rows.slice(1).forEach(row => {
      if (stageIndex >= 0) {
        const stage = String(row[stageIndex] || "").trim();
        if (replacements[stage]) {
          row[stageIndex] = replacements[stage];
          stageChanged = true;
        }
      }
      branchStageIndexes.forEach(index => {
        const stage = String(row[index] || "").trim();
        const replacement = stage === "ЧПУ" ? "Фрезеровка" : replacements[stage];
        if (!replacement) return;
        row[index] = replacement;
        changedBranchStageIndexes.add(index);
      });
      if (statusIndex >= 0 && String(row[statusIndex] || "").trim().toLowerCase() === "старт") {
        row[statusIndex] = "В работе";
        statusChanged = true;
      }
    });
    if (stageChanged) {
      sheet.getRange(2, stageIndex + 1, rows.length - 1, 1)
        .setValues(rows.slice(1).map(row => [row[stageIndex]]));
    }
    changedBranchStageIndexes.forEach(index => {
      sheet.getRange(2, index + 1, rows.length - 1, 1)
        .setValues(rows.slice(1).map(row => [row[index]]));
    });
    if (statusChanged) {
      sheet.getRange(2, statusIndex + 1, rows.length - 1, 1)
        .setValues(rows.slice(1).map(row => [row[statusIndex]]));
    }
    return rows;
  }

  // Продвигает операции, которые были завершены до появления явных переходов
  // «Присадка → Сборка», «Фрезеровка → В покраску» и упаковочных операций.
  static migrateCompletedManufacturingStages(rows) {
    if (!rows || rows.length < 2) return rows || [];
    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const headers = rows[0].map(value => String(value || "").trim());
    const stageIndex = headers.indexOf("Этап");
    const statusIndex = headers.indexOf("Статус");
    const facadePackagingStageIndex = headers.indexOf("Этап упаковки фасада");
    const facadePackagingStatusIndex = headers.indexOf("Статус упаковки фасада");
    const branches = [
      { kind: "ldsp", stage: headers.indexOf("Этап ЛДСП"), status: headers.indexOf("Статус ЛДСП") },
      { kind: "mdf", stage: headers.indexOf("Этап МДФ"), status: headers.indexOf("Статус МДФ") },
      { kind: "veneer", stage: headers.indexOf("Этап Шпон"), status: headers.indexOf("Статус Шпон") },
      { kind: "facade-packaging", stage: facadePackagingStageIndex, status: facadePackagingStatusIndex }
    ].filter(branch => branch.stage >= 0 && branch.status >= 0);
    const changedColumns = new Set();

    rows.slice(1).forEach(row => {
      branches.forEach(branch => {
        let currentStage = String(row[branch.stage] || "").trim();
        if (branch.kind === "veneer" && ["Упаковка Корпуса", "СГП Корпуса"].includes(currentStage)) {
          row[branch.stage] = currentStage === "Упаковка Корпуса" ? "Упаковка Фасада" : "СГП Фасада";
          currentStage = String(row[branch.stage]);
          changedColumns.add(branch.stage);
        }
        if (String(row[branch.status] || "").trim().toLowerCase() !== "готово") return;
        if (currentStage === "Присадка" || currentStage === "Присадка Шпон") {
          row[branch.stage] = "Сборка";
          row[branch.status] = "";
          changedColumns.add(branch.stage);
          changedColumns.add(branch.status);
        } else if (currentStage === "Фрезеровка") {
          row[branch.stage] = "В покраску";
          row[branch.status] = "";
          changedColumns.add(branch.stage);
          changedColumns.add(branch.status);
          if (stageIndex >= 0) {
            row[stageIndex] = "В покраску";
            changedColumns.add(stageIndex);
          }
          if (statusIndex >= 0) {
            row[statusIndex] = "";
            changedColumns.add(statusIndex);
          }
        } else if (currentStage === "Сборка") {
          row[branch.stage] = branch.kind === "veneer" ? "Упаковка Фасада" : "Упаковка Корпуса";
          row[branch.status] = "";
          changedColumns.add(branch.stage);
          changedColumns.add(branch.status);
        } else if (currentStage === "Упаковка Корпуса" || currentStage === "Упаковка Фасада") {
          row[branch.stage] = currentStage === "Упаковка Корпуса" ? "СГП Корпуса" : "СГП Фасада";
          row[branch.status] = "";
          changedColumns.add(branch.stage);
          changedColumns.add(branch.status);
          if (stageIndex >= 0) {
            row[stageIndex] = "СГП";
            changedColumns.add(stageIndex);
          }
          if (statusIndex >= 0) {
            row[statusIndex] = "";
            changedColumns.add(statusIndex);
          }
        } else if (currentStage === "СГП Корпуса" || currentStage === "СГП Фасада") {
          row[branch.stage] = "Монтаж";
          row[branch.status] = "";
          changedColumns.add(branch.stage);
          changedColumns.add(branch.status);
          if (stageIndex >= 0) {
            row[stageIndex] = "Монтаж";
            changedColumns.add(stageIndex);
          }
          if (statusIndex >= 0) {
            row[statusIndex] = "";
            changedColumns.add(statusIndex);
          }
        }
      });

      if (stageIndex >= 0 && String(row[stageIndex] || "").trim() === "Упаковка" &&
          facadePackagingStageIndex >= 0 && !String(row[facadePackagingStageIndex] || "").trim()) {
        row[facadePackagingStageIndex] = "Упаковка Фасада";
        changedColumns.add(facadePackagingStageIndex);
        if (facadePackagingStatusIndex >= 0) {
          row[facadePackagingStatusIndex] = "";
          changedColumns.add(facadePackagingStatusIndex);
        }
      }
    });

    changedColumns.forEach(index => {
      sheet.getRange(2, index + 1, rows.length - 1, 1)
        .setValues(rows.slice(1).map(row => [row[index]]));
    });
    return rows;
  }

  static migrateDictionaryWorkStatusName() {
    const sheet = this.getSheet(CONFIG.SHEETS.DICTIONARIES);
    const columns = this.getColumns(sheet);
    const statusColumn = columns["Статус"];
    if (!statusColumn || sheet.getLastRow() < 2) return;
    const range = sheet.getRange(2, statusColumn, sheet.getLastRow() - 1, 1);
    const values = range.getValues();
    let changed = false;
    values.forEach(row => {
      if (String(row[0] || "").trim().toLowerCase() === "старт") {
        row[0] = "В работе";
        changed = true;
      }
    });
    if (changed) range.setValues(values);
  }

  static ensureOrderColumn(fieldName) {
    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const columns = this.getColumns(sheet);

    if (!columns[fieldName]) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(fieldName);
    }

    return sheet;
  }

  // Однократная миграция названий старого производственного маршрута.
  // После замены повторные вызовы ничего не записывают.
  static migrateLegacyProductionStages() {
    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const columns = this.getColumns(sheet);
    const stageColumn = columns["Этап"];
    if (!stageColumn || sheet.getLastRow() < 2) return;

    const range = sheet.getRange(2, stageColumn, sheet.getLastRow() - 1, 1);
    const values = range.getValues();
    const replacements = {
      "Обработка": "В покраску",
      "Грунт": "Грунт 1 слой",
      "Белая шлифовка": "Шлифовка 1 слой",
      "Шлифовка": "Шлифовка 1 слой",
      "Покраска": "Покраска лицевой стороны"
    };
    let changed = false;
    values.forEach(row => {
      const current = String(row[0] || "").trim();
      if (!replacements[current]) return;
      row[0] = replacements[current];
      changed = true;
    });
    if (changed) range.setValues(values);
  }

  // Однократное переименование производственного участка. Сканируются этапные
  // колонки листа «Изделия», справочник и отдельный лист «Этапы»; статусы и очереди
  // намеренно не затрагиваются. Повторный запуск не выполняет записей.
  static migrateDeliveryStagesToSgp() {
    const replacements = {
      "Доставка": "СГП",
      "Доставка Корпуса": "СГП Корпуса",
      "Доставка Фасада": "СГП Фасада"
    };
    const migrateColumns = (sheet, columns) => {
      if (sheet.getLastRow() < 2) return 0;
      let changed = 0;
      columns.forEach(column => {
        if (!column) return;
        const range = sheet.getRange(2, column, sheet.getLastRow() - 1, 1);
        const values = range.getValues();
        let columnChanged = false;
        values.forEach(row => {
          const replacement = replacements[String(row[0] || "").trim()];
          if (!replacement) return;
          row[0] = replacement;
          changed += 1;
          columnChanged = true;
        });
        if (columnChanged) range.setValues(values);
      });
      return changed;
    };

    const orderSheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const orderColumns = this.getColumns(orderSheet);
    const orderStageColumns = Object.entries(orderColumns)
      .filter(([name]) => name === "Этап" || name.startsWith("Этап "))
      .map(([, column]) => column);
    const dictionarySheet = this.getSheet(CONFIG.SHEETS.DICTIONARIES);
    const dictionaryColumns = this.getColumns(dictionarySheet);
    const stageCatalogSheet = this.getSheet(CONFIG.SHEETS.STAGES);
    const ordersChanged = migrateColumns(orderSheet, orderStageColumns);
    const dictionaryChanged = migrateColumns(dictionarySheet, [dictionaryColumns["Этапы"]]);
    const stageCatalogChanged = migrateColumns(stageCatalogSheet, [1]);
    this.ensureOrderStageValidation();
    return { success: true, ordersChanged, dictionaryChanged, stageCatalogChanged };
  }

  // Однократно переименовывает рабочий статус во всех местах, где он хранится.
  // Старое значение оставлено только как условие миграции, чтобы уже созданные
  // заказы и список значений в справочнике стали единообразными.
  static migrateWorkStatusName() {
    [CONFIG.SHEETS.ORDERS, CONFIG.SHEETS.DICTIONARIES].forEach(sheetName => {
      const sheet = this.getSheet(sheetName);
      const columns = this.getColumns(sheet);
      const statusColumn = columns["Статус"];
      if (!statusColumn || sheet.getLastRow() < 2) return;

      const range = sheet.getRange(2, statusColumn, sheet.getLastRow() - 1, 1);
      const values = range.getValues();
      let changed = false;
      values.forEach(row => {
        if (String(row[0] || "").trim().toLowerCase() === "старт") {
          row[0] = "В работе";
          changed = true;
        }
      });
      if (changed) range.setValues(values);
    });
  }

  // Добавляет справочник роли только один раз. После создания списка его можно
  // свободно редактировать в Google Таблице — приложение значения не перезаписывает.
  static ensureDictionaryColumn(fieldName, defaultValues) {
    const sheet = this.getSheet(CONFIG.SHEETS.DICTIONARIES);
    const columns = this.getColumns(sheet);
    let column = columns[fieldName];

    if (!column) {
      column = sheet.getLastColumn() + 1;
      sheet.getRange(1, column).setValue(fieldName);
      if (defaultValues && defaultValues.length) {
        sheet.getRange(2, column, defaultValues.length, 1).setValues(defaultValues.map(value => [value]));
      }
      return sheet;
    }

    const existing = sheet.getRange(2, column, Math.max(sheet.getLastRow() - 1, 1), 1)
      .getValues().flat().filter(Boolean);
    if (!existing.length && defaultValues && defaultValues.length) {
      sheet.getRange(2, column, defaultValues.length, 1).setValues(defaultValues.map(value => [value]));
    }
    return sheet;
  }

  // Добавляет недостающие системные значения, не удаляя и не переставляя этапы,
  // которые пользователь вручную внёс в справочник.
  static syncDictionaryColumn(fieldName, values) {
    const sheet = this.getSheet(CONFIG.SHEETS.DICTIONARIES);
    const columns = this.getColumns(sheet);
    let column = columns[fieldName];
    if (!column) {
      column = sheet.getLastColumn() + 1;
      sheet.getRange(1, column).setValue(fieldName);
    }

    const rowCount = Math.max(sheet.getLastRow() - 1, 1);
    const rawValues = sheet.getRange(2, column, rowCount, 1).getValues().flat();
    const existing = rawValues.map(value => String(value || "").trim()).filter(Boolean);
    const expected = (values || []).map(value => String(value || "").trim()).filter(Boolean);
    const existingSet = new Set(existing);
    const missing = expected.filter(value => !existingSet.has(value));
    if (!missing.length) return sheet;

    let lastFilledOffset = -1;
    rawValues.forEach((value, index) => {
      if (String(value || "").trim()) lastFilledOffset = index;
    });
    const startRow = 2 + lastFilledOffset + 1;
    sheet.getRange(startRow, column, missing.length, 1).setValues(missing.map(value => [value]));
    return sheet;
  }

  static ensureOrderStageValidation() {
    const dictionarySheet = this.getSheet(CONFIG.SHEETS.DICTIONARIES);
    const dictionaryColumns = this.getColumns(dictionarySheet);
    const stageDictionaryColumn = dictionaryColumns["Этапы"];
    if (!stageDictionaryColumn) return;

    const stageValues = dictionarySheet
      .getRange(2, stageDictionaryColumn, Math.max(dictionarySheet.getLastRow() - 1, 1), 1)
      .getValues()
      .flat()
      .map(value => String(value || "").trim())
      .filter(Boolean);
    if (!stageValues.length) return;

    const sourceRange = dictionarySheet.getRange(2, stageDictionaryColumn, stageValues.length, 1);
    const orderSheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const orderColumns = this.getColumns(orderSheet);
    const stageColumn = orderColumns["Этап"];
    if (!stageColumn) return;

    const currentRule = orderSheet.getRange(2, stageColumn).getDataValidation();
    let isCurrent = false;
    if (currentRule && currentRule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const criteriaValues = currentRule.getCriteriaValues();
      const currentRange = criteriaValues && criteriaValues[0];
      isCurrent = currentRange &&
        currentRange.getSheet().getName() === sourceRange.getSheet().getName() &&
        currentRange.getA1Notation() === sourceRange.getA1Notation();
    }
    if (isCurrent) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    orderSheet
      .getRange(2, stageColumn, Math.max(orderSheet.getMaxRows() - 1, 1), 1)
      .setDataValidation(rule);
  }

  /**
   * Возвращает лист
   */
  static getSheet(sheetName) {
    if (!this._sheets) this._sheets = {};
    if (!this._sheets[sheetName]) {
      this._sheets[sheetName] = this.getSpreadsheet().getSheetByName(sheetName);
    }
    return this._sheets[sheetName];

  }

  /**
   * Возвращает все данные листа
   */
  static getData(sheetName) {

    return PerformanceService.measureRead(() => this
      .getSheet(sheetName)
      .getDataRange()
      .getValues());

  }

  /**
   * Возвращает массив объектов
   */
  static getObjects(sheetName) {

    const rows = this.getData(sheetName);

    if (rows.length < 2) {
      return [];
    }

    const headers = rows[0];

    return rows
      .slice(1)
      .map(row => {

        const obj = {};

        headers.forEach((header, index) => {
          obj[header] = row[index];
        });

        return obj;

      });

  }

  static objectsFromRows(rows) {
    if (!rows || rows.length < 2) return [];
    const headers = rows[0];
    return rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => { obj[header] = row[index]; });
      return obj;
    });
  }

  /**
   * Словарь колонок
   */
  static getColumns(sheet) {

    const headers = PerformanceService.measureRead(() => sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]);

    const columns = {};

    headers.forEach((name, index) => {

      columns[name] = index + 1;

    });

    return columns;

  }

  /**
   * Ключ бизнес-заказа. Нормализация защищает от дублей вида «35RM» / «35rm»
   * и от случайных неразрывных либо двойных пробелов в названии изделия.
   */
  static getNormalizedOrderKey(order) {
    const read = (field, property) => order &&
      (Object.prototype.hasOwnProperty.call(order, field) ? order[field] : order[property]);
    const normalize = value => String(value == null ? "" : value)
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLocaleLowerCase("ru-RU");

    return [
      normalize(read("Договор", "contract")),
      normalize(read("Клиент", "client")),
      normalize(read("Изделие", "product"))
    ].join("|");
  }

  /**
   * Проверяет уникальность договора, клиента и изделия на листе «Изделия».
   * Возвращает номер уже существующей строки либо null.
   */
  static findDuplicateOrderRow(order, sheet) {
    const targetSheet = sheet || this.getSheet(CONFIG.SHEETS.ORDERS);
    const rows = targetSheet.getDataRange().getValues();
    if (rows.length < 2) return null;

    const headers = rows[0].map(header => String(header || "").trim());
    const contractIndex = headers.indexOf("Договор");
    const clientIndex = headers.indexOf("Клиент");
    const productIndex = headers.indexOf("Изделие");
    if ([contractIndex, clientIndex, productIndex].some(index => index === -1)) {
      throw new Error("В листе «Изделия» нет колонок для проверки уникальности заказа.");
    }

    const newKey = this.getNormalizedOrderKey(order);
    for (let index = 1; index < rows.length; index++) {
      const rowKey = this.getNormalizedOrderKey({
        contract: rows[index][contractIndex],
        client: rows[index][clientIndex],
        product: rows[index][productIndex]
      });
      if (rowKey === newKey) return index + 1;
    }
    return null;
  }

  static assertOrderKeyUnique(order, sheet) {
    if (this.findDuplicateOrderRow(order, sheet)) {
      throw new Error(
        "Заказ с таким договором, клиентом и изделием уже существует. Откройте существующую строку вместо повторного добавления."
      );
    }
  }

  /**
   * Поиск строки по ключу
   */
  static findRowByKey(sheet, key) {

    const values = PerformanceService.measureRead(() => sheet.getDataRange().getValues());

    if (values.length < 2) {
      return null;
    }

    const columns = this.getColumns(sheet);

    const contractCol = columns["Договор"];
    const clientCol = columns["Клиент"];
    const productCol = columns["Изделие"];

    for (let row = 2; row <= values.length; row++) {

      const currentKey = [

        values[row - 1][contractCol - 1],
        values[row - 1][clientCol - 1],
        values[row - 1][productCol - 1]

      ].join("|");

      if (currentKey === key) {
        return row;
      }

    }

    return null;

  }
  // Контекст одной записи переиспользуется в рамках серверного действия:
  // поиск ключа и заголовков не должен повторно читать весь лист перед записью.
  static getOrderContext(key) {
    if (!this._orderContexts) this._orderContexts = {};
    if (this._orderContexts[key]) return this._orderContexts[key];
    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const rows = this.getData(CONFIG.SHEETS.ORDERS);
    const headers = rows[0] || [];
    const columns = {};
    headers.forEach((header, index) => { columns[String(header || "").trim()] = index + 1; });
    const lookupStartedAt = Date.now();
    const rowNumber = rows.findIndex((row, index) => index > 0 && [
      row[(columns["Договор"] || 1) - 1], row[(columns["Клиент"] || 1) - 1], row[(columns["Изделие"] || 1) - 1]
    ].join("|") === key) + 1;
    if (PerformanceService.current) {
      PerformanceService.current.lookupMs += Date.now() - lookupStartedAt;
    }
    if (rowNumber < 2) throw new Error("Заказ не найден: " + key);
    const context = { sheet, rows, columns, rowNumber };
    this._orderContexts[key] = context;
    return context;
  }

  static resetOrderContexts() {
    this._orderContexts = {};
  }

  static getActionOrderRows() {
    const contexts = Object.values(this._orderContexts || {});
    return contexts.length ? contexts[0].rows : null;
  }

  /**
   * Обновляет одно поле
   */
  static updateField(key, fieldName, value) {
    return this.updateFields(key, { [fieldName]: value });

  }


  /**
   * Обновляет несколько полей одновременно
   */
  static updateFields(key, fields) {
    this.assertProductionSheetWritable();
    const context = this.getOrderContext(key);
    const missing = Object.keys(fields || {}).filter(fieldName => !context.columns[fieldName]);
    if (missing.length) {
      this.ensureOrderColumns(missing);
      this.resetOrderContexts();
      return this.updateFields(key, fields);
    }
    const changes = Object.entries(fields || {})
      .map(([fieldName, value]) => ({ column: context.columns[fieldName], value }))
      .sort((left, right) => left.column - right.column);
    for (let index = 0; index < changes.length;) {
      const start = index;
      while (index + 1 < changes.length && changes[index + 1].column === changes[index].column + 1) index += 1;
      const group = changes.slice(start, index + 1);
      PerformanceService.measureWrite(() => context.sheet
        .getRange(context.rowNumber, group[0].column, 1, group.length)
        .setValues([group.map(change => change.value)]));
      group.forEach(change => { context.rows[context.rowNumber - 1][change.column - 1] = change.value; });
      index += 1;
    }
    return true;

  }

  /**
   * Обновляет несколько полей в уже найденной строке.
   *
   * Составной ключ заказа не содержит этап. Поэтому один договор может иметь
   * несколько строк с одинаковыми договором, клиентом и изделием (например,
   * «Запуск» и «Технолог»). В таких случаях обновление по одному ключу
   * небезопасно — вызывающий код должен передать точный номер строки.
   */
  static updateFieldsAtRow(rowNumber, fields) {
    this.assertProductionSheetWritable();

    const sheet = Object.prototype.hasOwnProperty.call(fields, "Статус")
      ? this.ensureOrderColumn("Статус")
      : this.getSheet(CONFIG.SHEETS.ORDERS);

    if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      throw new Error("Не найдена строка заказа для сохранения.");
    }

    const columns = this.getColumns(sheet);
    Object.entries(fields).forEach(([fieldName, value]) => {
      const column = columns[fieldName];
      if (column) sheet.getRange(rowNumber, column).setValue(value);
    });

    return true;
  }

  static appendOrder(order) {
    this.assertProductionSheetWritable();
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      if (order.manualClaim) {
        this.ensureOrderColumn("Ручная рекламация");
        this.ensureOrderColumn("Комментарий рекламации");
      }
      this.ensureOrderColumn("Статус");
      this.ensureOrderColumn("Дни ТЗ");
      this.ensureOrderColumn("Дни Проформы");
      this.ensureOrderColumn("Дни КД");
      this.ensureOrderColumn("Резерв КД — очередь");
      this.ensureOrderColumn("Проформа Согласована");
      this.ensureOrderColumn("Способ определения дедлайна");
      this.ensureOrderColumn("Срок изготовления, дней");
      this.ensureOrderColumn("Дата дедлайна вручную");
      this.ensureOrderColumn("Дедлайн");
      this.ensureOrderColumn("Папка Google Drive");
      this.ensureOrderColumn("Дни Шлифовка");
      this.ensureOrderColumn("Дни Грунт");
      this.ensureOrderColumn("Дни Белая шлифовка");
      this.ensureOrderColumn("Дни Покраска");
      this.ensureOrderColumn("Дни Подготовка");
      this.ensureOrderColumn("Дни Грунт 1 слой");
      this.ensureOrderColumn("Дни Грунт 2 слой");
      this.ensureOrderColumn("Дни Шлифовка 2 слой");
      this.ensureOrderColumn("Дни Покраска тыльной стороны");
      this.ensureOrderColumn("Дни Обклейка");
      this.ensureOrderColumn("Дни Покраска лицевой стороны");
      this.ensureOrderColumn("Дни Шлифовка перед лаком");
      this.ensureOrderColumn("Дни Лак");
      this.ensureOrderColumn("Дни Полировка");
      this.ensureOrderColumn("Лаковые операции");
      this.ensureOrderColumn("Исключить Грунт 2 слой");
      this.ensureOrderColumn("Исключить Шлифовка 2 слой");
      this.ensureOrderColumn("Исключить Покраска тыльной стороны");
      this.ensureOrderColumn("Исключить Обклейка");
      this.ensureOrderColumn("Очередь производства");
      this.ensureOrderColumn("Шлифовщик");
      this.ensureOrderColumn("Очередь шлифовки");
      this.ensureOrderColumn("Резерв Белая шлифовка — очередь");
      this.ensureOrderColumn("Плановая дата шлифовки");
      this.ensureOrderColumn("Маляр");
      this.ensureOrderColumn("Очередь грунт");
      this.ensureOrderColumn("Резерв Покраска — очередь");
      this.ensureOrderColumn("Плановая дата покраски");
      this.ensureOrderColumn("Ручной приоритет даты");
      this.ensureOrderColumn("Факт. начало работы");
      this.ensureOrderColumn("Факт. половина начала");
      this.ensureOrderColumn("Остаток длительности");
      this.ensureOrderColumn("Площадь покраски, м2");
      const sheet = this.ensureOrderColumn("Плановая дата");
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const normalizedHeaders = headers.map(header => String(header || "").trim());
      const requiredHeaders = ["Договор", "Клиент", "Изделие", "Менеджер", "Этап", "Ответственный", "Конструктор", "Длительность"];

      requiredHeaders.forEach(header => {
        if (!normalizedHeaders.includes(header)) {
          throw new Error("В листе «Изделия» нет колонки: " + header);
        }
      });

      const queueHeader = normalizedHeaders.includes("Очередность")
        ? "Очередность"
        : normalizedHeaders.includes("Очередь")
          ? "Очередь"
          : null;

      if (!queueHeader) {
        throw new Error("В листе «Изделия» нет колонки «Очередность».");
      }

      const rows = sheet.getDataRange().getValues();
      const queueIndex = normalizedHeaders.indexOf(queueHeader);
      // Повторная проверка выполняется уже под блокировкой: два пользователя,
      // нажавшие «Добавить заказ» одновременно, не смогут создать дубль.
      this.assertOrderKeyUnique(order, sheet);
      let maxQueue = 0;

      for (let index = 1; index < rows.length; index++) {
        const row = rows[index];
        maxQueue = Math.max(maxQueue, Number(row[queueIndex]) || 0);
      }

      const values = {
        "Договор": order.contract,
        "Клиент": order.client,
        "Изделие": order.product,
        "Менеджер": order.manager,
        "Этап": order.manualClaim ? "Рекламация" : "ТЗ",
        "Ответственный": order.manualClaim ? "Производство" : "Менеджер",
        "Ручная рекламация": order.manualClaim ? true : "",
        "Комментарий рекламации": order.manualClaim ? order.claimComment : "",
        "Конструктор": order.constructor,
        [queueHeader]: maxQueue + 1,
        "Длительность": order.duration,
        "Дни ТЗ": "",
        "Дни Проформы": "",
        "Дни КД": "",
        "Резерв КД — очередь": "",
        "Проформа Согласована": "",
        "Способ определения дедлайна": "От согласования проформы",
        "Срок изготовления, дней": 50,
        "Дата дедлайна вручную": "",
        "Папка Google Drive": order.cloudFolderUrl || "",
        "Дни Шлифовка": "",
        "Дни Грунт": "",
        "Дни Белая шлифовка": "",
        "Дни Покраска": "",
        "Дни Подготовка": "",
        "Дни Грунт 1 слой": "",
        "Дни Грунт 2 слой": "",
        "Дни Шлифовка 2 слой": "",
        "Дни Покраска тыльной стороны": "",
        "Дни Обклейка": "",
        "Дни Покраска лицевой стороны": "",
        "Исключить Грунт 2 слой": false,
        "Исключить Шлифовка 2 слой": false,
        "Исключить Покраска тыльной стороны": false,
        "Исключить Обклейка": false,
        "Очередь производства": "",
        "Плановая дата": "",
        "Ручной приоритет даты": false,
        "Факт. начало работы": "",
        "Факт. половина начала": "",
        "Остаток длительности": "",
        "Статус": "",
        "Площадь покраски, м2": ""
      };
      // Внизу листа заранее протянуты формулы, поэтому getLastRow() указывает
      // на технический хвост, а не на следующую строку заказа. Ищем первую
      // свободную строку именно по бизнес-ключу заказа и сохраняем формулы,
      // проверки данных и форматирование этой строки.
      const keyIndexes = ["Договор", "Клиент", "Изделие"]
        .map(header => normalizedHeaders.indexOf(header));
      let targetRow = rows.findIndex((row, index) =>
        index > 0 && keyIndexes.every(columnIndex =>
          String(row[columnIndex] == null ? "" : row[columnIndex]).trim() === ""
        )
      );

      if (targetRow === -1) {
        const lastGridRow = sheet.getMaxRows();
        sheet.insertRowAfter(lastGridRow);
        targetRow = lastGridRow;
      }

      // findIndex использует нулевой индекс массива, номер строки Sheets — с 1.
      targetRow += 1;
      Object.entries(values).forEach(([header, value]) => {
        const columnIndex = normalizedHeaders.indexOf(header);
        if (columnIndex >= 0) sheet.getRange(targetRow, columnIndex + 1).setValue(value);
      });
      this.setOrderDeadlineFormula(sheet, targetRow);

      return { success: true, queue: values[queueHeader] };
    } finally {
      lock.releaseLock();
    }
  }

  static saveQueueAndCalendar(orders) {
    this.assertProductionSheetWritable();

    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const queueHeader = headers.includes("Очередность") ? "Очередность" : "Очередь";

    orders.forEach(order => {
      this.updateFields(order.key, {
        [queueHeader]: order.queue,
        "Начало": order.startDate,
        "Окончание": order.finishDate,
        "Конструктор": order.constructor
      });
    });

  }

    /**
   * Сохранение рассчитанных дат
   */
  static saveCalendar(orders) {

    const sheet = this.getSheet(CONFIG.SHEETS.ORDERS);

    const values = sheet.getDataRange().getValues();

    if (values.length < 2) return;

    const headers = values[0];

    const contractCol = headers.indexOf("Договор");
    const clientCol = headers.indexOf("Клиент");
    const productCol = headers.indexOf("Изделие");

    const startCol = headers.indexOf("Начало");
    const finishCol = headers.indexOf("Окончание");

    if (
      contractCol === -1 ||
      clientCol === -1 ||
      productCol === -1 ||
      startCol === -1 ||
      finishCol === -1
    ) {
      throw new Error("Не найдены необходимые колонки.");
    }

    const index = {};

    for (let r = 1; r < values.length; r++) {

      const key = [
        values[r][contractCol],
        values[r][clientCol],
        values[r][productCol]
      ].join("|");

      index[key] = r + 1;

    }

    orders.forEach(order => {

      const key = [
        order.contract,
        order.client,
        order.product
      ].join("|");

      const row = index[key];

      if (!row) return;

      sheet.getRange(row, startCol + 1)
        .setValue(order.startDate);

      sheet.getRange(row, finishCol + 1)
        .setValue(order.finishDate);

    });

  }
  }
