/** Fixed, read-only data gateway for migration of the supplier order list.
 * Exposed only through the authenticated TimeWeb RPC dispatcher.
 * Preserve the actual GAS UI until the PHP projection is verified against it.
 */
function readSupplierOrderTableData() {
  // Transitional maintenance stays on the existing writer until its own migration.
  SupplierOrderService.reconcileMaterialStatuses();
  const spreadsheet = SheetService.getSpreadsheet();
  const timezone = Session.getScriptTimeZone();
  const read = (name, count, dateColumns) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    if (sheet.getLastRow() > 20001) throw new Error('TABLE_TOO_LARGE|Слишком много строк для загрузки.');
    const width = Math.min(count, sheet.getMaxColumns());
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
    return rows.map(row => row.map((value, index) => {
      if (!(value instanceof Date)) return value;
      return dateColumns.includes(index) ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd') : value.toISOString();
    }));
  };
  return {
    schema: 1,
    appVersion: CONFIG.APP_VERSION,
    orders: read(CONFIG.SHEETS.SUPPLIER_ORDERS, 11, [9]),
    lines: read(CONFIG.SHEETS.SUPPLIER_ORDER_LINES, 13, [12]),
    suppliers: read(CONFIG.SHEETS.SUPPLIERS, 3, [])
  };
}

/** Raw, bounded source for the Supply and Plate tables.
 * It intentionally contains no presentation projection. */
function readSupplyTableData() {
  // Transitional compatibility with the current Google UI. This maintenance
  // will move with the status writer before SQL reads are enabled.
  SupplierOrderService.reconcileMaterialStatuses();
  const spreadsheet = SheetService.getSpreadsheet();
  const timezone = Session.getScriptTimeZone();
  const read = (name, width, dateColumns) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    if (sheet.getLastRow() > 20001) throw new Error('TABLE_TOO_LARGE|Слишком много строк для загрузки.');
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(width, sheet.getMaxColumns())).getValues();
    return rows.map(row => row.map((value, index) => value instanceof Date
      ? (dateColumns.includes(index) ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd') : value.toISOString())
      : value));
  };
  const orderDirectory = () => {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ORDERS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(value => String(value || '').trim());
    const fields = ['Договор', 'Клиент', 'Изделие', 'Дедлайн'].map(name => headers.indexOf(name));
    if (fields.some(index => index < 0)) throw new Error('ORDER_SCHEMA|Не найдены колонки изделий.');
    return values.slice(1).map(row => fields.map((index) => {
      const value = row[index];
      // SupplyService returns this value through JSON.stringify, which uses an
      // ISO timestamp for a Date. Preserve that wire contract exactly.
      return value instanceof Date ? value.toISOString() : value;
    }));
  };
  return {
    schema: 1,
    appVersion: CONFIG.APP_VERSION,
    supply: read(CONFIG.SHEETS.SUPPLY, 29, []),
    plate: read(CONFIG.SHEETS.PLATE_MATERIALS, 19, []),
    suppliers: read(CONFIG.SHEETS.SUPPLIERS, 6, []),
    orderDirectory: orderDirectory()
  };
}

/** SQL mirror receiver. Only Api.handlePost exposes this class, after secret
 * authentication. No browser-callable global function or business callbacks. */
class TimewebSheetMirror {
  static reserveInvoiceFileId() {
    const result = Drive.Files.generateIds({ maxResults: 1 });
    if (!result || !result.ids || result.ids.length !== 1) throw new Error('DRIVE_ID|Не удалось выделить ID файла.');
    return { id: String(result.ids[0]) };
  }

  static applyInvoice(change) {
    const sheet = SupplierInvoiceService.ensureInvoicesSheet();
    const existing = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues().slice(1).filter(row => String(row[0]) === change.id) : [];
    if (existing.length > 1) throw new Error('INVOICE_DUPLICATE|Повторяющийся ID счёта.');
    if (existing.length === 1) {
      if (String(existing[0][1]) !== change.orderId || String(existing[0][5]) !== change.driveFileId) throw new Error('INVOICE_CONFLICT|ID счёта уже занят.');
      return;
    }
    const folder = SupplierInvoiceService.ensureInvoiceFolder(change.orderNumber);
    const marker = 'dcw-invoice:' + change.id + ':' + change.sha256;
    const name = GoogleDriveService.getSafeFileName(change.name);
    const blob = Utilities.newBlob(Utilities.base64Decode(change.base64), 'application/pdf', name);
    let file;
    try {
      file = Drive.Files.insert({ id: change.driveFileId, title: name, mimeType: 'application/pdf', description: marker, parents: [{ id: folder.getId() }] }, blob);
    } catch (error) {
      try { file = Drive.Files.get(change.driveFileId); } catch (_) { throw error; }
      if (!file || file.description !== marker || file.trashed || (file.labels && file.labels.trashed)) throw error;
    }
    if (!file || String(file.id) !== change.driveFileId) throw new Error('INVOICE_FILE|Не удалось сохранить файл счёта.');
    const literal = value => typeof value === 'string' && /^[=']/.test(value) ? "'" + value : value;
    sheet.appendRow([change.id, change.orderId, change.number, '', '', change.driveFileId, name, 'application/pdf', 'Не проверялся', '', 'Быстрое прикрепление без сверки', change.uploadedAt].map(literal));
  }

  static manage(action) {
    if (!['status', 'export', 'enable', 'disable', 'orders'].includes(action)) throw new Error('BAD_REQUEST|Unknown mirror action');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const properties = PropertiesService.getScriptProperties();
      if (action === 'enable' || action === 'disable') {
        properties.setProperty('TIMEWEB_SHEET_SYNC_ENABLED', action === 'enable' ? 'enabled' : 'disabled');
      }
      const result = {
        schema: 1,
        enabled: properties.getProperty('TIMEWEB_SHEET_SYNC_ENABLED') === 'enabled',
        sequence: properties.getProperty('TIMEWEB_SHEET_SYNC_SEQUENCE') || '0'
      };
      if (action === 'orders') {
        const spreadsheet = SheetService.getSpreadsheet();
        const read = (name, width, dateColumns) => {
          const sheet = spreadsheet.getSheetByName(name);
          if (!sheet || sheet.getLastRow() < 2) return [];
          if (sheet.getLastRow() > 20001) throw new Error('TABLE_TOO_LARGE|Order snapshot too large');
          return sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues().map(row => row.map((value, index) =>
            value instanceof Date && dateColumns.includes(index) ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd') : value));
        };
        result.orders = read(CONFIG.SHEETS.SUPPLIER_ORDERS, 11, [9]);
        result.lines = read(CONFIG.SHEETS.SUPPLIER_ORDER_LINES, 13, [12]);
      }
      if (action === 'export') {
        const spreadsheet = SheetService.getSpreadsheet();
        const definitions = { supply: [CONFIG.SHEETS.SUPPLY, 29], plate: [CONFIG.SHEETS.PLATE_MATERIALS, 19],
          supplier: [CONFIG.SHEETS.SUPPLIERS, 6], route: [CONFIG.SHEETS.CONSUMABLES, 7] };
        result.tables = {};
        Object.entries(definitions).forEach(([key, [name, width]]) => {
          const sheet = spreadsheet.getSheetByName(name);
          if (!sheet || sheet.getLastRow() < 1 || sheet.getMaxColumns() < width) throw new Error('SYNC_SCHEMA|Mirror sheet missing');
          if (sheet.getLastRow() > 20001) throw new Error('TABLE_TOO_LARGE|Mirror sheet too large');
          result.tables[key] = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues() : [];
        });
      }
      return result;
    } finally { lock.releaseLock(); }
  }

  static apply(packet) {
    const properties = PropertiesService.getScriptProperties();
    if (properties.getProperty('TIMEWEB_SHEET_SYNC_ENABLED') !== 'enabled') {
      throw new Error('SYNC_DISABLED|Зеркалирование MySQL выключено.');
    }
    if (!packet || packet.schema !== 1 || !/^[1-9][0-9]{0,19}$/.test(packet.sequence) ||
        !Array.isArray(packet.changes) || !packet.changes.length || packet.changes.length > 20000) {
      throw new Error('BAD_REQUEST|Некорректный пакет зеркалирования.');
    }
    const sequence = String(packet.sequence);
    const definitions = {
      supply: { name: CONFIG.SHEETS.SUPPLY, width: 29 },
      plate: { name: CONFIG.SHEETS.PLATE_MATERIALS, width: 19 },
      supplier: { name: CONFIG.SHEETS.SUPPLIERS, width: 6 },
      route: { name: CONFIG.SHEETS.CONSUMABLES, width: 7 },
      supplier_order: { name: CONFIG.SHEETS.SUPPLIER_ORDERS, width: 11 },
      supplier_order_patch: { name: CONFIG.SHEETS.SUPPLIER_ORDERS, width: 11 },
      supplier_line: { name: CONFIG.SHEETS.SUPPLIER_ORDER_LINES, width: 13 },
      invoice_attachment: { name: CONFIG.SHEETS.SUPPLIER_INVOICES, width: 12 }
    };
    const rowsOf = change => Object.prototype.hasOwnProperty.call(change, 'rows')
      ? change.rows : change.row === null ? [] : [change.row];
    let totalRows = 0;
    // Validate whole groups before writing: duplicate source IDs are legitimate
    // in imported snapshots; each packet replaces the entire SQL-owned group.
    packet.changes.forEach(change => {
      const definition = change && Object.prototype.hasOwnProperty.call(definitions, change.table) && definitions[change.table];
      if (!definition || typeof change.id !== 'string' || !change.id || Array.from(change.id).length > (change.table === 'supplier' ? 512 : 191) ||
          (Object.prototype.hasOwnProperty.call(change, 'rows') && Object.prototype.hasOwnProperty.call(change, 'row'))) {
        throw new Error('BAD_REQUEST|Некорректная строка зеркалирования.');
      }
      if (change.table === 'invoice_attachment') {
        if (!/^[0-9a-f-]{36}$/.test(change.id) || !/^[A-Za-z0-9_-]{10,200}$/.test(change.driveFileId || '') ||
            !/^[0-9a-f]{64}$/.test(change.sha256 || '') || typeof change.base64 !== 'string' || change.base64.length > 7500000 ||
            !/^[A-Za-z0-9+/]*={0,2}$/.test(change.base64) || !['orderId', 'orderNumber', 'name', 'number', 'uploadedAt'].every(key => typeof change[key] === 'string' && change[key].length <= 2000)) {
          throw new Error('BAD_REQUEST|Некорректный файл счёта.');
        }
        return;
      }
      if (change.table === 'supplier_order_patch') {
        const fields = change.fields;
        const valid = {
          status: value => ['Сформирован', 'Счёт получен', 'Выгружен', 'Оплачен', 'Товар получен'].includes(value),
          invoiceNumbers: value => typeof value === 'string' && value.length <= 20000,
          updatedAt: value => typeof value === 'string' && value.length <= 64,
          total: value => value === '' || (typeof value === 'number' && Number.isFinite(value) && value >= 0),
          lineCount: value => Number.isInteger(value) && value >= 0,
          deliveryDate: value => typeof value === 'string' && (value === '' || /^\d{4}-\d{2}-\d{2}$/.test(value)),
          manualTotal: value => typeof value === 'boolean'
        };
        if (!fields || Array.isArray(fields) || typeof fields !== 'object' || !Object.keys(fields).length ||
            Object.entries(fields).some(([name, value]) => !Object.prototype.hasOwnProperty.call(valid, name) || !valid[name](value)) ||
            Object.prototype.hasOwnProperty.call(change, 'row') || Object.prototype.hasOwnProperty.call(change, 'rows')) {
          throw new Error('BAD_REQUEST|Некорректное изменение заказа.');
        }
        return;
      }
      const rows = rowsOf(change);
      if (!Array.isArray(rows)) throw new Error('BAD_REQUEST|Некорректная группа строк.');
      totalRows += rows.length;
      if (totalRows > 20000) throw new Error('BAD_REQUEST|Слишком много строк.');
      rows.forEach(row => {
        if (!Array.isArray(row) || row.length > definition.width ||
            row.some(value => value !== null && !['string', 'number', 'boolean'].includes(typeof value))) {
          throw new Error('BAD_REQUEST|Некорректная строка зеркалирования.');
        }
        if (['supply', 'plate', 'supplier_order', 'supplier_line'].includes(change.table) && String(row[0]) !== change.id) throw new Error('BAD_REQUEST|ID строки не совпадает.');
        if (change.table === 'supplier' && this.normal(row[0]) !== this.normal(change.id)) throw new Error('BAD_REQUEST|Имя поставщика не совпадает.');
        if (change.table === 'route' && (row.length !== 4 || this.routeKey(row[0], row[1]) !== change.id)) throw new Error('BAD_REQUEST|Ключ маршрута не совпадает.');
      });
    });
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const last = properties.getProperty('TIMEWEB_SHEET_SYNC_SEQUENCE') || '0';
      if (sequence.length < last.length || (sequence.length === last.length && sequence <= last)) {
        return { sequence: sequence, duplicate: true };
      }
      const spreadsheet = SheetService.getSpreadsheet();
      const plans = {};
      packet.changes.forEach(change => {
        if (change.table === 'invoice_attachment') return;
        const definition = definitions[change.table];
        let plan = plans[definition.name];
        if (!plan) {
          const sheet = spreadsheet.getSheetByName(definition.name);
          if (!sheet || sheet.getLastRow() < 1 || sheet.getMaxColumns() < definition.width) {
            throw new Error('SYNC_SCHEMA|Лист зеркала не подготовлен.');
          }
          if (sheet.getLastRow() > 20001) throw new Error('TABLE_TOO_LARGE|Зеркало превышает лимит строк.');
          const rows = sheet.getLastRow() > 1
            ? sheet.getRange(2, 1, sheet.getLastRow() - 1, definition.width).getValues() : [];
          const index = new Map();
          rows.forEach((row, position) => {
            const key = change.table === 'supplier' ? this.normal(row[0]) :
              change.table === 'route' ? this.routeKey(row[0], row[1]) : String(row[0]);
            if (!key) return;
            index.set(key, (index.get(key) || []).concat(position));
          });
          plan = plans[definition.name] = { sheet: sheet, rows: rows, width: definition.width, dirty: new Set(), index: index, fields: new Map() };
        }
        const key = change.table === 'supplier' ? this.normal(change.id) : change.id;
        const matches = plan.index.get(key) || [];
        if (change.table === 'supplier_order_patch') {
          if (matches.length !== 1) throw new Error('SYNC_ORDER_MISSING|Заказ не найден или имеет дублирующий ID.');
          const index = matches[0];
          const positions = { status: 3, updatedAt: 5, total: 6, lineCount: 7, invoiceNumbers: 8, deliveryDate: 9, manualTotal: 10 };
          const values = Object.entries(change.fields).map(([name, value]) => [positions[name], value]);
          if (!plan.fields.has(index)) plan.fields.set(index, new Map());
          values.forEach(([column, value]) => { plan.rows[index][column] = value; plan.fields.get(index).set(column, value); });
          return;
        }
        const sources = rowsOf(change);
        const indices = [];
        sources.forEach((input, position) => {
          const source = change.table === 'route'
            ? [input[0], input[1], this.normal(input[0]), this.normal(input[1]).replace(/\s+/g, ''),
              input[2] ? 'Да' : '', input[3] ? 'Да' : '', ''] : input;
          const index = position < matches.length ? matches[position] : plan.rows.length;
          const previous = plan.rows[index] || [];
          const row = Array.from({ length: plan.width }, (_, column) =>
            change.table === 'supplier' && column >= input.length ? previous[column] || '' : source[column] == null ? '' : source[column]);
          plan.rows[index] = row;
          plan.dirty.add(index);
          indices.push(index);
        });
        matches.slice(sources.length).forEach(index => {
          plan.rows[index] = Array(plan.width).fill(''); plan.dirty.add(index);
        });
        if (indices.length) plan.index.set(key, indices); else plan.index.delete(key);
        if (plan.rows.length > 20000) throw new Error('TABLE_TOO_LARGE|Зеркало превышает лимит строк.');
      });
      // All plans are validated before the first write. If Sheets fails midway,
      // retry reapplies these absolute rows/deletions without creating new IDs.
      Object.values(plans).forEach(plan => {
        const indices = Array.from(plan.dirty).sort((a, b) => a - b);
        if (!indices.length && !plan.fields.size) return;
        const required = indices.length ? indices[indices.length - 1] + 2 : 0;
        if (plan.sheet.getMaxRows() < required) {
          plan.sheet.insertRowsAfter(plan.sheet.getMaxRows(), required - plan.sheet.getMaxRows());
        }
        // Write only changed rows, keeping formulas/formatting in unrelated rows.
        for (let offset = 0; offset < indices.length;) {
          const first = indices[offset];
          let last = first;
          while (++offset < indices.length && indices[offset] === last + 1) last = indices[offset];
          const literal = plan.rows.slice(first, last + 1).map(row => row.map(value =>
            typeof value === 'string' && /^[=']/.test(value) ? "'" + value : value));
          plan.sheet.getRange(first + 2, 1, literal.length, plan.width).setValues(literal);
        }
        // Apply order status/invoice patches by cell to preserve formulas and
        // columns owned by other features. A simultaneous full-row update has
        // already included the pending fields in plan.rows.
        plan.fields.forEach((columns, index) => {
          if (plan.dirty.has(index)) return;
          columns.forEach((value, column) => {
            const literal = typeof value === 'string' && /^[=']/.test(value) ? "'" + value : value;
            plan.sheet.getRange(index + 2, column + 1, 1, 1).setValues([[literal]]);
          });
        });
      });
      packet.changes.filter(change => change.table === 'invoice_attachment').forEach(change => this.applyInvoice(change));
      SpreadsheetApp.flush();
      properties.setProperty('TIMEWEB_SHEET_SYNC_SEQUENCE', sequence);
      return { sequence: sequence, duplicate: false };
    } finally {
      lock.releaseLock();
    }
  }

  static normal(value) {
    return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim().toLowerCase().replace(/\s+/g, ' ');
  }

  static routeKey(name, article) {
    return Array.from(this.normal(article).replace(/\s+/g, '') || this.normal(name)).slice(0, 191).join('');
  }
}
