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
    orders: read(CONFIG.SHEETS.SUPPLIER_ORDERS, 10, [9]),
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
  static manage(action) {
    if (!['status', 'export', 'enable', 'disable'].includes(action)) throw new Error('BAD_REQUEST|Unknown mirror action');
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
      route: { name: CONFIG.SHEETS.CONSUMABLES, width: 7 }
    };
    // Validate the entire packet before making any sheet changes.
    packet.changes.forEach(change => {
      const definition = change && Object.prototype.hasOwnProperty.call(definitions, change.table) && definitions[change.table];
      if (!definition || typeof change.id !== 'string' || !change.id || Array.from(change.id).length > 191 ||
          (change.row !== null && (!Array.isArray(change.row) || change.row.length > definition.width ||
            change.row.some(value => value !== null && !['string', 'number', 'boolean'].includes(typeof value))))) {
        throw new Error('BAD_REQUEST|Некорректная строка зеркалирования.');
      }
      if (change.row !== null && ['supply', 'plate'].includes(change.table) && String(change.row[0]) !== change.id) {
        throw new Error('BAD_REQUEST|ID строки не совпадает.');
      }
      if (change.table === 'supplier' && change.row !== null && this.normal(change.row[0]) !== this.normal(change.id)) {
        throw new Error('BAD_REQUEST|Имя поставщика не совпадает.');
      }
      if (change.table === 'route' && (change.row === null || change.row.length !== 4 ||
          this.routeKey(change.row[0], change.row[1]) !== change.id)) {
        throw new Error('BAD_REQUEST|Ключ маршрута не совпадает.');
      }
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
        const definition = definitions[change.table];
        let plan = plans[change.table];
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
          plan = plans[change.table] = { sheet: sheet, rows: rows, width: definition.width, dirty: new Set(), index: index };
        }
        const key = change.table === 'supplier' ? this.normal(change.id) : change.id;
        const matches = plan.index.get(key) || [];
        if (matches.length > 1) throw new Error('SYNC_DUPLICATE_ID|Исправьте дублирующий ID в зеркале.');
        if (change.row === null) {
          if (matches.length) { plan.rows[matches[0]] = Array(plan.width).fill(''); plan.dirty.add(matches[0]); plan.index.delete(key); }
        } else {
          const source = change.table === 'route'
            ? [change.row[0], change.row[1], this.normal(change.row[0]), this.normal(change.row[1]).replace(/\s+/g, ''),
              change.row[2] ? 'Да' : '', change.row[3] ? 'Да' : '', ''] : change.row;
          const row = Array.from({ length: plan.width }, (_, index) => source[index] == null ? '' : source[index]);
          const index = matches.length ? matches[0] : plan.rows.length;
          plan.rows[index] = row;
          plan.index.set(key, [index]);
          plan.dirty.add(index);
        }
        if (plan.rows.length > 20000) throw new Error('TABLE_TOO_LARGE|Зеркало превышает лимит строк.');
      });
      // All plans are validated before the first write. If Sheets fails midway,
      // retry reapplies these absolute rows/deletions without creating new IDs.
      Object.values(plans).forEach(plan => {
        const indices = Array.from(plan.dirty).sort((a, b) => a - b);
        if (!indices.length) return;
        const required = indices[indices.length - 1] + 2;
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
      });
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
