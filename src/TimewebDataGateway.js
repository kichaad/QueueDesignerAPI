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
    return values.slice(1).map(row => fields.map((index, field) => {
      const value = row[index];
      return value instanceof Date && field === 3 ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd') : value;
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
