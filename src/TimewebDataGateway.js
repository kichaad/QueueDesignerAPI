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
