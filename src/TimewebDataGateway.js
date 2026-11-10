/** Fixed, read-only data gateway for migration of the supplier order list.
 * This is intentionally not registered with the RPC dispatcher yet.
 * Preserve the actual GAS UI until the PHP projection is verified against it.
 */
function readSupplierOrderTableData() {
  const spreadsheet = SheetService.getSpreadsheet();
  const timezone = Session.getScriptTimeZone();
  const read = (name, count, dateColumns) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, count).getValues();
    return rows.map(row => row.map((value, index) => {
      if (!(value instanceof Date)) return value;
      return dateColumns.includes(index) ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd') : value.toISOString();
    }));
  };
  return {
    schema: 1,
    orders: read(CONFIG.SHEETS.SUPPLIER_ORDERS, 10, [9]),
    lines: read(CONFIG.SHEETS.SUPPLIER_ORDER_LINES, 13, [12]),
    suppliers: read(CONFIG.SHEETS.SUPPLIERS, 3, [])
  };
}
