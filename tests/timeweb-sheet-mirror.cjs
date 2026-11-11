const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const properties = new Map([['TIMEWEB_SHEET_SYNC_ENABLED', 'enabled']]);
let fail = false, flushed = false, locked = false;
class Sheet {
  constructor(width, rows = []) { this.width = width; this.rows = rows; this.capacity = 100; }
  getLastRow() { return this.rows.length + 1; }
  getMaxColumns() { return this.width; }
  getMaxRows() { return this.capacity; }
  insertRowsAfter(after, count) { this.capacity += count; }
  getRange(start, col, count, width) {
    return {
      getValues: () => this.rows.slice(start - 2, start - 2 + count).map(r => r.slice()),
      setValues: rows => {
        if (fail) { fail = false; throw Error('simulated Sheets outage'); }
        rows.forEach((row, index) => { const target = this.rows[start - 2 + index] || Array(this.width).fill(''); row.forEach((v, i) => { target[col - 1 + i] = typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v; }); this.rows[start - 2 + index] = target; });
      }
    };
  }
}
const supply = new Sheet(29), plate = new Sheet(19), suppliers = new Sheet(6), routes = new Sheet(7);
const orders = new Sheet(11), lines = new Sheet(13);
const sheets = { supply, plate, suppliers, routes, orders, lines };
const context = vm.createContext({
  CONFIG: { SHEETS: { SUPPLY: 'supply', PLATE_MATERIALS: 'plate', SUPPLIERS: 'suppliers', CONSUMABLES: 'routes', SUPPLIER_ORDERS: 'orders', SUPPLIER_ORDER_LINES: 'lines' } },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => properties.get(k), setProperty: (k, v) => properties.set(k, v) }) },
  LockService: { getScriptLock: () => ({ waitLock: () => { assert(!locked); locked = true; }, releaseLock: () => { locked = false; } }) },
  SheetService: { getSpreadsheet: () => ({ getSheetByName: name => sheets[name] }) },
  SpreadsheetApp: { flush: () => { flushed = true; } }
});
vm.runInContext(fs.readFileSync('src/TimewebDataGateway.js', 'utf8') + '\nthis.mirror = TimewebSheetMirror;', context);
const apply = (sequence, changes) => context.mirror.apply({ schema: 1, sequence, changes });
const row = (table, id, data) => ({ table, id, row: data });
apply('1', [row('supply', 's1', ['s1', '=literal']), row('plate', 'p1', ['p1', "'literal"])]);
assert.equal(supply.rows[0][1], '=literal');
assert.equal(plate.rows[0][1], "'literal");
assert(flushed);
assert.equal(properties.get('TIMEWEB_SHEET_SYNC_SEQUENCE'), '1');
assert(apply('1', [row('supply', 's1', ['s1', 'duplicate'])]).duplicate);
assert.equal(supply.rows[0][1], '=literal');
fail = true;
assert.throws(() => apply('2', [row('supply', 's1', ['s1', 'new'])]), /outage/);
assert.equal(properties.get('TIMEWEB_SHEET_SYNC_SEQUENCE'), '1');
apply('2', [row('supply', 's1', ['s1', 'new'])]);
assert.equal(supply.rows[0][1], 'new');
apply('3', [row('supply', 's1', null)]);
assert(supply.rows.every(r => r[0] === ''));
apply('4', [row('supplier', 'Поставщик', ['Поставщик', '', true]), row('route', 'a42', ['Петля', 'A 42', true, false])]);
assert.equal(routes.rows[0][4], 'Да');
apply('5', [row('supplier', 'ПОСТАВЩИК', ['ПОСТАВЩИК', 'https://example.test'])]);
assert.equal(suppliers.rows.length, 1);
assert(apply('4', [row('supplier', 'Поставщик', null)]).duplicate);
assert.throws(() => apply('6', [row('supply', 'x', ['wrong'])]), /ID/);
assert.throws(() => apply('6', [row('__proto__', 'x', [])]), /Некорректная/);
supply.rows = [['duplicate'], ['duplicate']];
apply('6', [{ table: 'supply', id: 'duplicate', rows: [['duplicate', 'A'], ['duplicate', 'B']] }]);
assert.equal(supply.rows.filter(r => r[0] === 'duplicate').length, 2);
apply('7', [{ table: 'supply', id: 'duplicate', rows: [['duplicate', 'remaining B']] }]);
assert.equal(supply.rows.filter(r => r[0] === 'duplicate').length, 1);
assert.equal(supply.rows[0][1], 'remaining B');
apply('8', [{ table: 'supply', id: 'duplicate', rows: [] }]);
assert(supply.rows.every(r => r[0] === ''));
assert.throws(() => apply('9', [{ table: 'plate', id: 'p1', rows: [['p1'], ['wrong']] }]), /ID/);
assert(!locked);
properties.set('TIMEWEB_SHEET_SYNC_ENABLED', '');
assert.throws(() => apply('7', []), /SYNC_DISABLED/);
console.log('Sheet mirror: upsert/delete, literal text, retry, sequence fence, validation and lock release verified');

assert.equal(context.mirror.manage('status').enabled, false);
assert.equal(context.mirror.manage('enable').enabled, true);
assert.equal(context.mirror.manage('supplier-maintenance-on').supplierMaintenance, true);
assert.equal(context.mirror.manage('supplier-maintenance-off').supplierMaintenance, false);
assert.equal(context.mirror.manage('export').tables.plate[0][0], 'p1');
assert.equal(context.mirror.manage('disable').enabled, false);
assert.throws(() => context.mirror.manage('unknown'), /BAD_REQUEST/);
assert(!locked);
console.log('Mirror administration enable/disable/status/export verified');

properties.set('TIMEWEB_SHEET_SYNC_ENABLED', 'enabled');
apply('10', [row('supplier_order', 'o1', ['o1', 'ЗП-1', 'Blum']), row('supplier_line', 'l1', ['l1', 'o1', 'm1'])]);
assert.equal(orders.rows[0][0], 'o1'); assert.equal(lines.rows[0][1], 'o1');
assert(apply('10', [row('supplier_order', 'o1', ['o1', 'duplicate'])]).duplicate);
assert.throws(() => apply('11', [row('supplier_line', 'l2', ['wrong'])]), /ID/);
console.log('Supplier order mirror validates and idempotently applies order and line IDs');

orders.rows[0][6] = '=preserved-formula'; orders.rows[0][10] = true;
apply('11', [{table:'supplier_order_patch',id:'o1',fields:{status:'Оплачен',invoiceNumbers:'A1',updatedAt:'2026-11-11T10:00:00Z'}}]);
assert.equal(orders.rows[0][3], 'Оплачен'); assert.equal(orders.rows[0][8], 'A1');
assert.equal(orders.rows[0][6], '=preserved-formula'); assert.equal(orders.rows[0][10], true);
assert.throws(() => apply('12', [{table:'supplier_order_patch',id:'o1',fields:{status:'Оплачен',invoiceNumbers:'A1',updatedAt:'x',unknown:'0'}}]), /Некорректное/);
assert.throws(() => apply('12', [{table:'supplier_order_patch',id:'missing',fields:{status:'Оплачен',invoiceNumbers:'A1',updatedAt:'x'}}]), /SYNC_ORDER_MISSING/);
console.log('Order patches preserve unrelated cells and validate the target');
