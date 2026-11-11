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
        rows.forEach((row, index) => { this.rows[start - 2 + index] = row.map(v => typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v); });
      }
    };
  }
}
const supply = new Sheet(29), plate = new Sheet(19), suppliers = new Sheet(6), routes = new Sheet(7);
const sheets = { supply, plate, suppliers, routes };
const context = vm.createContext({
  CONFIG: { SHEETS: { SUPPLY: 'supply', PLATE_MATERIALS: 'plate', SUPPLIERS: 'suppliers', CONSUMABLES: 'routes' } },
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
assert.throws(() => apply('6', [row('supply', 'duplicate', null)]), /дублирующий/);
assert(!locked);
properties.set('TIMEWEB_SHEET_SYNC_ENABLED', '');
assert.throws(() => apply('7', []), /SYNC_DISABLED/);
console.log('Sheet mirror: upsert/delete, literal text, retry, sequence fence, validation and lock release verified');
