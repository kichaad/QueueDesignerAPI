const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
let savedFile, rows = [], insertions = 0, failAppend = true;
const sheet = { getLastRow: () => rows.length + 1, getDataRange: () => ({ getValues: () => [[], ...rows] }), appendRow: row => { if (failAppend) { failAppend = false; throw Error('sheet unavailable'); } rows.push(row); } };
const ctx = vm.createContext({
  Drive: { Files: {
    generateIds: () => ({ ids: ['reservedDriveId123456'] }),
    insert: (resource) => { insertions++; if (savedFile) throw Error('file already exists'); savedFile = resource; return resource; },
    get: id => { if (!savedFile || id !== savedFile.id) throw Error('missing'); return savedFile; }
  } },
  SupplierInvoiceService: { ensureInvoicesSheet: () => sheet, ensureInvoiceFolder: () => ({ getId: () => 'folder' }) },
  GoogleDriveService: { getSafeFileName: name => name },
  Utilities: { base64Decode: s => Buffer.from(s, 'base64'), newBlob: () => ({}) }
});
vm.runInContext(fs.readFileSync('src/TimewebDataGateway.js', 'utf8') + '\nthis.mirror=TimewebSheetMirror;', ctx);
const change = { id: '3df32199-f6c8-459b-b334-a2fd777e0923', orderId: 'order', orderNumber: 'N', driveFileId: 'reservedDriveId123456', sha256: 'a'.repeat(64), name: 'invoice.pdf', base64: 'JVBERi0=', number: '42', uploadedAt: '2026-01-01' };
assert.equal(ctx.mirror.reserveInvoiceFileId().id, change.driveFileId);
assert.throws(() => ctx.mirror.applyInvoice(change), /sheet unavailable/);
ctx.mirror.applyInvoice(change); ctx.mirror.applyInvoice(change);
assert.equal(rows.length, 1); assert.equal(rows[0][5], change.driveFileId); assert.equal(insertions, 2);
console.log('Invoice retry reuses reserved Drive ID and creates one invoice row');
