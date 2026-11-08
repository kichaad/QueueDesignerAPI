const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');

// Encoding adapter for the Apps Script Blob used in these local tests.
const cp1251Decoder = new TextDecoder('windows-1251');
const cp1251Map = new Map(Array.from({length:256},(_,byte)=>[cp1251Decoder.decode(Uint8Array.of(byte)),byte]));
function newBlob() {
 let bytes;
 return {
  setDataFromString(text, charset) {
   assert.equal(charset,'windows-1251');
   bytes=Uint8Array.from([...text].map(c=>cp1251Map.get(c) ?? 63));return this;
  },
  getBytes:()=>bytes,
  getDataAsString:charset=>{assert.equal(charset,'windows-1251');return cp1251Decoder.decode(bytes);},
 };
}

function fixture() {
  const orders = [['ID','Number','Supplier','Status'], ['a','MDM-1','МДМ','Счёт получен'], ['b','MDM-2','МДМ','Счёт получен'], ['c','Other','Blum','Счёт получен']];
  const lines = [['ID','Order'], ['l1','a','m1','36','Client','Desk',' A B ', 'Article',2,'шт'], ['l2','b','m2','37','Client','Shelf','A B','Article',3,'шт'], ['l3','b','m3','37','Client','Shelf','ZERO01','Other',0.5,'м']];
  let writes = 0;
  const sheet = rows => ({ getRange:(r,c)=>({setNumberFormat(){return this;},setValue(value){writes++;rows[r-1][c-1]=value;}}), getLastRow:()=>rows.length,getDataRange:()=>({getValues:()=>rows}), getRangeList:refs=>({setNumberFormat(){return this;},setValue:value=>{
    writes++; for (const ref of refs) { const col=ref.charCodeAt(0)-65, row=Number(ref.slice(1))-1;rows[row][col]=value; }
  }}) });
  const orderSheet = sheet(orders), lineSheet=sheet(lines), cache=new Map();
  const context=vm.createContext({CONFIG:{SHEETS:{SUPPLIER_ORDERS:'orders',SUPPLIER_ORDER_LINES:'lines'}},SheetService:{getSpreadsheet:()=>({getSheetByName:name=>name==='orders'?orderSheet:lineSheet})},Utilities:{newBlob,DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(_,s)=>crypto.createHash('sha256').update(s).digest(),base64Encode:b=>Buffer.from(b).toString('base64'),getUuid:()=>crypto.randomUUID(),formatDate:()=> '20261105_120000'},Session:{getScriptTimeZone:()=> 'Europe/Moscow'},LockService:{getDocumentLock:()=>({waitLock(){},releaseLock(){}})},CacheService:{getUserCache:()=>({get:k=>cache.get(k),put:(k,v)=>cache.set(k,v)})}});
  vm.runInContext(fs.readFileSync('src/SupplierOrderService.js','utf8')+'\nthis.Service=SupplierOrderService;',context);
  const Service=context.Service;
  Service.ensureOrdersSheet=()=>orderSheet;Service.ensureLinesSheet=()=>lineSheet;
  return {Service,orders,lines,cache,get writes(){return writes;}};
}

test('default, per-line exceptions, clearing and maximum date survive reads',()=>{
 const f=fixture();
 f.lines.push(['l4','a','m4','36','Client','Desk',' A B ', 'Article',1,'шт']);
 f.Service.saveDeliveryDates('a',{date:'2026-12-10'});
 assert.equal(f.Service.getOrders().find(o=>o.id==='a').deliveryDate,'2026-12-10');
 f.Service.saveDeliveryDates('a',{date:'2027-01-05',lineIds:['l4']});
 f.Service.saveDeliveryDates('a',{date:'2026-12-15'});
 let details=f.Service.getOrderDetails('a');
 assert.equal(details.order.deliveryDate,'2027-01-05');
 assert.equal(details.lines.length,2);
 assert.equal(details.detailedLines.find(l=>l.lineId==='l1').deliveryDate,'2026-12-15');
 f.Service.saveDeliveryDates('a',{date:'',lineIds:['l4']});
 details=f.Service.getOrderDetails('a');
 assert.equal(details.order.deliveryDate,'2026-12-15');assert.equal(details.lines.length,1);
 f.Service.saveDeliveryDates('a',{date:''});
 assert.equal(f.Service.getOrders().find(o=>o.id==='a').deliveryDateMissingCount,2);
});
test('old rows are readable, status and ownership validation prevent writes',()=>{
 const f=fixture();assert.equal(f.Service.getOrders()[0].deliveryDate,'');
 for(const date of ['2026-02-30','2026-13-01','10.12.2026','abc']) assert.throws(()=>f.Service.saveDeliveryDates('a',{date}));
 assert.throws(()=>f.Service.saveDeliveryDates('a',{date:'2026-12-01',lineIds:['l2']}));
 f.orders[1][3]='Сформирован';assert.throws(()=>f.Service.saveDeliveryDates('a',{date:'2026-12-01'}));
 assert.equal(f.writes,0);
 for(const status of ['Счёт получен','Выгружен','Оплачен','Товар получен']){f.orders[1][3]=status;assert.doesNotThrow(()=>f.Service.saveDeliveryDates('a',{date:'2028-02-29'}));}
});
