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
  const sheet = rows => ({ getRange:(r,c)=>({setValue:value=>{writes++;rows[r-1][c-1]=value;}}), getLastRow:()=>rows.length,getDataRange:()=>({getValues:()=>rows}), getRangeList:refs=>({setNumberFormat:()=>{},setValue:value=>{
    writes++; for (const ref of refs) { const col=ref.charCodeAt(0)-65, row=Number(ref.slice(1))-1;rows[row][col]=value; }
  }}) });
  const orderSheet = sheet(orders), lineSheet=sheet(lines), cache=new Map();
  const context=vm.createContext({CONFIG:{SHEETS:{SUPPLIER_ORDERS:'orders',SUPPLIER_ORDER_LINES:'lines'}},SheetService:{getSpreadsheet:()=>({getSheetByName:name=>name==='orders'?orderSheet:lineSheet})},Utilities:{newBlob,DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(_,s)=>crypto.createHash('sha256').update(s).digest(),base64Encode:b=>Buffer.from(b).toString('base64'),getUuid:()=>crypto.randomUUID(),formatDate:()=> '20261105_120000'},Session:{getScriptTimeZone:()=> 'Europe/Moscow'},LockService:{getDocumentLock:()=>({waitLock(){},releaseLock(){}})},CacheService:{getUserCache:()=>({get:k=>cache.get(k),put:(k,v)=>cache.set(k,v)})}});
  vm.runInContext(fs.readFileSync('src/SupplierOrderService.js','utf8')+'\nthis.Service=SupplierOrderService;',context);
  const Service=context.Service;
  Service.ensureOrdersSheet=()=>orderSheet;Service.ensureLinesSheet=()=>lineSheet;
  return {Service,orders,lines,cache,get writes(){return writes;}};
}

test('CSV has no header/BOM/trailing delimiters, combines articles and keeps internal spaces',()=>{
 const f=fixture(),file=f.Service.prepareMdmExport(['a','b']);
 assert.equal(file.csv,'A B;5\r\nZERO01;0,5');assert.equal(file.articleCount,2);assert.equal(file.orderCount,2);
 assert.equal(f.writes,0);assert.equal(f.orders[1][3],'Счёт получен');
 f.Service.completeMdmExport(file.token);assert.equal(f.orders[1][3],'Выгружен');assert.equal(f.orders[2][3],'Выгружен');assert.equal(f.orders[3][3],'Счёт получен');
 assert.doesNotThrow(()=>f.Service.completeMdmExport(file.token));
});

test('only explicitly selected whole orders enter CSV and status update',()=>{
 const f=fixture(),file=f.Service.prepareMdmExport(['a']);assert.equal(file.csv,'A B;2');
 f.Service.completeMdmExport(file.token);assert.equal(f.orders[2][3],'Счёт получен');
});

test('rejects wrong supplier/status, missing orders, invalid articles and quantities',()=>{
 for(const ids of [[],['missing'],['c']])assert.throws(()=>fixture().Service.prepareMdmExport(ids));
 for(const value of ['', 'BAD;CODE', 'BAD\nCODE']){const f=fixture();f.lines[1][6]=value;assert.equal(f.Service.prepareMdmExport(['a']).issues.length,1);assert.equal(f.writes,0);}
 for(const value of [0,-1,'bad',Infinity]){const f=fixture();f.lines[1][8]=value;assert.equal(f.Service.prepareMdmExport(['a']).issues.length,1);}
 const f=fixture();f.orders[1][3]='Оплачен';assert.throws(()=>f.Service.prepareMdmExport(['a']));
});

test('500-article limit is enforced before writing',()=>{
 const f=fixture();f.lines.splice(1);for(let i=0;i<501;i++)f.lines.push(['l'+i,'a','','','','','ART'+i,'',1,'шт']);
 assert.throws(()=>f.Service.prepareMdmExport(['a']),/500/);assert.equal(f.writes,0);

});

test('changed source, changed status and expired token cannot mark orders exported',()=>{
 const f=fixture(),file=f.Service.prepareMdmExport(['a']);f.lines[1][8]=99;
 assert.throws(()=>f.Service.completeMdmExport(file.token),/изменились/);assert.equal(f.writes,0);
 const g=fixture(),file2=g.Service.prepareMdmExport(['a']);g.orders[1][3]='Оплачен';assert.throws(()=>g.Service.completeMdmExport(file2.token));assert.equal(g.writes,0);
 const h=fixture();assert.throws(()=>h.Service.completeMdmExport('expired'),/истекла/);assert.equal(h.writes,0);
});


test('matches the supplied MDM CSV sample byte-for-byte',()=>{
 const expected = 'P1903171 AB;1\r\nK020.M00L/RU;2\r\n6 34200 20 ZN;10\r\nTE93;1\r\nDA50MID;1\r\nK-075 PR;6\r\nD/15 CR.CH;1\r\nF-EX1000/650/1250 SL;10\r\nGL3.152A.2000.WH PR;1';
 const f=fixture();f.lines.splice(1);
 expected.split('\r\n').forEach((text,i)=>{const [article,quantity]=text.split(';');f.lines.push(['line'+i,'a','','','','',article,'',Number(quantity),'шт']);});
 assert.equal(f.Service.prepareMdmExport(['a']).csv,expected);
});


test('same article quantities are combined despite different or empty units',()=>{
 for(const unit of ['м', '', 'компл']) {
  const f=fixture();f.lines[2][9]=unit;
  const file=f.Service.prepareMdmExport(['a','b']);
  assert.equal(file.csv,'A B;5\r\nZERO01;0,5');
  assert.equal(f.writes,0);
  f.Service.completeMdmExport(file.token);
  assert.equal(f.orders[1][3],'Выгружен');assert.equal(f.orders[2][3],'Выгружен');
 }
});


test('all problematic lines include order, product, name, reason and editable article',()=>{
 const f=fixture();f.lines[1][6]='';f.lines[2][6]='BAD;CODE';f.lines[3][8]=0;
 const result=f.Service.prepareMdmExport(['a','b']);
 assert.equal(result.issues.length,3);assert.equal(result.token,undefined);assert.equal(f.writes,0);
 assert.equal(result.issues[0].orderNumber,'MDM-1');assert.equal(result.issues[0].lineId,'l1');
 assert.equal(result.issues[0].name,'Article');assert.ok(result.issues[0].product.includes('Desk'));
 assert.equal(result.issues[0].articleError,true);assert.equal(result.issues[2].articleError,false);
});

test('article correction preserves leading zeros, quantities and status; invalidates old export',()=>{
 const f=fixture();const file=f.Service.prepareMdmExport(['a']);
 f.Service.setLineArticle('a',['l1'],' 001ABC ',' A B ');
 assert.equal(f.lines[1][6],'001ABC');assert.equal(f.lines[1][8],2);assert.equal(f.orders[1][3],'Счёт получен');
 assert.throws(()=>f.Service.completeMdmExport(file.token),/изменились/);
 assert.equal(f.Service.prepareMdmExport(['a']).csv,'001ABC;2');
 assert.throws(()=>f.Service.setLineArticle('a',['l1'],'NEW',' A B '),/изменились/);
 assert.throws(()=>f.Service.setLineArticle('a',['l2'],'NEW','A B'),/изменились/);
 for(const value of ['', 'BAD;ARTICLE', '=FORMULA'])assert.throws(()=>f.Service.setLineArticle('a',['l1'],value,'001ABC'));
});

test('fixing a missing article allows previously blocked CSV preparation',()=>{
 const f=fixture();f.lines[1][6]='';assert.equal(f.Service.prepareMdmExport(['a']).issues.length,1);
 f.Service.setLineArticle('a',['l1'],'FIXED','');
 assert.equal(f.Service.prepareMdmExport(['a']).csv,'FIXED;2');
});


test('order list flags problematic articles and clears warning after correction',()=>{
 const f=fixture();f.lines[1][6]='';f.lines[2][6]='BAD;ARTICLE';
 let orders=f.Service.getOrders();assert.equal(orders.find(o=>o.id==='a').articleIssueCount,1);assert.equal(orders.find(o=>o.id==='b').articleIssueCount,1);
 f.Service.setLineArticle('a',['l1'],'001','');
 orders=f.Service.getOrders();assert.equal(orders.find(o=>o.id==='a').articleIssueCount,0);
});


test('Cyrillic CSV is delivered as Windows-1251 bytes without BOM',()=>{
 const f=fixture();f.lines[1][6]='ФПС 087-23ЧР';
 const file=f.Service.prepareMdmExport(['a']);
 const bytes=Buffer.from(file.base64,'base64');
 assert.equal(file.encoding,'windows-1251');
 assert.equal(cp1251Decoder.decode(bytes),'ФПС 087-23ЧР;2');
 assert.equal(bytes[0],0xD4);assert.notEqual(bytes[0],0xD0);
 assert.equal(f.writes,0);
});

test('unsupported Unicode cannot silently become question marks',()=>{
 const f=fixture();f.lines[1][6]='CODE😀';
 assert.throws(()=>f.Service.prepareMdmExport(['a']),/Windows-1251/);assert.equal(f.writes,0);
});
