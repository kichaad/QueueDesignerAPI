const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');

function fixture(frozen=true) {
 const state={
  headers:['Договор','Клиент','Изделие','Менеджер','Этап'],
  rows:[['44RM26','Клиент А','Кухня','Менеджер 1','Проформа'],
        ['44RM26','Клиент А','Кухня','Менеджер 1','Технолог']]
 };
 const sheet={
   getDataRange(){return {getValues:()=>[state.headers,...state.rows].map(r=>[...r])};},
   getLastColumn(){return state.headers.length;},
   getRange(row,col){return {setValue(v){if(row===1)state.headers[col-1]=v;
      else state.rows[row-2][col-1]=v;}};}
 };
 const ctx={
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>{
    if(k==='TIMEWEB_PRODUCTION_LINK_ENABLED')return 'enabled';
    if(k==='TIMEWEB_PRODUCTION_GOOGLE_WRITES_DISABLED')return frozen?'enabled':'';
    return '';
  }})},
  LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
  CONFIG:{SHEETS:{ORDERS:'Изделия'}},
  SheetService:{productionLock:()=>({waitLock(){},releaseLock(){}}),getSpreadsheet:()=>({getSheetByName:()=>sheet}),
    ensureOrderColumn(name){
       if(!state.headers.includes(name)){
         state.headers.push(name);state.rows.forEach(r=>r.push(''));
       }
       return sheet;
    }}
 };
 vm.runInNewContext(fs.readFileSync('src/TimewebProductionMigrationLink.js','utf8')+
   '\nthis.linking=TimewebProductionMigrationLink;',ctx);
 return {link:rows=>ctx.linking.apply(rows),state};
}
const a={productionOrderId:101,sheetRow:2,key:'44RM26|Клиент А|Кухня',
         manager:'Менеджер 1',stage:'Проформа'};
const b={productionOrderId:102,sheetRow:3,key:'44RM26|Клиент А|Кухня',
         manager:'Менеджер 1',stage:'Технолог'};
let f=fixture();
assert.equal(f.link([a,b]).tagged,2);
assert.deepEqual([f.state.rows[0][5],f.state.rows[1][5]],[101,102]);
assert.equal(f.link([a,b]).tagged,2);
f=fixture();
assert.throws(()=>f.link([a,{...b,stage:'КД'}]),/SYNC_CONFLICT/);
assert.equal(f.state.headers.includes('ID заказа на производство'),false);
f=fixture();
assert.throws(()=>f.link([a,{...b,productionOrderId:101}]),/SYNC_CONFLICT/);
f=fixture(false);
assert.throws(()=>f.link([a]),/MIGRATION_NOT_FROZEN/);
console.log('One-time production SQL ID mapping: freeze, duplicates, conflicts, retry passed');
