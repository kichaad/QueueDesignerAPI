const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const headers=['Договор','Клиент','Изделие','Менеджер','Этап','Статус','ID заказа на производство','Комментарий'];
const rows=[headers,['D','Client','Kitchen','Manager','Проформа','',101,''],['P','Client','Production','Manager','Распил','',102,'']];
const props=new Map([['TIMEWEB_CONSTRUCTOR_SQL_ENABLED','enabled']]);let held=false,acquired=0,sent=0;
const sheet={getLastColumn:()=>headers.length,getLastRow:()=>rows.length,getMaxRows:()=>100,
 getDataRange:()=>({getValues:()=>rows.map(r=>r.slice()),getDisplayValues:()=>rows.map(r=>r.map(String))}),
 getRange:(r,c,n=1,w=1)=>({getValues:()=>rows.slice(r-1,r-1+n).map(a=>a.slice(c-1,c-1+w)),
  setNumberFormat(){return this;},getValue:()=>rows[r-1][c-1],setValue:v=>{rows[r-1]??=[];rows[r-1][c-1]=v;},
  setValues:vs=>vs.forEach((a,i)=>a.forEach((v,j)=>{rows[r-1+i][c-1+j]=v;}))})};
const context=vm.createContext({CONFIG:{SHEETS:{ORDERS:'Изделия',DICTIONARIES:'Справочники'}},
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k),setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},
 LockService:{getScriptLock:()=>({waitLock(){assert.equal(held,false,'nested lock must be reentrant');held=true;acquired++;},releaseLock(){held=false;}})},
 SpreadsheetApp:{flush(){}},PerformanceService:{measureWrite:fn=>fn()},
 OrderMapper:{map:x=>x},OrderService:{resolveOverallStage:()=>null},
 NotificationService:{getEventDefinition:()=>({}),deliverSqlEvent:()=>{sent++;return {status:'SENT'};}}
});
vm.runInContext(fs.readFileSync('src/SheetService.js','utf8')+'\nthis.S=SheetService;',context);
context.S.getSpreadsheet=()=>({getSheetByName:()=>sheet});
context.S.getColumns=()=>Object.fromEntries(headers.map((h,i)=>[h,i+1]));
context.S.getOrderContext=key=>({sheet,rows,columns:context.S.getColumns(),rowNumber:key==='prod'?3:2});
context.S.ensureOrderColumn=h=>{if(!headers.includes(h)){headers.push(h);rows.slice(1).forEach(r=>r.push(''));}return sheet;};
vm.runInContext(fs.readFileSync('src/TimewebConstructorSql.js','utf8')+'\nthis.C=TimewebConstructorSql;',context);
const lock=context.S.productionLock();lock.waitLock(100);
assert.throws(()=>context.S.updateFields('design',{'Статус':'В работе'}),/SQL_MASTER/);
assert.equal(held,true);lock.releaseLock();assert.equal(held,false);assert.equal(acquired,1);
context.S.updateFields('prod',{'Статус':'В работе'});assert.equal(rows[2][5],'В работе');
assert.throws(()=>context.S.updateFields('prod',{'Этап':'КД'}),/SQL_MASTER/);
const packet={schema:1,eventId:1,productionOrderId:101,fields:{'ID заказа на производство':101,'Комментарий':'=not-a-formula','Этап':'КД'},notification:{type:'KD_ASSIGNED',before:{},after:{}}};
context.C.apply(packet);assert.equal(rows[1][4],'КД');assert.equal(rows[1][7],"'=not-a-formula");assert.equal(sent,1);
rows[1][4]='Распил'; // Production advances after the SQL handoff.
context.C.apply(packet);assert.equal(rows[1][4],'Распил','duplicate packet overwrote newer production state');assert.equal(sent,1);
assert.throws(()=>context.C.apply({...packet,eventId:2,fields:{'Дедлайн':'bad'}}),/BAD_REQUEST/);
assert.equal(held,false);
console.log('Constructor fence, nested lock, production isolation, literal text and event replay: OK');

rows[1][4]='КД';
context.S.ensureOrderColumn('Очередь производства');const prodCol=headers.indexOf('Очередь производства');rows[2][prodCol]=50;
context.C.apply({schema:1,eventId:2,productionOrderId:101,fields:{'ID заказа на производство':101,'Этап':'Очередь','Очередь производства':10},notification:null});
assert.equal(rows[1][prodCol],60,'production handoff must follow the live production queue');
console.log('SQL handoff allocates a production priority after all existing production orders');
