const fs=require('node:fs'),vm=require('node:vm'),test=require('node:test'),assert=require('node:assert/strict');
const ctx=vm.createContext({});
vm.runInContext(fs.readFileSync('src/SupplyService.js','utf8')+'\nthis.Service=SupplyService;',ctx);
const Service=ctx.Service;
test('five supply states and priority red > lilac > yellow > green',()=>{
 const cases=[
  [[], 'none'],[['Новый'],'red'],[['Заказ поставщику'],'red'],[[''],'red'],
  [['Счёт получен'],'lilac'],[['Счет получен'],'lilac'],
  [['Оплачено'],'yellow'],[['Товар получен'],'green'],
  [['Оплачено','Товар получен'],'yellow'],
  [['Счёт получен','Оплачено','Товар получен'],'lilac'],
  [['Новый','Счёт получен','Оплачено','Товар получен'],'red'],
  [['Заказ поставщику','Счёт получен'],'red'],
  [['Товар получен','Товар получен'],'green'],
 ];
 for(const [statuses,state] of cases)assert.equal(Service.getMasterSupplySignal(statuses,false).state,state,JSON.stringify(statuses));
});
test('LDSP passed-saw exception still applies only when positions are absent',()=>{
 assert.equal(Service.getMasterSupplySignal([],true).state,'green');
 assert.equal(Service.getMasterSupplySignal(['Новый'],true).state,'red');
 assert.equal(Service.getMasterSupplySignal(['Счёт получен'],true).state,'lilac');
});
test('Excel signal cells use saturated fills, black text and borders for all five states',()=>{
 const text=fs.readFileSync('src/OrderService.js','utf8');
 const start=text.indexOf('        masterSupplyColumns.forEach((group, index) => {');
 const end=text.indexOf('        sheet.setColumnWidth(11, 280);',start);
 assert.ok(start>0&&end>start);
 const calls=[];
 const range={setBackgrounds:v=>{calls.push(['colors',v]);return range;},setFontColor:v=>{calls.push(['font',v]);return range;},setFontWeight:()=>range,setBorder:(...v)=>{calls.push(['border',v]);return range;}};
 const source=['none','red','lilac','yellow','green'].map(state=>({supplySignals:{hardware:{state}}}));
 vm.runInNewContext(text.slice(start,end),{masterSupplyColumns:[{key:'hardware'}],source,sheet:{setColumnWidth(){},getRange:()=>range},SpreadsheetApp:{BorderStyle:{SOLID:'solid'}}});
 assert.deepEqual(JSON.parse(JSON.stringify(calls[0][1])),[['#ffffff'],['#ff3030'],['#d43dff'],['#ffe600'],['#00d45a']]);
 assert.equal(calls[1][1],'#000000');assert.equal(calls[2][1][6],'#000000');
});
