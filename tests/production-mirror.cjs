const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');

function fixture(enabled=true,duplicate=false) {
  const state={
    headers:['Договор','Клиент','Изделие','Менеджер','Этап','Статус','Очередность'],
    rows:[['44RM26','Клиент','Кухня','Менеджер','ТЗ','',10]]
  };
  if(duplicate)state.rows.push([...state.rows[0]]);
  const range=(row,column,height=1,width=1)=>({
    getValues(){return Array.from({length:height},(_,i)=>Array.from({length:width},(_,j)=>
      row+i===1 ? state.headers[column+j-1] : state.rows[row+i-2]?.[column+j-1]??''));},
    setValue(value) {
      if(row===1)state.headers[column-1]=value;
      else state.rows[row-2][column-1]=value;
    }
  });
  const sheet={
    getLastRow(){return state.rows.length+1;},
    getDataRange(){return {getValues:()=>[state.headers,...state.rows].map(r=>[...r])};},
    getRange:range,
    getLastColumn(){return state.headers.length;}
  };
  const store={
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>k==='TIMEWEB_PRODUCTION_MIRROR_ENABLED' && enabled?'enabled':''})},
    SheetService:{
      getSpreadsheet:()=>({getSheetByName:()=>sheet}),
      ensureOrderColumn(name){
        if(!state.headers.includes(name)){state.headers.push(name);state.rows.forEach(r=>r.push(''));}
        return sheet;
      },
      appendOrder(order){
        if(state.rows.some(r=>r.slice(0,3).join('|')===
          [order.contract,order.client,order.product].join('|')))throw Error('Duplicate order');
        const row=Array(state.headers.length).fill('');
        row[0]=order.contract;row[1]=order.client;row[2]=order.product;row[3]=order.manager;
        state.rows.push(row);
      }
    },
    CONFIG:{SHEETS:{ORDERS:'Изделия'}}
  };
  vm.runInNewContext(fs.readFileSync('src/TimewebProductionMirror.js','utf8')+
    '\nthis.receiver=TimewebProductionMirror;',store);
  return {apply:p=>store.receiver.apply(p),state};
}
const packet={
 schema:1,event:'upsert',productionOrderId:543,
 fields:{'ID заказа на производство':543,'Договор':'44RM26','Клиент':'Клиент',
         'Изделие':'Кухня','Менеджер':'Менеджер'}
};
let app=fixture();
assert.equal(app.apply(packet).productionOrderId,543);
assert.equal(app.state.rows[0][7],543);
assert.equal(app.apply(packet).existing,true);
assert.equal(app.state.rows.length,1);
const richer=structuredClone(packet);
richer.fields['Этап']='КД';
richer.fields['Статус']='В работе';
richer.fields['Очередность']=25;
app.apply(richer);
assert.equal(app.state.rows[0][4],'КД');
assert.equal(app.state.rows[0][5],'В работе');
assert.equal(app.state.rows[0][6],25);
app=fixture();
let fresh=structuredClone(packet);fresh.fields['Изделие']='Шкаф';
assert.equal(app.apply(fresh).existing,false);
assert.equal(app.state.rows.length,2);
app=fixture(false);
assert.throws(()=>app.apply(packet),/SYNC_DISABLED/);
app=fixture(true,true);
assert.throws(()=>app.apply(packet),/SYNC_CONFLICT/);
app=fixture();
app.apply(packet);
let conflict=structuredClone(packet);conflict.productionOrderId=600;
conflict.fields['ID заказа на производство']=600;
assert.throws(()=>app.apply(conflict),/SYNC_CONFLICT/);
console.log('Production mirror: matching, append, idempotence, fencing and conflict tests passed');
