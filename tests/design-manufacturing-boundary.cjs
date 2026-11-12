const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const ctx=vm.createContext({SheetService:{productionLock:()=>({waitLock(){},releaseLock(){}}),ensureOrderColumn(){}}});
vm.runInContext(fs.readFileSync('src/OrderService.js','utf8')+'\nthis.S=OrderService;',ctx);
const S=ctx.S;let patch;
S.updateRouteFields=(_key,_order,fields)=>{patch=fields;};
S.getAdditionalCorpusBranchConfig=()=>({selectedProperty:'hasCorpusMdf',stageField:'Этап Корпус МДФ',initialStage:'Распил МДФ',statusField:'Статус Корпус МДФ'});
for(const stage of ['ТЗ','Технолог','Проформа','Согласование','Правки','КД','Очередь','Запуск']){
 const order={stage,hasCorpusLdsp:true,hasCorpusMdf:false,hasPaintedLdsp:false,corpusMdfStage:'Распил МДФ',toJSON(){return {...this};}};
 S.getOrderByKey=()=>order;
 S.setCorpusTypes('fixture',{hasCorpusLdsp:true,hasCorpusMdf:true,hasPaintedLdsp:false});
 assert.equal(patch['Этап Корпус МДФ'],undefined,stage+' must not start a branch');
 assert.equal(S.getManufacturingTasks([order]).length,0,stage+' must not create a workshop task');
 order.hasCorpusMdf=true;
 assert.doesNotThrow(()=>S.setCorpusTypes('fixture',{hasCorpusLdsp:true,hasCorpusMdf:false}));
}
const launched={stage:'Сборка',hasCorpusLdsp:true,hasCorpusMdf:false,hasPaintedLdsp:false,toJSON(){return {...this};}};
S.getOrderByKey=()=>launched;
S.setCorpusTypes('fixture',{hasCorpusLdsp:true,hasCorpusMdf:true});
assert.equal(patch['Этап Корпус МДФ'],'Распил МДФ');
launched.hasCorpusMdf=true;launched.corpusMdfStage='Распил МДФ';
assert.equal(S.getManufacturingTasks([launched]).length,1);
assert.throws(()=>S.setCorpusTypes('fixture',{hasCorpusLdsp:true,hasCorpusMdf:false}),/нельзя отменить/);
console.log('Design/release boundary, draft composition changes and launched branch history: OK');
