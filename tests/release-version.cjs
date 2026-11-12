const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const context=vm.createContext({
 CONFIG:{APP_VERSION:'2026.11.12.001'},
 PropertiesService:{getScriptProperties:()=>({getProperty:()=>null})},
 SheetService:{resetOrderContexts(){},getActionOrderRows(){return [];}},
 PerformanceService:{start(){},finish(){}},
 OrderService:{getOrders:section=>({section,orders:[]}),assignOrderConstructor:()=>({success:true})}
});
vm.runInContext(fs.readFileSync('src/Code.js','utf8'),context);
for(const version of ['2026.11.12.001','2026.11.09.001']){
 context.version=version;
 assert.equal(vm.runInContext('getOrders("constructors",version).appVersion',context),version);
 assert.equal(vm.runInContext('runQueueAction("assignOrderConstructor",["key","constructor"],"constructors",version).snapshot.appVersion',context),version);
}
for(const version of ['','unknown','2026.11.08.001']){
 context.version=version;
 assert.throws(()=>vm.runInContext('getOrders("constructors",version)',context),/APP_VERSION_MISMATCH/);
}
console.log('Coordinated release preserves old/current snapshot contracts and rejects unknown versions');
