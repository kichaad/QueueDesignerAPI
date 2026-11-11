const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync('src/Code.js','utf8');let until=Date.now()+60000;
const c=vm.createContext({PropertiesService:{getScriptProperties:()=>({getProperty:()=>String(until)})}});
const fn=source.match(/function assertSupplierOrderWritable\(\) \{[\s\S]*?\n\}/);assert(fn);vm.runInContext(fn[0],c);
assert.throws(()=>c.assertSupplierOrderWritable(),/MAINTENANCE/);until=Date.now()-1;assert.doesNotThrow(()=>c.assertSupplierOrderWritable());
for(const method of ['createSupplierOrder','addSupplierOrderMaterials','deleteSupplierOrder','setSupplierOrderStatusFast','attachSupplierOrderInvoice'])assert.match(source,new RegExp('function '+method+'\\([^\\n]*\\) \\{\\n  assertSupplierOrderWritable\\(\\);'));
console.log('Supplier order write maintenance is time limited');
