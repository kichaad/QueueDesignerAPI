const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
let count = 0;
for (const name of fs.readdirSync('src')) {
  if (name.endsWith('.js')) {
    new vm.Script(fs.readFileSync(path.join('src', name), 'utf8'), { filename: name });
    count++;
  }
}
for (const name of ['package.json', 'package-lock.json', '.clasp.json', 'src/appsscript.json']) {
  JSON.parse(fs.readFileSync(name, 'utf8'));
}
if (count !== 26) throw new Error(`Expected 26 server files, found ${count}; review source inventory`);
console.log(`Syntax OK: ${count} server files and 4 JSON files. Google runtime and business behavior are not tested.`);
