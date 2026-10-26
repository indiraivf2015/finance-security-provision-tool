// Assembles the single-file offline tool: dist/Security_Provision_Tool.html
const fs = require('fs'); const path = require('path'); const crypto = require('crypto');
const tpl = fs.readFileSync(path.join(__dirname, 'src/ui.html'), 'utf8');
const exceljs = fs.readFileSync(path.join(__dirname, 'node_modules/exceljs/dist/exceljs.min.js'), 'utf8');
const engine = fs.readFileSync(path.join(__dirname, 'src/engine.js'), 'utf8');
for (const [n, s] of [['exceljs', exceljs], ['engine', engine]]) if (/<\/script/i.test(s)) throw new Error(n + ' contains </script');
const html = tpl.replace('/*__EXCELJS__*/', () => exceljs).replace('/*__ENGINE__*/', () => engine);
fs.mkdirSync(path.join(__dirname, 'dist'), { recursive: true });
const out = path.join(__dirname, 'dist/Security_Provision_Tool.html');
fs.writeFileSync(out, html);
console.log(out, (html.length / 1024).toFixed(0) + ' KB', 'sha256', crypto.createHash('sha256').update(html).digest('hex').slice(0, 16));
