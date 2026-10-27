const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage();await p.goto('file://'+process.argv[2]);await p.waitForLoadState('networkidle');
await p.pdf({path:process.argv[3],format:'A4',printBackground:true,margin:{top:'16mm',bottom:'16mm',left:'15mm',right:'15mm'}});await b.close();})();
