const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const p = await b.newPage();
  await p.goto('file:///home/user/-/contract/answer-sheet.html', { waitUntil: 'networkidle' });
  await p.emulateMedia({ media: 'print' });
  await p.pdf({ path: '/home/user/-/contract/聖建様_確認事項へのご回答.pdf',
                format: 'A4', printBackground: true,
                margin: { top: '16mm', bottom: '16mm', left: '15mm', right: '15mm' } });
  await b.close();
})();
