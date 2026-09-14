const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const p = await b.newPage();
  await p.goto('file:///home/user/-/design/identity-verification-spec.html', { waitUntil: 'networkidle' });
  await p.emulateMedia({ media: 'print' });
  await p.pdf({ path: '/home/user/-/design/聖建様_本人確認システム設計書.pdf',
                format: 'A4', printBackground: true,
                margin: { top: '14mm', bottom: '14mm', left: '12mm', right: '12mm' } });
  await b.close();
})();
