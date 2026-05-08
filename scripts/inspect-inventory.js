const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const dir = path.join(__dirname, '..', 'inventory');
const files = fs.readdirSync(dir).filter(f => f.endsWith('.xlsx'));

for (const f of files) {
  console.log(`\n=================================================`);
  console.log(`ФАЙЛ: ${f}`);
  console.log(`=================================================`);
  const wb = XLSX.readFile(path.join(dir, f));

  console.log('\n--- Листы ---');
  for (const name of wb.SheetNames) {
    const sh = wb.Sheets[name];
    const ref = sh['!ref'] || '(пусто)';
    console.log(`  "${name}"  ${ref}`);
  }

  console.log('\n--- Заголовки/маркеры по листам ---');
  for (const name of wb.SheetNames) {
    const sh = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: null });
    rows.slice(0, 5).forEach((r, i) => {
      const filled = r.filter(v => v !== null && v !== '').map(v => String(v).slice(0, 90));
      if (filled.length) console.log(`  [${name}] row ${i}: ${JSON.stringify(filled)}`);
    });
  }

  console.log('\n--- Поиск всех "квартал N" ---');
  const found = new Set();
  for (const name of wb.SheetNames) {
    const sh = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: null });
    rows.forEach(r => r.forEach(c => {
      if (typeof c === 'string') {
        const m = [...c.matchAll(/(?:квартал|кв\.?)\s*[№#]?\s*(\d+)/gi)];
        m.forEach(x => found.add(parseInt(x[1])));
      }
    }));
  }
  console.log(`  Кварталы упоминаются: ${[...found].sort((a, b) => a - b).join(', ') || 'нет'}`);

  // В свод (если есть) ищем структуру: квартал, сорт, га, кустов
  const svodName = wb.SheetNames.find(n => n.toLowerCase().includes('свод'));
  if (svodName) {
    console.log(`\n--- "${svodName}" (целиком, непустые строки) ---`);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[svodName], { header: 1, defval: null });
    rows.forEach((r, i) => {
      const filled = r.filter(v => v !== null && v !== '');
      if (filled.length) console.log(`  row ${i}: ${JSON.stringify(filled.map(v => String(v).slice(0, 70)))}`);
    });
  }
}
