// Локальна перевірка без Telegram:
//   npm run report -- "шлях/до/накладної.mmo"
// Поруч з'явиться готовий .xlsx — зручно, щоб подивитися результат чи налагодити.

import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { buildReport } from './excel-report.ts';
import { decodeMmo, parseMmo } from './mmo-parser.ts';
import { loadPriceList } from './price-list.ts';
import { PRICE_LIST_PATH } from './config.ts';

const input = process.argv[2];
if (!input) {
  console.error('Використання: npm run report -- <файл.mmo>');
  process.exit(1);
}

const invoice = parseMmo(decodeMmo(await readFile(input)));
const priceList = await loadPriceList(PRICE_LIST_PATH);
const report = await buildReport(invoice, priceList, PRICE_LIST_PATH);

const outPath = path.join(path.dirname(input), report.fileName);
await writeFile(outPath, report.buffer);

console.log(`Готово: ${outPath}`);
console.log(`Товарів: ${invoice.items.length}, сума за накладною: ${invoice.totalWithVat}`);
if (report.missingInPriceList.length > 0) {
  console.log('Немає в прайсі:', report.missingInPriceList.map((i) => i.code).join(', '));
}
