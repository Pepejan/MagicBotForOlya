// Побудова Excel-файлу з накладної — так, як сестра робила вручну:
// бере її файл з прайсом (data/price-list.xlsx) і в аркуш "Аркуш1" вставляє
// весь .mmo з клітинки A1, рядок за рядком, у тому самому порядку.
// Далі в рядках товарів, як і раніше:
//   N — термін придатності з прайсу (VLOOKUP)
//   T — діюча ціна з прайсу (VLOOKUP)
//   W — сума = T * P (ціна × кількість)
// Жодних кольорів і груп — оформлення як в оригінальному файлі.

import ExcelJS from 'exceljs';
import { isItemRow, parseNumber, type Invoice, type InvoiceItem } from './mmo-parser.ts';
import { findPriceSheet, type PriceList } from './price-list.ts';

export interface Report {
  fileName: string;
  buffer: Uint8Array;
  missingInPriceList: InvoiceItem[]; // товари, яких немає в прайсі (або в них не вказана ціна)
}

const DATA_SHEET = 'Аркуш1';

// Явний формат, щоб дата виглядала однаково в будь-якому Excel, Google Таблицях і на телефоні.
const DATE_FORMAT = 'dd.mm.yyyy';

// Ширини колонок як в оригінальному "Аркуш1" (якщо в шаблоні аркуша немає).
const DEFAULT_WIDTHS: Record<string, number> = { A: 13.33, N: 12.33, T: 10.33, U: 11.55, W: 9.11 };

// Номери колонок Excel (з одиниці)
const COL = { code: 1, expiry: 14, quantity: 16, price: 20, invoiceSum: 21, sum: 23 } as const;

export async function buildReport(invoice: Invoice, priceList: PriceList, templatePath: string): Promise<Report> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const priceSheetRef = quoteSheetName(findPriceSheet(workbook).name);
  const sheet = recreateDataSheet(workbook);
  const missingInPriceList: InvoiceItem[] = [];

  invoice.rows.forEach((cols, i) => {
    const row = sheet.getRow(i + 1);
    cols.forEach((raw, c) => {
      const cell = row.getCell(c + 1);
      cell.value = toCellValue(raw);
      if (cell.value instanceof Date) cell.numFmt = DATE_FORMAT;
    });

    // Перші два рядки (шапка документа) не чіпаємо, навіть якщо вони схожі на товар.
    if (i < 2 || !isItemRow(cols)) return;

    const r = row.number;
    const code = cols[COL.code - 1]!.trim();
    const entry = priceList.get(code);
    const quantity = parseNumber(cols[COL.quantity - 1]!);
    if (entry?.price == null) {
      missingInPriceList.push(invoice.items.find((item) => item.code === code)!);
    }

    // result — це значення, яке Excel покаже одразу, ще до перерахунку формул.
    const notFound: ExcelJS.CellErrorValue = { error: '#N/A' };
    row.getCell(COL.expiry).value = {
      formula: `VLOOKUP(A${r},${priceSheetRef}!$A:$D,4,0)`,
      result: entry ? (entry.expiry ?? 0) : notFound,
    };
    row.getCell(COL.price).value = {
      formula: `VLOOKUP(A${r},${priceSheetRef}!$A:$D,3,0)`,
      result: entry ? (entry.price ?? 0) : notFound,
    };
    row.getCell(COL.sum).value = {
      formula: `T${r}*P${r}`,
      result: entry ? Math.round(quantity * (entry.price ?? 0) * 100) / 100 : notFound,
    };

    // Не 'mm-dd-yy': це вбудований формат Excel "коротка дата", і він залежить від мови
    // системи — на телефоні чи в англомовному Excel виходило 10-01-30 замість 01.10.2030.
    row.getCell(COL.expiry).numFmt = DATE_FORMAT;
    row.getCell(COL.price).numFmt = '0.00';
    row.getCell(COL.invoiceSum).numFmt = '0.00';
    row.getCell(COL.sum).numFmt = '0.00';
  });

  widenDateColumns(sheet);

  // Відкривати файл одразу на "Аркуш1".
  const activeTab = workbook.worksheets.indexOf(sheet);
  workbook.views = [
    { x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab, visibility: 'visible' },
  ];

  const buffer = new Uint8Array(await workbook.xlsx.writeBuffer());
  return {
    fileName: `Накладна_${safeFileName(invoice.number)}_${invoice.date}.xlsx`,
    buffer,
    missingInPriceList,
  };
}

/** Замінює "Аркуш1" у шаблоні на чистий, зберігаючи ширини колонок. */
function recreateDataSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet {
  const old = workbook.getWorksheet(DATA_SHEET);
  const widths: Record<string, number> = { ...DEFAULT_WIDTHS };
  if (old) {
    old.columns?.forEach((col) => {
      if (col.letter && col.width) widths[col.letter] = col.width;
    });
    workbook.removeWorksheet(old.id);
  }

  const sheet = workbook.addWorksheet(DATA_SHEET);
  for (const [letter, width] of Object.entries(widths)) {
    sheet.getColumn(letter).width = width;
  }
  return sheet;
}

/**
 * Якщо дата не влазить у колонку, Excel показує "#######".
 * Тому кожну колонку, де є дата (сама або результат формули), робимо не вужчою за MIN_DATE_WIDTH.
 */
const MIN_DATE_WIDTH = 11.5; // "01.10.2030" + невеликий запас

function widenDateColumns(sheet: ExcelJS.Worksheet): void {
  sheet.columns.forEach((column) => {
    let hasDate = false;
    column.eachCell?.((cell) => {
      const value = cell.value;
      const result = value !== null && typeof value === 'object' && 'formula' in value ? value.result : value;
      if (result instanceof Date) hasDate = true;
    });
    if (hasDate && (column.width ?? 0) < MIN_DATE_WIDTH) {
      column.width = MIN_DATE_WIDTH;
    }
  });
}

/**
 * Перетворює текст з .mmo на значення комірки так, як це зробив би Excel при вставці:
 * числа — числами, дати — датами, решта — текстом.
 */
function toCellValue(raw: string): ExcelJS.CellValue {
  const text = raw.trim();
  if (text === '') return null;

  const date = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(text);
  if (date) return new Date(Date.UTC(Number(date[3]), Number(date[2]) - 1, Number(date[1])));

  // "155", "1 060,08", "0" — але не "0000000000" чи IBAN (там провідні нулі / літери)
  if (/^-?(0|[1-9]\d{0,2}([\s ]\d{3})*|[1-9]\d*)(,\d+)?$/.test(text)) return parseNumber(text);

  return text;
}

/** У формулах назву аркуша беремо в лапки: 'Прайс-лист'. */
function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function safeFileName(s: string): string {
  return s.replace(/[\\/:*?"<>|]/g, '_');
}
