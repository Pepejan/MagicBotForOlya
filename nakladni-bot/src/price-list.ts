// Завантаження прайсу: код товару → діюча ціна з ПДВ і термін придатності.
// Береться аркуш "Прайс-лист" (або перший аркуш, якщо такого немає).
// Колонки: A — код, B — назва, C — ціна, D — термін придатності.

import ExcelJS from 'exceljs';

export interface PriceEntry {
  code: string;
  name: string;
  price: number | null; // у прайсі ціна є не для всіх товарів
  expiry: Date | null;
}

/** Ключ — код товару рядком ("15022"), бо в Excel він буває і числом, і текстом. */
export type PriceList = Map<string, PriceEntry>;

const SHEET_NAME = 'Прайс-лист';

export async function loadPriceList(source: string | Uint8Array): Promise<PriceList> {
  const workbook = new ExcelJS.Workbook();
  if (typeof source === 'string') {
    await workbook.xlsx.readFile(source);
  } else {
    // Тип Buffer у exceljs застарілий, тому приводимо явно.
    await workbook.xlsx.load(source as unknown as ExcelJS.Buffer);
  }

  const sheet = findPriceSheet(workbook);

  // Шапка може бути не в першому рядку (зараз вона в 3-му), тому шукаємо її.
  let headerRow = 0;
  sheet.eachRow((row, rowNumber) => {
    if (headerRow === 0 && cellText(row.getCell(1).value).toLowerCase().includes('код')) {
      headerRow = rowNumber;
    }
  });
  if (headerRow === 0) throw new Error('Не знайшов шапку прайсу (колонка "Код ...")');

  const priceList: PriceList = new Map();
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRow) return;

    const code = cellText(row.getCell(1).value);
    if (code === '') return;

    const price = unwrap(row.getCell(3).value);
    const expiry = unwrap(row.getCell(4).value);

    priceList.set(code, {
      code,
      name: cellText(row.getCell(2).value),
      price: typeof price === 'number' ? price : null,
      expiry: expiry instanceof Date ? expiry : null,
    });
  });

  return priceList;
}

/** Аркуш із прайсом: "Прайс-лист", а якщо такого немає — перший аркуш. */
export function findPriceSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet {
  const sheet = workbook.getWorksheet(SHEET_NAME) ?? workbook.worksheets[0];
  if (!sheet) throw new Error('У файлі немає жодного аркуша');
  return sheet;
}

/** Комірка може містити формулу чи форматований текст — дістаємо саме значення. */
function unwrap(value: ExcelJS.CellValue): ExcelJS.CellValue {
  if (value !== null && typeof value === 'object' && !(value instanceof Date)) {
    if ('result' in value) return unwrap(value.result as ExcelJS.CellValue);
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
  }
  return value;
}

function cellText(value: ExcelJS.CellValue): string {
  const v = unwrap(value);
  return v === null || v === undefined ? '' : String(v).trim();
}
