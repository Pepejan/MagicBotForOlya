// Розбір .mmo файлу — це вивантаження "Видаткової накладної" з 1С.
//
// Формат (текст у кодуванні Windows-1251, поля розділені табом):
//   рядок 1: РАСХОДНАЯ_НАКЛАДНАЯ  <ЄДРПОУ> <ЄДРПОУ> версия_3
//   рядок 2: шапка документа — номер, дата, постачальник, ..., суми
//   рядок 3: "Заказ ..." (нам не потрібен)
//   далі:    по одному рядку на товар

export interface InvoiceItem {
  code: string;
  name: string;
  quantity: number;
  priceNoVat: number; // ціна за одиницю без ПДВ (як у накладній)
  sumWithVat: number; // сума рядка з ПДВ (як у накладній)
}

export interface Invoice {
  number: string;
  date: string; // лишаємо рядком "18.11.2026" — нам його лише показати
  supplier: string;
  totalNoVat: number;
  totalWithVat: number;
  items: InvoiceItem[];
  rows: string[][]; // усі рядки файлу як є (розбиті по табу) — щоб вставити їх у Excel без змін
}

// Номери колонок (з нуля). Саме ці позиції в "Аркуш1" відповідали
// колонкам A, B, P, T, U — так сестра раніше вставляла дані вручну.
const ITEM_COL = {
  code: 0, // A
  name: 1, // B
  quantity: 15, // P
  priceNoVat: 19, // T
  sumWithVat: 20, // U
} as const;

const HEADER_COL = {
  number: 0,
  date: 1,
  supplier: 3,
  totalNoVat: 11,
  totalWithVat: 12,
} as const;

/** 1С зберігає файл у Windows-1251, тож звичайний toString('utf8') дасть кракозябри. */
export function decodeMmo(bytes: Uint8Array): string {
  return new TextDecoder('windows-1251').decode(bytes);
}

export function parseMmo(text: string): Invoice {
  const lines = text.split(/\r?\n/).filter((line) => line.trim() !== '');

  if (!lines[0]?.startsWith('РАСХОДНАЯ_НАКЛАДНАЯ')) {
    throw new Error('Це не схоже на видаткову накладну з 1С (.mmo)');
  }

  const rows = lines.map((line) => line.split('\t'));
  const header = rows[1] ?? [];

  const items = rows
    .slice(2)
    .filter(isItemRow)
    .map(
      (cols): InvoiceItem => ({
        code: cols[ITEM_COL.code]!.trim(),
        name: cols[ITEM_COL.name]!.trim(),
        quantity: parseNumber(cols[ITEM_COL.quantity]!),
        priceNoVat: parseNumber(cols[ITEM_COL.priceNoVat]!),
        sumWithVat: parseNumber(cols[ITEM_COL.sumWithVat]!),
      }),
    );

  if (items.length === 0) {
    throw new Error('У накладній не знайдено жодного товару');
  }

  return {
    number: header[HEADER_COL.number]?.trim() ?? '',
    date: header[HEADER_COL.date]?.trim() ?? '',
    supplier: header[HEADER_COL.supplier]?.trim() ?? '',
    totalNoVat: parseNumber(header[HEADER_COL.totalNoVat] ?? '0'),
    totalWithVat: parseNumber(header[HEADER_COL.totalWithVat] ?? '0'),
    items,
    rows,
  };
}

/**
 * Рядок товару — це той, де є всі колонки, а перша — числовий код.
 * Так автоматично пропускаються шапка, "Заказ ..." та інші службові рядки.
 */
export function isItemRow(cols: string[]): boolean {
  return cols.length > ITEM_COL.sumWithVat && /^\d+$/.test(cols[ITEM_COL.code]!.trim());
}

/** "1 060,08" → 1060.08. 1С ставить пробіл (часто нерозривний) між тисячами і кому замість крапки. */
export function parseNumber(raw: string): number {
  const value = Number(raw.replace(/[\s ]/g, '').replace(',', '.'));
  if (Number.isNaN(value)) {
    throw new Error(`Не вдалося прочитати число: "${raw}"`);
  }
  return value;
}
