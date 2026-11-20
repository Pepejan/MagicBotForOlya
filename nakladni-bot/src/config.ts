import path from 'node:path';

// Прайс лежить у data/ поруч із кодом. Його можна оновити, просто надіславши боту новий .xlsx.
export const PRICE_LIST_PATH = path.join(import.meta.dirname, '..', 'data', 'price-list.xlsx');
