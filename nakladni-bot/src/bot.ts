// Telegram-бот:
//   • надіслали .mmo  → відповідає гарною Excel-таблицею по накладній;
//   • надіслали .xlsx → оновлює прайс (ціни й терміни придатності).

import { copyFile, writeFile } from 'node:fs/promises';
import { Bot, InputFile, type Context } from 'grammy';
import { PRICE_LIST_PATH } from './config.ts';
import { buildReport } from './excel-report.ts';
import { decodeMmo, parseMmo } from './mmo-parser.ts';
import { loadPriceList, type PriceList } from './price-list.ts';

const token = process.env.BOT_TOKEN;
if (!token) {
  console.error('Немає BOT_TOKEN. Створи файл .env за зразком .env.example');
  process.exit(1);
}

// Хто може користуватися ботом: MY_ID, SISTER_ID і EXTRA_IDS (кілька ID через кому) з .env.
// Якщо всі порожні, доступ мають усі (зручно на старті, щоб дізнатися свій ID).
const allowedUserIds = new Set(
  [process.env.MY_ID, process.env.SISTER_ID, ...(process.env.EXTRA_IDS ?? '').split(',')]
    .map((id) => id?.trim())
    .filter((id): id is string => Boolean(id))
    .map(Number),
);
console.log(`Доступ мають: ${allowedUserIds.size === 0 ? 'всі' : `${allowedUserIds.size} користувач(і)`}`);

let priceList: PriceList = await loadPriceList(PRICE_LIST_PATH);
console.log(`Прайс завантажено: ${priceList.size} позицій`);

const bot = new Bot(token);

// --- Доступ ---
bot.use(async (ctx, next) => {
  const userId = ctx.from?.id;
  if (allowedUserIds.size === 0 || (userId !== undefined && allowedUserIds.has(userId))) {
    return next();
  }
  await ctx.reply(`⛔ Немає доступу. Твій ID: ${userId} — передай його власнику бота.`);
});

bot.command(['start', 'help'], (ctx) =>
  ctx.reply(
    [
      '👋 Привіт! Я перетворюю накладні з 1С на гарні Excel-таблиці.',
      '',
      '📄 Надішли файл *.mmo*: у відповідь отримаєш таблицю в тому ж порядку, з цінами й термінами з прайсу.',
      '💰 Надішли файл *.xlsx* з аркушем «Прайс-лист», і я оновлю прайс.',
      '',
      `Зараз у прайсі ${priceList.size} позицій.`,
      `Твій ID: \`${ctx.from?.id}\``,
    ].join('\n'),
    { parse_mode: 'Markdown' },
  ),
);

bot.on('message:document', async (ctx) => {
  const fileName = ctx.message.document.file_name ?? '';
  try {
    if (/\.mmo$/i.test(fileName)) {
      await handleInvoice(ctx);
    } else if (/\.xlsx$/i.test(fileName)) {
      await handlePriceList(ctx);
    } else {
      await ctx.reply('Я розумію лише .mmo (накладна) або .xlsx (прайс).');
    }
  } catch (err) {
    console.error(`Помилка з файлом "${fileName}":`, err);
    await ctx.reply(`❌ Не вийшло обробити "${fileName}": ${err instanceof Error ? err.message : err}`);
  }
});

bot.on('message', (ctx) => ctx.reply('Надішли мені файл накладної .mmo 📄'));

bot.catch((err) => console.error('Помилка бота:', err));

// --- Обробники ---

async function handleInvoice(ctx: Context): Promise<void> {
  const invoice = parseMmo(decodeMmo(await downloadDocument(ctx)));
  const report = await buildReport(invoice, priceList, PRICE_LIST_PATH);

  const lines = [
    `✅ Накладна № ${invoice.number} від ${invoice.date}`,
    `Товарів: ${invoice.items.length}, сума з ПДВ: ${formatMoney(invoice.totalWithVat)} грн`,
  ];
  if (report.missingInPriceList.length > 0) {
    lines.push('', '⚠️ Немає в прайсі або без ціни:');
    for (const item of report.missingInPriceList) lines.push(`• ${item.code}: ${item.name}`);
  }

  await ctx.replyWithDocument(new InputFile(report.buffer, report.fileName), {
    caption: lines.join('\n').slice(0, 1024), // ліміт підпису в Telegram
  });
}

async function handlePriceList(ctx: Context): Promise<void> {
  const bytes = await downloadDocument(ctx);

  // Спершу перевіряємо, що файл справді читається як прайс, і лише тоді замінюємо старий.
  const newPriceList = await loadPriceList(bytes);
  if (newPriceList.size === 0) throw new Error('у файлі не знайдено жодної позиції');

  await copyFile(PRICE_LIST_PATH, PRICE_LIST_PATH.replace(/\.xlsx$/, '.backup.xlsx'));
  await writeFile(PRICE_LIST_PATH, bytes);
  priceList = newPriceList;

  await ctx.reply(`💰 Прайс оновлено: ${priceList.size} позицій.`);
}

async function downloadDocument(ctx: Context): Promise<Uint8Array> {
  const file = await ctx.getFile();
  const res = await fetch(`https://api.telegram.org/file/bot${token}/${file.file_path}`);
  if (!res.ok) throw new Error(`не вдалося завантажити файл (${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

function formatMoney(x: number): string {
  return x.toLocaleString('uk-UA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

console.log('Бот запущено. Ctrl+C, щоб зупинити.');
await bot.start();
