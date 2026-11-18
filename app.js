'use strict';

/* Деньги Monee — картина денег бизнеса из таблицы «Баланс ДДС».
   Только чтение. Последние посчитанные цифры (не сам файл) хранятся на этом
   устройстве (localStorage), чтобы приложение открывалось без входа и без
   интернета. Кнопка «Выйти» стирает их. */

const CONFIG = {
  // «Номер приложения» для входа через Google (OAuth Client ID). Не секрет.
  // Проект Google Cloud «Dengi Monee» (monee-dengi), тестовый режим. Если пусто — кнопка входа выключена.
  googleClientId: '135942497577-e5nf1dq48f29qtr261sapubthi53ca42.apps.googleusercontent.com',
  fileName: 'Баланс ДДС Ильгиз.xlsx',
  scope: 'https://www.googleapis.com/auth/drive.readonly',
};

const LS_KEY = 'monee-dengi:settings';
const SS_KEY = 'monee-dengi:session'; // старое место (до 18.11.2026), только стираем
const LAST_KEY = 'monee-dengi:last';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const GSHEET_MIME = 'application/vnd.google-apps.spreadsheet';

/* ---------- мелкие помощники ---------- */

class AppError extends Error {}

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (ch) => (
  { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const NF = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
const NBSP = ' ';
const rub = (v) => (v == null ? '—' : NF.format(Math.round(v)) + NBSP + '₽');
const signedRub = (v) => (Math.round(v) > 0 ? '+' : Math.round(v) < 0 ? '−' : '') + NF.format(Math.abs(Math.round(v))) + NBSP + '₽';
const pct = (v) => (v == null ? '—' : v.toLocaleString('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + NBSP + '%');
const cap = (s) => { const t = String(s).trim(); return t.charAt(0).toUpperCase() + t.slice(1); };
function shortRub(v) {
  const a = Math.abs(v);
  if (a >= 1e6) return (v / 1e6).toLocaleString('ru-RU', { maximumFractionDigits: 1 }) + NBSP + 'млн';
  if (a >= 1e3) return Math.round(v / 1e3) + NBSP + 'тыс';
  return String(Math.round(v));
}

function store(kind) {
  try { return kind === 'local' ? window.localStorage : window.sessionStorage; } catch { return null; }
}
function load(kind, key) {
  try { const s = store(kind)?.getItem(key); return s ? JSON.parse(s) : null; } catch { return null; }
}
function save(kind, key, val) {
  try {
    const st = store(kind); if (!st) return;
    if (val == null) st.removeItem(key); else st.setItem(key, JSON.stringify(val));
  } catch { /* хранилище недоступно — работаем без него */ }
}

/* ---------- месяцы ---------- */

const MONTH_STEMS = { янв: 0, фев: 1, мар: 2, апр: 3, май: 4, мая: 4, июн: 5, июл: 6, авг: 7, сен: 8, окт: 9, ноя: 10, дек: 11 };
const M_NAME = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
const M_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
const M_DAT = ['январю', 'февралю', 'марту', 'апрелю', 'маю', 'июню', 'июлю', 'августу', 'сентябрю', 'октябрю', 'ноябрю', 'декабрю'];
const M_SHORT = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];

function parseMonth(v) {
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) { // дата Excel
      const d = new Date(Math.round((v - 25569) * 864e5));
      return d.getUTCFullYear() * 12 + d.getUTCMonth();
    }
    return null;
  }
  if (typeof v !== 'string') return null;
  const m = v.trim().toLowerCase().match(/^([а-яё]{3,9})\.?\s*[-–.\s]\s*'?(\d{2}|\d{4})$/);
  if (!m) return null;
  const idx = MONTH_STEMS[m[1].slice(0, 3)];
  if (idx == null) return null;
  let y = +m[2]; if (y < 100) y += 2000;
  return y * 12 + idx;
}
const mYear = (k) => Math.floor(k / 12);
const mName = (k) => `${cap(M_NAME[k % 12])} ${mYear(k)}`;
const mGen = (k) => `${M_GEN[k % 12]} ${mYear(k)}`;
const mDat = (k) => M_DAT[k % 12];

/* ---------- чтение xlsx без библиотек ---------- */

async function inflateRaw(bytes) {
  if (typeof DecompressionStream === 'undefined') {
    throw new AppError('Браузер слишком старый и не умеет открывать xlsx. Обнови браузер.');
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function readZip(buffer) {
  const u8 = new Uint8Array(buffer);
  const dv = new DataView(buffer);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new AppError('Это не файл Excel (xlsx). Выбери «Баланс ДДС» в формате xlsx.');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  const entries = new Map();
  for (let k = 0; k < count && p + 46 <= u8.length; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true);
    const xlen = dv.getUint16(p + 30, true);
    const clen = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true);
    entries.set(dec.decode(u8.subarray(p + 46, p + 46 + nlen)), { method, csize, off });
    p += 46 + nlen + xlen + clen;
  }
  return {
    async text(name) {
      const e = entries.get(name);
      if (!e) return null;
      const start = e.off + 30 + dv.getUint16(e.off + 26, true) + dv.getUint16(e.off + 28, true);
      const raw = u8.subarray(start, start + e.csize);
      let data;
      if (e.method === 0) data = raw;
      else if (e.method === 8) data = await inflateRaw(raw);
      else throw new AppError('Файл сжат непривычным способом. Пересохрани его в Excel или Google Таблицах как xlsx.');
      return dec.decode(data);
    },
  };
}

const parseXml = (s) => new DOMParser().parseFromString(s, 'application/xml');
const tags = (node, name) => node.getElementsByTagNameNS('*', name);
function colIndex(ref) {
  let n = 0;
  for (const ch of ref) { const c = ch.charCodeAt(0); if (c < 65 || c > 90) break; n = n * 26 + (c - 64); }
  return n - 1;
}

function norm(s) {
  return String(s ?? '').toLowerCase().replace(/ё/g, 'е').replace(/[ \s]+/g, ' ').trim().replace(/\s*:$/, '');
}
// ключ для поиска строки: без нумерации «1.», «2.1.» и дефиса в начале
const rowKey = (s) => norm(s).replace(/^\d+(\.\d+)*\.?\s*/, '').replace(/^-\s*/, '');
const cleanName = (s) => cap(String(s).trim().replace(/^-\s*/, '').replace(/^\d+(\.\d+)*\.?\s+/, '').replace(/\s*:$/, ''));

async function readWorkbook(buffer, wanted) {
  const zip = readZip(buffer);
  const wbXml = await zip.text('xl/workbook.xml');
  if (!wbXml) throw new AppError('Это не файл Excel (xlsx). Выбери «Баланс ДДС» в формате xlsx.');
  const rels = {};
  const relsXml = await zip.text('xl/_rels/workbook.xml.rels');
  if (relsXml) for (const r of tags(parseXml(relsXml), 'Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
  const sheets = {};
  for (const s of tags(parseXml(wbXml), 'sheet')) {
    const rid = s.getAttribute('r:id')
      || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    let t = rels[rid] || '';
    t = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
    sheets[norm(s.getAttribute('name'))] = t;
  }
  let shared = [];
  const ssXml = await zip.text('xl/sharedStrings.xml');
  if (ssXml) {
    shared = [...tags(parseXml(ssXml), 'si')].map((si) => [...tags(si, 't')]
      .filter((t) => t.parentNode.localName !== 'rPh').map((t) => t.textContent).join(''));
  }
  const out = {};
  for (const name of wanted) {
    const path = sheets[norm(name)];
    if (!path) throw new AppError(`В файле нет вкладки «${name}». Проверь, что выбран «Баланс ДДС».`);
    const text = await zip.text(path);
    if (!text) throw new AppError(`Не получилось прочитать вкладку «${name}».`);
    out[name] = parseSheet(text, shared);
  }
  return out;
}

function parseSheet(text, shared) {
  const doc = parseXml(text);
  const grid = [];
  for (const row of tags(doc, 'row')) {
    const ra = row.getAttribute('r');
    const r = ra ? +ra - 1 : grid.length;
    const arr = grid[r] = [];
    let auto = 0;
    for (const c of row.children) {
      if (c.localName !== 'c') continue;
      const ref = c.getAttribute('r');
      const ci = ref ? colIndex(ref) : auto;
      auto = ci + 1;
      const t = c.getAttribute('t');
      const vEl = tags(c, 'v')[0];
      let val = null;
      if (t === 's') val = vEl ? (shared[+vEl.textContent] ?? null) : null;
      else if (t === 'inlineStr') val = [...tags(c, 't')].map((x) => x.textContent).join('');
      else if (t === 'str') val = vEl ? vEl.textContent : null;
      else if (t === 'b' || t === 'e') val = null;
      else if (vEl) { const n = parseFloat(vEl.textContent); val = Number.isFinite(n) ? n : null; }
      arr[ci] = val;
    }
  }
  return grid;
}

/* ---------- поиск строк по названию ---------- */

class Sheet {
  constructor(name, grid) {
    this.name = name;
    this.grid = grid;
    this.labels = Array.from(grid, (r) => (r && typeof r[0] === 'string' && r[0].trim() ? r[0] : null));
    this.keys = this.labels.map((l) => (l == null ? null : rowKey(l)));
  }
  find(label, { from = 0, to = Infinity, prefix = false, optional = false } = {}) {
    const k = rowKey(label);
    const end = Math.min(to, this.grid.length);
    for (let r = Math.max(0, from); r < end; r++) {
      const x = this.keys[r];
      if (x != null && (prefix ? x.startsWith(k) : x === k)) return r;
    }
    if (optional) return -1;
    throw new AppError(`Не нашёл строку «${label}» на вкладке «${this.name}». Возможно, её переименовали — проверь таблицу. Неверных цифр не показываю.`);
  }
  num(r, c) {
    if (r < 0 || c == null) return null;
    const v = this.grid[r]?.[c];
    return typeof v === 'number' && Number.isFinite(v) ? v : null;
  }
}

function monthHeader(sh, maxRows = 6) {
  let best = null;
  for (let r = 0; r < Math.min(maxRows, sh.grid.length); r++) {
    const row = sh.grid[r] || [];
    const cols = []; const seen = new Set();
    for (let c = 1; c < row.length; c++) {
      const k = parseMonth(row[c]);
      if (k != null && !seen.has(k)) { seen.add(k); cols.push({ c, key: k }); }
    }
    if (!best || cols.length > best.cols.length) best = { r, cols };
  }
  if (!best || best.cols.length < 2) {
    throw new AppError(`На вкладке «${sh.name}» не нашёл строку с месяцами (вида «окт.-26»).`);
  }
  return best;
}

/* ---------- вкладка «Баланс» ---------- */

function parseBalance(sh) {
  const { cols } = monthHeader(sh);
  const rA = sh.find('АКТИВЫ');
  const rP = sh.find('ПАССИВЫ', { from: rA + 1 });
  const R = {
    rA, rP,
    assets: sh.find('Итого', { from: rA + 1, to: rP }),
    cab: sh.find('Личный кабинет', { from: rA, to: rP }),
    goods: sh.find('Товары', { from: rA, to: rP, prefix: true }),
    money: sh.find('Деньги', { from: rA, to: rP }),
    liab: sh.find('Итого', { from: rP + 1 }),
    equity: sh.find('Собственные средства', { from: rP }),
    debt: sh.find('Задолженность', { from: rP }),
    check: sh.find('Сверка баланса', { from: rP }),
  };
  R.credit = sh.find('Кредит', { from: R.debt, to: R.check, prefix: true });
  R.tax = sh.find('Налоги', { from: R.debt, to: R.check });
  const months = cols.filter((m) => sh.num(R.assets, m.c) != null);
  if (!months.length) throw new AppError('На вкладке «Баланс» нет месяца с цифрой в строке «Итого» активов.');
  return { sh, R, months };
}

function balanceSnapshot(b, c, pc) {
  const { sh, R } = b;
  const pair = (r) => ({ v: sh.num(r, c), pv: pc == null ? null : sh.num(r, pc) });
  const heads = [R.cab, R.goods, R.money, R.rP].sort((x, y) => x - y);
  const end = (s) => heads.find((h) => h > s) ?? R.rP;
  const items = (s) => {
    const out = [];
    for (let r = s + 1; r < end(s); r++) {
      const lab = sh.labels[r];
      if (!lab) continue;
      const p = pair(r);
      if (p.v == null && p.pv == null) continue;
      out.push({ name: cleanName(lab), sub: /^\s*-/.test(lab), ...p });
    }
    return out;
  };
  const groups = [];
  for (const it of items(R.cab)) {
    if (it.sub && groups.length) groups.at(-1).items.push(it);
    else groups.push({ ...it, items: [] });
  }
  const credit = pair(R.credit); const tax = pair(R.tax); const debt = pair(R.debt);
  const parts = [{ name: 'Кредиты и займы', ...credit }, { name: 'Налоги', ...tax }];
  const other = {
    name: 'Прочие',
    v: debt.v == null ? null : debt.v - (credit.v ?? 0) - (tax.v ?? 0),
    pv: debt.pv == null ? null : debt.pv - (credit.pv ?? 0) - (tax.pv ?? 0),
  };
  if (Math.abs(other.v ?? 0) >= 1 || Math.abs(other.pv ?? 0) >= 1) parts.push(other);
  return {
    assets: pair(R.assets),
    liabilities: pair(R.liab),
    equity: pair(R.equity),
    debt: { ...debt, parts },
    money: { ...pair(R.money), items: items(R.money) },
    cabinets: { ...pair(R.cab), groups },
    goods: { ...pair(R.goods), items: items(R.goods) },
    check: sh.num(R.check, c),
  };
}

/* ---------- вкладка «ДДС» ---------- */

function parseDDS(sh) {
  let hdr = -1; let bestCount = 0;
  for (let r = 0; r < Math.min(6, sh.grid.length); r++) {
    const n = (sh.grid[r] || []).filter((v) => parseMonth(v) != null).length;
    if (n > bestCount) { bestCount = n; hdr = r; }
  }
  if (hdr < 0) throw new AppError('На вкладке «ДДС» не нашёл строку с месяцами.');
  const top = sh.grid[hdr] || []; const sub = sh.grid[hdr + 1] || [];
  const byMonth = new Map();
  let cur = null;
  for (let c = 1; c < Math.max(top.length, sub.length); c++) {
    const k = parseMonth(top[c]);
    if (k != null) cur = k;
    if (cur == null) continue;
    if (!byMonth.has(cur)) byMonth.set(cur, { total: null, cols: [] });
    const e = byMonth.get(cur);
    if (e.total == null && norm(sub[c]) === 'итого') e.total = c; else e.cols.push(c);
  }
  const rIn = sh.find('Входящий денежный поток', { prefix: true });
  const rOut = sh.find('Исходящий денежный поток', { from: rIn + 1, prefix: true });
  const rEnd = sh.find('Остаток денежных средств на конец периода', { from: rOut + 1, prefix: true });
  const rDiff = sh.find('Разница', { from: rEnd, optional: true });
  const rTrIn = sh.find('перевод между счетами', { from: rIn + 1, to: rOut, prefix: true, optional: true });
  const rTrOut = sh.find('перевод между счетами', { from: rOut + 1, to: rEnd, prefix: true, optional: true });

  return {
    flows(k) {
      const e = byMonth.get(k);
      if (!e || e.total == null) return null;
      const c = e.total;
      const inflow = sh.num(rIn, c); const outflow = sh.num(rOut, c);
      if (inflow == null && outflow == null) return null;
      const expenses = [];
      for (let r = rOut + 1; r < rEnd; r++) {
        if (r === rTrOut || !sh.labels[r]) continue;
        const v = sh.num(r, c);
        if (v != null && v > 0) expenses.push({ name: cleanName(sh.labels[r]), v });
      }
      expenses.sort((a, b) => b.v - a.v);
      return {
        inflow: inflow == null ? null : inflow - (sh.num(rTrIn, c) ?? 0),
        outflow: outflow == null ? null : outflow - (sh.num(rTrOut, c) ?? 0),
        expenses: expenses.slice(0, 10),
      };
    },
    diff(k) {
      if (rDiff < 0) return null;
      const e = byMonth.get(k);
      if (!e) return null;
      const t = sh.num(rDiff, e.total);
      if (t != null) return t;
      const parts = e.cols.map((c) => sh.num(rDiff, c)).filter((v) => v != null);
      return parts.length ? parts.reduce((a, b) => a + b, 0) : null;
    },
  };
}

/* ---------- вкладка «PNL» ---------- */

function parsePNL(sh) {
  const { cols } = monthHeader(sh);
  const byMonth = new Map(cols.map((m) => [m.key, m.c]));
  const rNet = sh.find('За минусом самовыкупов', { prefix: true });
  const rWB = sh.find('Валовая выручка WB', { from: rNet + 1, prefix: true });
  const rOz = sh.find('Валовая выручка Ozon', { from: rNet + 1, prefix: true });
  const rProfit = sh.find('Чистая прибыль', { prefix: true });
  const rMargin = sh.find('Маржинальность по ЧП', { prefix: true, optional: true });
  return {
    month(k) {
      const c = byMonth.get(k);
      if (c == null) return null;
      const revTotal = sh.num(rNet, c); const profit = sh.num(rProfit, c);
      let margin = sh.num(rMargin, c);
      if (margin != null && Math.abs(margin) <= 1.5) margin *= 100;
      if (margin == null && profit != null && revTotal) margin = (profit / revTotal) * 100;
      return { revWB: sh.num(rWB, c), revOzon: sh.num(rOz, c), revTotal, profit, margin };
    },
  };
}

/* ---------- сборка картины ---------- */

function buildModel(grids, source) {
  const b = parseBalance(new Sheet('Баланс', grids['Баланс']));
  const dds = parseDDS(new Sheet('ДДС', grids['ДДС']));
  const pnl = parsePNL(new Sheet('PNL', grids.PNL));
  const last = b.months.at(-1);
  const prev = b.months.length > 1 ? b.months.at(-2) : null;
  const now = balanceSnapshot(b, last.c, prev?.c);
  const series = [];
  for (let k = last.key - 11; k <= last.key; k++) {
    const p = pnl.month(k); const f = dds.flows(k);
    if (!p && !f) continue;
    series.push({
      key: k,
      revWB: p?.revWB ?? null, revOzon: p?.revOzon ?? null, revTotal: p?.revTotal ?? null,
      profit: p?.profit ?? null, margin: p?.margin ?? null,
      inflow: f?.inflow ?? null, outflow: f?.outflow ?? null, expenses: f?.expenses ?? [],
    });
  }
  return {
    v: 1, source, loadedAt: Date.now(),
    month: last.key, prevMonth: prev?.key ?? null,
    now, sverka: { balance: now.check, dds: dds.diff(last.key) }, series,
  };
}

async function modelFromBuffer(buffer, source) {
  const grids = await readWorkbook(buffer, ['Баланс', 'ДДС', 'PNL']);
  return buildModel(grids, source);
}

/* ---------- выдуманные цифры для примера ---------- */

function demoModel() {
  const d = new Date();
  const last = d.getFullYear() * 12 + d.getMonth() - 1;
  let seed = 20261;
  const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
  const r1000 = (v) => Math.round(v / 1000) * 1000 + Math.round(rnd() * 999);
  const pair = (base, swing = 0.12) => ({ v: r1000(base), pv: r1000(base * (1 - swing + rnd() * swing * 2)) });
  const money = [
    ['Расчётный счёт 1', 1.9e6], ['Расчётный счёт 2', 640e3], ['Карта предпринимателя', 210e3],
    ['Счёт маркетплейса', 380e3], ['Карта банка', 95e3], ['Накопительный счёт', 1.2e6], ['Долг карго', -150e3],
  ].map(([name, v]) => ({ name, sub: false, ...pair(v) }));
  const sum = (arr, f) => arr.reduce((a, x) => a + (x[f] ?? 0), 0);
  const grp = (name, list) => {
    const items = list.map(([n, v]) => ({ name: n, sub: true, ...pair(v) }));
    return { name, sub: false, v: sum(items, 'v'), pv: sum(items, 'pv'), items };
  };
  const groups = [
    grp('Ozon', [['В обработке', 820e3], ['Доступно к выводу', 260e3], ['На выводе', 140e3]]),
    grp('WB', [['В обработке (выводится)', 1.1e6], ['Доступно к выводу', 310e3], ['На балансе', 90e3]]),
  ];
  const goodsItems = [['В пути', 2.4e6], ['Едет на ВБ', 600e3], ['На складе на Ozon', 3.1e6], ['На складе на WB', 4.2e6], ['Фулфилмент/склад', 1.7e6]]
    .map(([n, v]) => ({ name: n, sub: true, ...pair(v) }));
  const moneyT = { v: sum(money, 'v'), pv: sum(money, 'pv') };
  const cabT = { v: sum(groups, 'v'), pv: sum(groups, 'pv') };
  const goodsT = { v: sum(goodsItems, 'v'), pv: sum(goodsItems, 'pv') };
  const assets = { v: moneyT.v + cabT.v + goodsT.v, pv: moneyT.pv + cabT.pv + goodsT.pv };
  const parts = [{ name: 'Кредиты и займы', ...pair(3.4e6, 0.05) }, { name: 'Налоги', ...pair(420e3) }];
  const debt = { v: sum(parts, 'v'), pv: sum(parts, 'pv'), parts };
  const equity = { v: assets.v - debt.v, pv: assets.pv - debt.pv };
  const names = ['закупка товаров Китай', 'доставка Китай/РФ', 'рекламные затраты', 'услуги фулфилмента/ЦПТ', 'налоги',
    'ЗП менеджер ВБ', 'ЗП менеджер ОЗОН', 'аренда помещения', 'товары для упаковки товаров', 'дизайн, инфографика',
    'аналитические сервисы', 'связь (тел., инт.)'];
  const series = [];
  for (let k = last - 11; k <= last; k++) {
    const rev = 3.1e6 + rnd() * 2.6e6;
    const wb = rev * (0.48 + rnd() * 0.1); const oz = rev * 0.95 - wb;
    const profit = rev * (rnd() < 0.12 ? -0.03 : 0.05 + rnd() * 0.11);
    const outflow = rev * (0.82 + rnd() * 0.2);
    let left = outflow;
    const expenses = names.map((n, i) => {
      const share = i === 0 ? 0.34 + rnd() * 0.1 : (0.5 / names.length) * (0.4 + rnd() * 1.4);
      const v = Math.min(left, outflow * share); left -= v; return { name: cap(n), v: Math.round(v) };
    }).sort((a, b2) => b2.v - a.v).slice(0, 10);
    series.push({
      key: k, revWB: Math.round(wb), revOzon: Math.round(oz), revTotal: Math.round(rev), profit: Math.round(profit),
      margin: (profit / rev) * 100, inflow: Math.round(rev * (0.9 + rnd() * 0.15)), outflow: Math.round(outflow), expenses,
    });
  }
  return {
    v: 1, source: { type: 'demo', name: 'Выдуманные цифры' }, loadedAt: Date.now(),
    month: last, prevMonth: last - 1,
    now: {
      assets, liabilities: { ...assets }, equity, debt,
      money: { ...moneyT, items: money }, cabinets: { ...cabT, groups }, goods: { ...goodsT, items: goodsItems }, check: 0,
    },
    sverka: { balance: 0, dds: 0 }, series,
  };
}

/* ---------- вход через Google и чтение с Диска ---------- */

const google_ = { token: null, expires: 0, client: null, ready: null };

function gisReady() {
  if (!CONFIG.googleClientId) return Promise.reject(new AppError('Вход через Google ещё не настроен.'));
  if (!google_.ready) {
    google_.ready = new Promise((resolve, reject) => {
      if (window.google?.accounts?.oauth2) { resolve(); return; }
      const s = document.createElement('script');
      s.src = 'https://accounts.google.com/gsi/client';
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => { google_.ready = null; reject(new AppError('Не загрузился вход Google. Проверь интернет или VPN — или открой файл вручную.')); };
      document.head.append(s);
    });
  }
  return google_.ready;
}

function requestToken() {
  return new Promise((resolve, reject) => {
    if (!window.google?.accounts?.oauth2) { reject(new AppError('Вход Google ещё загружается. Нажми ещё раз через пару секунд.')); return; }
    if (!google_.client) {
      google_.client = window.google.accounts.oauth2.initTokenClient({
        client_id: CONFIG.googleClientId, scope: CONFIG.scope, callback: () => {},
      });
    }
    google_.client.callback = (resp) => {
      if (resp.error) { reject(new AppError('Google не дал доступ: ' + resp.error)); return; }
      if (!window.google.accounts.oauth2.hasGrantedAllScopes(resp, CONFIG.scope)) {
        reject(new AppError('Нужно разрешить чтение Google Диска — без этого таблицу не открыть.')); return;
      }
      google_.token = resp.access_token;
      google_.expires = Date.now() + (Number(resp.expires_in) || 3600) * 1000 - 60000;
      resolve();
    };
    google_.client.error_callback = (err) => {
      reject(new AppError(err?.type === 'popup_closed' ? 'Окно входа Google закрыто. Нажми «Войти через Google» ещё раз.'
        : 'Окно входа Google не открылось. Разреши всплывающие окна для этого сайта.'));
    };
    google_.client.requestAccessToken({ prompt: '' });
  });
}

async function drive(url) {
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + google_.token } });
  if (res.status === 401) { google_.token = null; throw new AppError('Вход в Google истёк. Нажми «Войти через Google» ещё раз.'); }
  if (res.status === 403 || res.status === 404) throw Object.assign(new AppError('noaccess'), { noAccess: true });
  if (!res.ok) throw new AppError(`Google Диск ответил ошибкой ${res.status}. Попробуй позже или открой файл вручную.`);
  return res;
}

const NO_ACCESS = () => new AppError(`Нет доступа к файлу «${CONFIG.fileName}» у этого аккаунта Google. Войди аккаунтом, на Диске которого лежит таблица.`);

async function findDriveFile() {
  const fields = 'id,name,mimeType,modifiedTime';
  const settings = load('local', LS_KEY) || {};
  if (settings.fileId) {
    try {
      const r = await drive(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(settings.fileId)}?fields=${fields}&supportsAllDrives=true`);
      return await r.json();
    } catch (e) { if (!e.noAccess) throw e; }
  }
  const q = `name = '${CONFIG.fileName.replace(/'/g, "\\'")}' and trashed = false`;
  const url = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}&fields=files(${fields})&orderBy=modifiedTime desc&pageSize=10&supportsAllDrives=true&includeItemsFromAllDrives=true`;
  let list;
  try { list = await (await drive(url)).json(); } catch (e) { if (e.noAccess) throw NO_ACCESS(); throw e; }
  const files = list.files || [];
  const file = files.find((f) => f.mimeType === XLSX_MIME) || files.find((f) => f.mimeType === GSHEET_MIME);
  if (!file) throw NO_ACCESS();
  save('local', LS_KEY, { ...settings, fileId: file.id });
  return file;
}

async function loadFromGoogle() {
  if (!google_.token || Date.now() > google_.expires) await requestToken();
  showLoading('Ищу таблицу на Google Диске…');
  const file = await findDriveFile();
  showLoading('Читаю таблицу…');
  const base = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}`;
  let res;
  try {
    res = await drive(file.mimeType === GSHEET_MIME
      ? `${base}/export?mimeType=${encodeURIComponent(XLSX_MIME)}`
      : `${base}?alt=media&supportsAllDrives=true`);
  } catch (e) { if (e.noAccess) throw NO_ACCESS(); throw e; }
  const buffer = await res.arrayBuffer();
  return modelFromBuffer(buffer, { type: 'google', name: file.name, modified: file.modifiedTime || null });
}

/* ---------- состояние и экраны ---------- */

const app = document.getElementById('app');
const fileInput = document.getElementById('file-input');
const state = { screen: 'login', model: null, error: null, selected: null, loadingText: '' };

function persist() {
  save('local', LAST_KEY, state.model ? { model: state.model, screen: state.screen, selected: state.selected } : null);
}

function go(screen) {
  state.screen = screen;
  persist();
  render();
  window.scrollTo(0, 0);
}

function showLoading(text) {
  state.loadingText = text;
  if (state.screen !== 'loading') { state.prevScreen = state.screen; state.screen = 'loading'; }
  render();
}

function toast(text) {
  const t = document.getElementById('toast');
  t.textContent = text;
  t.classList.add('show');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.remove('show'), 2600);
}

async function run(task, okText) {
  const back = state.model ? (state.screen === 'loading' ? state.prevScreen : state.screen) : 'login';
  try {
    state.error = null;
    const model = await task();
    state.model = model;
    if (!state.selected || !model.series.some((s) => s.key === state.selected)) state.selected = model.series.at(-1)?.key ?? null;
    go(back === 'login' ? 'now' : back);
    if (okText) toast(okText);
  } catch (e) {
    if (!(e instanceof AppError)) console.warn(e);
    const msg = e instanceof AppError ? e.message : 'Не получилось прочитать таблицу. Проверь, что это «Баланс ДДС» в формате xlsx.';
    if (state.model) { state.screen = back; render(); toast(msg); } else { state.error = msg; go('login'); }
  }
}

function render() {
  const s = state.screen;
  document.body.dataset.screen = s;
  if (s === 'loading') app.innerHTML = viewLoading();
  else if (s === 'now' && state.model) app.innerHTML = viewNow(state.model);
  else if (s === 'months' && state.model) { app.innerHTML = viewMonths(state.model); drawChart(); }
  else app.innerHTML = viewLogin();
}

/* ---------- открытые данные: курсы ЦБ и производственный календарь ----------
   Файл data.json сам обновляет задание GitHub Actions (fetch.py, .github/workflows/update.yml).
   Там только общие открытые данные, ничего личного. */

const market = { status: 'loading', data: null };
const UFA_TZ = 'Asia/Yekaterinburg';
const RATE = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const CUR_SIGN = { CNY: '¥', USD: '$', EUR: '€' };

function ymd(iso) { const [y, m, d] = String(iso).slice(0, 10).split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); }
function dayMonth(iso) { const d = ymd(iso); return `${d.getUTCDate()}${NBSP}${M_GEN[d.getUTCMonth()]}`; }
function whenUfa(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  const day = new Intl.DateTimeFormat('ru-RU', { timeZone: UFA_TZ, day: 'numeric', month: 'long' }).format(d);
  const time = new Intl.DateTimeFormat('ru-RU', { timeZone: UFA_TZ, hour: '2-digit', minute: '2-digit' }).format(d);
  return `${day} в ${time}`;
}
function plural(n, one, few, many) {
  const a = Math.abs(n) % 100; const b = a % 10;
  if (a > 10 && a < 20) return many;
  if (b > 1 && b < 5) return few;
  return b === 1 ? one : many;
}

async function loadMarket() {
  market.status = 'loading'; paintMarket();
  try {
    const res = await fetch(`data.json?t=${Date.now()}`, { cache: 'no-store' });
    if (!res.ok) throw new Error(String(res.status));
    const d = await res.json();
    if (!d || (!d.rates && !d.calendar)) throw new Error('пустой файл');
    market.data = d; market.status = 'ok';
  } catch {
    market.status = 'error';
  }
  paintMarket();
}

function paintMarket() {
  document.querySelectorAll('[data-market]').forEach((el) => { el.innerHTML = viewMarketBody(); });
}

function viewMarket(compact = false) {
  return `<section class="block market${compact ? ' market-compact' : ''}" aria-label="Курсы ЦБ и календарь">
    <h2><span class="key key-market"></span>Курсы ЦБ и календарь</h2>
    <div data-market aria-live="polite">${viewMarketBody()}</div>
  </section>`;
}

function viewMarketBody() {
  if (market.status === 'loading') return '<p class="mk-note">Загружаю курсы и календарь…</p>';
  if (market.status === 'error' || !market.data) {
    return `<div class="mk-error" role="alert">Не получилось загрузить курсы и календарь. Проверь интернет и нажми «Повторить». Остальное приложение работает как обычно.
      <button class="btn mk-retry" data-act="market-retry">Повторить</button></div>`;
  }
  const d = market.data;
  const parts = [];
  const r = d.rates;
  if (r && r.items?.length) {
    const rows = r.items.map((it) => {
      const ch = Math.round(it.change * 100) / 100;
      const dl = ch === 0 ? '<span class="dl">без изменений</span>'
        : `<span class="dl">${ch > 0 ? '▲' : '▼'}${NBSP}${RATE.format(Math.abs(ch))}</span>`;
      return `<li class="row"><span class="row-name"><b class="cur-sign">${esc(CUR_SIGN[it.code] || '')}</b> ${esc(it.name)}</span><span class="row-val">${RATE.format(it.value)}${NBSP}₽</span>${dl}</li>`;
    }).join('');
    parts.push(`<p class="mk-cap">Курс ЦБ на ${esc(dayMonth(r.date))}${r.ok === false ? ` <span class="mk-stale">(источник не ответил, данные от ${esc(whenUfa(r.fetched_at))})</span>` : ''}</p><ul class="ledger">${rows}</ul>`);
  }
  const c = d.calendar;
  if (c && c.workdays_in_month != null) {
    const mon = M_GEN[Number(c.month.slice(5, 7)) - 1];
    const today = c.today_is_workday ? (c.today_is_short ? 'Сегодня сокращённый рабочий день' : 'Сегодня рабочий день') : 'Сегодня выходной';
    const tomorrow = c.tomorrow_is_workday == null ? '' : (c.tomorrow_is_workday ? 'завтра рабочий' : 'завтра выходной');
    const hol = (c.next_holidays || []).map((h) => {
      const span = h.from === h.to ? dayMonth(h.from) : `${dayMonth(h.from)} — ${dayMonth(h.to)}`;
      return `<li class="row"><span class="row-name">${esc(h.name)}</span><span class="row-val mk-dates">${esc(span)}</span><span class="dl">${h.days}${NBSP}${plural(h.days, 'день', 'дня', 'дней')} отдыха</span></li>`;
    }).join('');
    parts.push(`<p class="mk-cap">Производственный календарь${c.ok === false ? ` <span class="mk-stale">(источник не ответил, данные от ${esc(whenUfa(c.fetched_at))})</span>` : ''}</p>
      <ul class="ledger">
        <li class="row"><span class="row-name">Рабочих дней до конца ${esc(mon)}</span><span class="row-val">${c.workdays_left} из ${c.workdays_in_month}</span><span class="dl">${esc(today)}${tomorrow ? `, ${esc(tomorrow)}` : ''}</span></li>
        ${hol}
      </ul>`);
  }
  const age = (Date.now() - new Date(d.updated_at).getTime()) / 36e5;
  const old = age > 48 ? ' <span class="mk-stale">— давно не обновлялось, проверь задание в GitHub Actions</span>' : '';
  parts.push(`<p class="mk-upd">Обновлено ${esc(whenUfa(d.updated_at))} (время Уфы)${old}</p>`);
  return parts.join('');
}

/* ---------- экран «Вход» ---------- */

function viewLogin() {
  const ready = Boolean(CONFIG.googleClientId);
  return `
  <main class="login wrap">
    <svg class="login-mark" viewBox="0 0 120 64" aria-hidden="true">
      <rect x="0" y="6" width="58" height="20" rx="4" class="seg-money"/><rect x="61" y="6" width="22" height="20" rx="4" class="seg-cab"/><rect x="86" y="6" width="34" height="20" rx="4" class="seg-goods"/>
      <rect x="0" y="34" width="40" height="20" rx="4" class="seg-debt"/><rect x="43" y="34" width="77" height="20" rx="4" class="seg-equity"/>
    </svg>
    <h1>Деньги Monee · пульт</h1>
    <p class="lead">Сколько денег на счетах, во что они вложены, сколько бизнес должен и сколько заработал. Цифры берутся из таблицы «Баланс ДДС».</p>
    ${state.error ? `<div class="alert" role="alert">${esc(state.error)}</div>` : ''}
    <div class="login-actions">
      <button class="btn btn-primary" data-act="google" ${ready ? '' : 'disabled aria-describedby="google-note"'}>Войти через Google</button>
      ${ready ? '' : '<p class="note" id="google-note">Вход через Google ещё не настроен. Пока открой файл вручную.</p>'}
      <button class="btn" data-act="file">Открыть файл вручную</button>
      <button class="btn btn-quiet" data-act="demo">Посмотреть на выдуманных цифрах</button>
    </div>
    ${viewMarket(true)}
    <p class="fine">Файл читается только на этом устройстве. Цифры никуда не отправляются; последние запоминаются на этом устройстве, кнопка «Выйти» их стирает. Таблица не меняется.</p>
  </main>`;
}

function viewLoading() {
  return `<main class="loading wrap" aria-busy="true"><div class="loading-bar" aria-hidden="true"><span></span></div><p>${esc(state.loadingText || 'Читаю таблицу…')}</p></main>`;
}

/* ---------- экран «Сейчас» ---------- */

function delta(v, pv, { goodUp = true, prevLabel = '' } = {}) {
  if (v == null || pv == null) return '<span class="dl dl-none"></span>';
  const d = Math.round(v - pv);
  if (d === 0) return `<span class="dl">без изменений${prevLabel}</span>`;
  const good = goodUp ? d > 0 : d < 0;
  return `<span class="dl ${good ? 'dl-good' : 'dl-bad'}">${d > 0 ? '▲' : '▼'}${NBSP}${signedRub(d)}${prevLabel}</span>`;
}

function ledgerRow(it, opts) {
  return `<li class="row${it.sub ? ' row-sub' : ''}"><span class="row-name">${esc(it.name)}</span><span class="row-val">${rub(it.v)}</span>${delta(it.v, it.pv, opts)}</li>`;
}

function sverkaBadge(m) {
  const { balance, dds } = m.sverka;
  const known = [balance, dds].filter((x) => x != null);
  if (!known.length) return { cls: 'check-unknown', text: 'Сверки нет', detail: 'В таблице нет цифр сверки за этот месяц.' };
  const bad = known.some((x) => Math.abs(x) >= 1);
  const parts = [];
  if (balance != null) parts.push(`«Сверка баланса»: ${rub(balance)}`);
  if (dds != null) parts.push(`«Разница» в ДДС: ${rub(dds)}`);
  if (!bad) return { cls: 'check-ok', text: 'Сверка сходится', detail: parts.join(', ') };
  const worst = known.reduce((a, x) => (Math.abs(x) > Math.abs(a) ? x : a), 0);
  return { cls: 'check-bad', text: `Расхождение ${rub(worst)}`, detail: parts.join(', ') };
}

function sourceLine(m) {
  const t = new Date(m.loadedAt).toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
  if (m.source.type === 'demo') return 'Выдуманные цифры для примера';
  const d = new Date(m.loadedAt);
  const today = new Date().toDateString() === d.toDateString();
  const when = today ? `сегодня в ${t}` : `${d.toLocaleDateString('ru-RU', { day: 'numeric', month: 'long' })} в ${t}`;
  if (m.source.type === 'google') return `Из Google Диска, загружено ${when}`;
  return `Из файла «${esc(m.source.name)}», загружено ${when}`;
}

function nav(screen) {
  return `<nav class="dock" aria-label="Действия">
    ${screen === 'now' ? '<button class="btn" data-act="months">Месяцы</button>' : '<button class="btn" data-act="back">Назад</button>'}
    <button class="btn" data-act="refresh">Обновить</button>
    <button class="btn btn-quiet" data-act="logout">Выйти</button>
  </nav>`;
}

function scaleBar(n) {
  const pos = (x) => Math.max(0, x ?? 0);
  const money = pos(n.money.v); const cab = pos(n.cabinets.v); const goods = pos(n.goods.v);
  const other = pos((n.assets.v ?? 0) - money - cab - goods);
  const topTotal = money + cab + goods + other || 1;
  const debt = pos(n.debt.v); const eq = pos(n.equity.v);
  const botTotal = debt + eq || 1;
  const width = Math.max(topTotal, botTotal);
  const seg = (cls, v, total, label) => (v > 0
    ? `<span class="seg ${cls}" style="--w:${((v / width) * 100).toFixed(3)}%" title="${esc(label)}: ${rub(v)}"></span>` : '');
  const share = (v, t) => Math.round((v / t) * 100);
  return `
  <div class="scale" role="img" aria-label="Активы: деньги ${share(money, topTotal)}%, кабинеты ${share(cab, topTotal)}%, товар ${share(goods, topTotal)}%. Источники: долги ${share(debt, botTotal)}%, собственные средства ${share(eq, botTotal)}%.">
    <div class="scale-row"><span class="scale-cap">Что есть</span><div class="bar">${seg('seg-money', money, topTotal, 'Деньги на счетах')}${seg('seg-cab', cab, topTotal, 'Деньги в кабинетах')}${seg('seg-goods', goods, topTotal, 'Товар')}${seg('seg-other', other, topTotal, 'Прочее')}</div></div>
    <div class="scale-row"><span class="scale-cap">Чьё это</span><div class="bar">${seg('seg-debt', debt, botTotal, 'Долги')}${seg('seg-equity', eq, botTotal, 'Собственные средства')}</div></div>
    <ul class="legend">
      <li><i class="seg-money"></i>Счета ${share(money, topTotal)}%</li>
      <li><i class="seg-cab"></i>Кабинеты ${share(cab, topTotal)}%</li>
      <li><i class="seg-goods"></i>Товар ${share(goods, topTotal)}%</li>
      ${other > 0 ? `<li><i class="seg-other"></i>Прочее ${share(other, topTotal)}%</li>` : ''}
      <li><i class="seg-debt"></i>Долги ${share(debt, botTotal)}%</li>
      <li><i class="seg-equity"></i>Свои ${share(eq, botTotal)}%</li>
    </ul>
  </div>`;
}

function viewNow(m) {
  const n = m.now;
  const badge = sverkaBadge(m);
  const prevLabel = m.prevMonth != null ? ` к${NBSP}${mDat(m.prevMonth)}` : '';
  const cabinets = n.cabinets.groups.map((g) => `
      <li class="group"><div class="row row-group"><span class="row-name">${esc(g.name)}</span><span class="row-val">${rub(g.v)}</span>${delta(g.v, g.pv)}</div>
        ${g.items.length ? `<ul>${g.items.map((it) => ledgerRow(it)).join('')}</ul>` : ''}</li>`).join('');
  return `
  <div class="wrap page">
    <header class="top">
      <div class="top-main">
        <p class="brand">Деньги Monee · пульт</p>
        <h1>На конец ${esc(mGen(m.month))}</h1>
        <p class="sub">${sourceLine(m)}</p>
      </div>
      ${nav('now')}
    </header>

    <section class="hero" aria-label="Главные цифры">
      <div class="figs">
        <div class="fig fig-main">
          <span class="fig-label">Активы</span>
          <span class="fig-num">${rub(n.assets.v)}</span>
          ${delta(n.assets.v, n.assets.pv, { prevLabel })}
        </div>
        <div class="fig fig-debt">
          <span class="fig-label">Долги</span>
          <span class="fig-num">${rub(n.debt.v)}</span>
          ${delta(n.debt.v, n.debt.pv, { goodUp: false, prevLabel })}
        </div>
        <div class="fig fig-eq">
          <span class="fig-label">Собственные средства</span>
          <span class="fig-num">${rub(n.equity.v)}</span>
          ${delta(n.equity.v, n.equity.pv, { prevLabel })}
        </div>
      </div>
      <div class="hero-side">
        <div class="check ${badge.cls}">
          <span class="check-dot" aria-hidden="true"></span>
          <span><strong>${esc(badge.text)}</strong><small>${esc(badge.detail)}</small></span>
        </div>
        ${scaleBar(n)}
      </div>
    </section>

    <div class="blocks">
      <section class="block block-money">
        <h2><span class="key key-money"></span>Деньги на счетах <span class="h-sum">${rub(n.money.v)}</span></h2>
        <ul class="ledger">${n.money.items.map((it) => ledgerRow(it)).join('')}</ul>
      </section>
      <section class="block">
        <h2><span class="key key-cab"></span>Деньги в кабинетах <span class="h-sum">${rub(n.cabinets.v)}</span></h2>
        <ul class="ledger">${cabinets}</ul>
      </section>
      <section class="block">
        <h2><span class="key key-goods"></span>Товар по себестоимости <span class="h-sum">${rub(n.goods.v)}</span></h2>
        <ul class="ledger">${n.goods.items.map((it) => ledgerRow({ ...it, sub: false })).join('')}</ul>
      </section>
      <section class="block">
        <h2><span class="key key-debt"></span>Долги <span class="h-sum">${rub(n.debt.v)}</span></h2>
        <ul class="ledger">${n.debt.parts.map((it) => ledgerRow(it, { goodUp: false })).join('')}</ul>
      </section>
      ${viewMarket()}
    </div>
  </div>`;
}

/* ---------- экран «Месяцы» ---------- */

function viewMonths(m) {
  const series = m.series;
  const sel = series.find((s) => s.key === state.selected) || series.at(-1);
  const range = series.length ? `${mName(series[0].key)} — ${mName(series.at(-1).key).toLowerCase()}` : '';
  const cell = (v) => (v == null ? '—' : `${NF.format(Math.round(v))}<span class="cur">${NBSP}₽</span>`);
  const rows = series.slice().reverse().map((s) => `
    <button class="mrow${s.key === sel?.key ? ' is-sel' : ''}" data-month="${s.key}" aria-pressed="${s.key === sel?.key}">
      <span class="c c-month">${esc(mName(s.key))}</span>
      <span class="c c-wb" data-label="WB">${cell(s.revWB)}</span>
      <span class="c c-oz" data-label="Ozon">${cell(s.revOzon)}</span>
      <span class="c c-rev" data-label="Выручка">${cell(s.revTotal)}</span>
      <span class="c c-profit${(s.profit ?? 0) < 0 ? ' neg' : ''}" data-label="Чистая прибыль">${cell(s.profit)}</span>
      <span class="c c-margin${(s.margin ?? 0) < 0 ? ' neg' : ''}" data-label="Маржа">${pct(s.margin)}</span>
      <span class="c c-in" data-label="Приток">${cell(s.inflow)}</span>
      <span class="c c-out" data-label="Отток">${cell(s.outflow)}</span>
    </button>`).join('');
  return `
  <div class="wrap wrap-wide page">
    <header class="top">
      <div class="top-main">
        <p class="brand">Деньги Monee · пульт</p>
        <h1>Месяцы</h1>
        <p class="sub">${esc(range)}</p>
      </div>
      ${nav('months')}
    </header>

    <section class="chart-block" aria-label="Выручка и чистая прибыль по месяцам">
      <div class="chart-legend"><span><i class="sw sw-rev"></i>Выручка</span><span><i class="sw sw-profit"></i>Чистая прибыль</span></div>
      <div id="chart" class="chart"></div>
    </section>

    <div class="months-layout">
      <section class="mtable" aria-label="Цифры по месяцам">
        <div class="mhead" aria-hidden="true">
          <span class="c c-month">Месяц<small>суммы в ₽</small></span><span class="c c-wb">WB</span><span class="c c-oz">Ozon</span><span class="c c-rev">Выручка</span>
          <span class="c c-profit">Чистая прибыль</span><span class="c c-margin">Маржа</span><span class="c c-in">Приток</span><span class="c c-out">Отток</span>
        </div>
        ${rows || '<p class="note">В таблице нет месяцев с цифрами.</p>'}
        <p class="fine">Выручка — без самовыкупов, из вкладки «PNL». Приток и отток — из «ДДС», без переводов между своими счетами.</p>
      </section>
      ${sel ? viewWhere(sel) : ''}
    </div>
  </div>`;
}

function viewWhere(s) {
  const max = s.expenses[0]?.v || 1;
  const list = s.expenses.length
    ? `<ol class="spend">${s.expenses.map((e) => `
        <li><span class="spend-name">${esc(cap(e.name))}</span><span class="spend-val">${rub(e.v)}</span>
        <span class="spend-bar" style="--w:${((e.v / max) * 100).toFixed(1)}%" aria-hidden="true"></span></li>`).join('')}</ol>`
    : '<p class="note">В «ДДС» нет расходов за этот месяц.</p>';
  return `
  <section class="where" id="where" aria-live="polite">
    <h2>${esc(mName(s.key))}</h2>
    <dl class="kv">
      <div><dt>Выручка WB</dt><dd>${rub(s.revWB)}</dd></div>
      <div><dt>Выручка Ozon</dt><dd>${rub(s.revOzon)}</dd></div>
      <div><dt>Выручка всего</dt><dd>${rub(s.revTotal)}</dd></div>
      <div><dt>Чистая прибыль</dt><dd class="${(s.profit ?? 0) < 0 ? 'neg' : ''}">${rub(s.profit)} <small>${pct(s.margin)}</small></dd></div>
      <div><dt>Пришло денег</dt><dd>${rub(s.inflow)}</dd></div>
      <div><dt>Ушло денег</dt><dd>${rub(s.outflow)}</dd></div>
    </dl>
    <h3>Куда ушли деньги</h3>
    ${list}
  </section>`;
}

function drawChart() {
  const box = document.getElementById('chart');
  if (!box || !state.model) return;
  const series = state.model.series;
  if (!series.length) { box.innerHTML = ''; return; }
  const W = Math.max(280, box.clientWidth);
  const H = W < 600 ? 190 : 240;
  const padL = W < 600 ? 40 : 56; const padB = 26; const padT = 10;
  const vals = series.flatMap((s) => [s.revTotal ?? 0, s.profit ?? 0]);
  const max = Math.max(1, ...vals); const min = Math.min(0, ...vals);
  const y = (v) => padT + (H - padT - padB) * (1 - (v - min) / (max - min));
  const step = (W - padL) / series.length;
  const bw = Math.max(4, Math.min(22, step * 0.32));
  const ticks = niceTicks(min, max, 4);
  const sel = state.selected ?? series.at(-1).key;
  let g = '';
  for (const t of ticks) g += `<line x1="${padL}" x2="${W}" y1="${y(t)}" y2="${y(t)}" class="grid${t === 0 ? ' zero' : ''}"/><text x="${padL - 6}" y="${y(t) + 4}" class="tick" text-anchor="end">${t === 0 ? '0' : shortRub(t)}</text>`;
  series.forEach((s, i) => {
    const cx = padL + step * i + step / 2;
    const rev = s.revTotal ?? 0; const pr = s.profit ?? 0;
    const isSel = s.key === sel;
    g += `<g class="grp${isSel ? ' is-sel' : ''}" data-month="${s.key}" tabindex="0" role="button" aria-label="${esc(mName(s.key))}: выручка ${rub(s.revTotal)}, чистая прибыль ${rub(s.profit)}">
      <rect class="hit" x="${padL + step * i + 1}" y="${padT}" width="${step - 2}" height="${H - padT}" rx="6"/>
      <rect class="b-rev" x="${cx - bw - 1}" y="${Math.min(y(rev), y(0))}" width="${bw}" height="${Math.max(1, Math.abs(y(0) - y(rev)))}" rx="2"/>
      <rect class="b-profit${pr < 0 ? ' neg' : ''}" x="${cx + 1}" y="${Math.min(y(pr), y(0))}" width="${bw}" height="${Math.max(1, Math.abs(y(0) - y(pr)))}" rx="2"/>
      <text x="${cx}" y="${H - 8}" class="mlab" text-anchor="middle">${M_SHORT[s.key % 12]}</text></g>`;
  });
  box.innerHTML = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${g}</svg>`;
}

function niceTicks(min, max, n) {
  const span = max - min || 1;
  const raw = span / n; const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((f) => f * mag).find((s) => s >= raw) || raw;
  const out = [];
  for (let v = Math.ceil(min / step) * step; v <= max + 1e-9; v += step) out.push(Math.round(v));
  return out;
}

function selectMonth(key) {
  state.selected = key;
  persist();
  render();
  if (window.matchMedia('(max-width: 899px)').matches) {
    document.getElementById('where')?.scrollIntoView({ behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth', block: 'nearest' });
  }
}

/* ---------- действия ---------- */

function startGoogle(okText) {
  // requestAccessToken должен вызываться прямо из нажатия, иначе браузер заблокирует окно входа
  run(() => loadFromGoogle(), okText);
}

app.addEventListener('click', (ev) => {
  const monthEl = ev.target.closest('[data-month]');
  if (monthEl) { selectMonth(+monthEl.dataset.month); return; }
  const btn = ev.target.closest('[data-act]');
  if (!btn || btn.disabled) return;
  const act = btn.dataset.act;
  if (act === 'market-retry') loadMarket();
  else if (act === 'google') startGoogle();
  else if (act === 'file') { fileInput.value = ''; fileInput.click(); }
  else if (act === 'demo') run(async () => demoModel());
  else if (act === 'months') go('months');
  else if (act === 'back') go('now');
  else if (act === 'refresh') {
    const type = state.model?.source.type;
    if (type === 'google') startGoogle('Обновлено');
    else if (type === 'file') { fileInput.value = ''; fileInput.click(); }
    else run(async () => demoModel(), 'Обновлено');
  } else if (act === 'logout') {
    if (google_.token && window.google?.accounts?.oauth2) window.google.accounts.oauth2.revoke(google_.token, () => {});
    google_.token = null;
    state.model = null; state.error = null; state.selected = null;
    go('login');
    toast('Цифры стёрты с этого устройства');
  }
});

app.addEventListener('keydown', (ev) => {
  const el = ev.target.closest?.('g[data-month]');
  if (el && (ev.key === 'Enter' || ev.key === ' ')) { ev.preventDefault(); selectMonth(+el.dataset.month); }
});

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (!file) return;
  const wasModel = Boolean(state.model);
  showLoading('Читаю таблицу…');
  run(async () => modelFromBuffer(await file.arrayBuffer(), { type: 'file', name: file.name }), wasModel ? 'Обновлено' : null);
});

let resizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => { if (state.screen === 'months') drawChart(); }, 120);
});

/* ---------- старт ---------- */

(function start() {
  const saved = load('local', LAST_KEY) || load('session', SS_KEY);
  save('session', SS_KEY, null);
  if (saved?.model?.v === 1) {
    state.model = saved.model;
    state.selected = saved.selected ?? saved.model.series.at(-1)?.key ?? null;
    state.screen = saved.screen === 'months' ? 'months' : 'now';
  }
  render();
  loadMarket();
  if (CONFIG.googleClientId) gisReady().catch(() => {});
}());
