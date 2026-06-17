const fs = require('fs');
const path = require('path');
const os = require('os');
const zlib = require('zlib');
const { chromium } = require('playwright-core');

const REQUIRED_FIELDS = [
  'date',
  'platform',
  'account',
  'impressions',
  'clicks',
  'cost',
  'conversions',
];
const COPIED_TEMPLATE_FIELDS = new Set(['platform', 'account']);

const ZERO_AFTER_COPY_HEADERS = [];
const CLEAR_AFTER_COPY_HEADERS = [];
const BALANCE_ORDER_HEADERS = ['\u8ba2\u5355'];
const BALANCE_AMOUNT_HEADERS = ['\u4ea7\u51fa\u91d1\u989d'];
const REBATE_HEADERS = ['\u8fd4\u70b9'];
const FINANCE_COST_HEADERS = ['\u8d22\u52a1\u6d88\u8017'];
const PROFIT_HEADERS = ['\u5229\u6da6'];
const ROI_HEADERS = ['ROI'];
const DEFAULT_BALANCE_RECONCILE_FILE = path.join(
  os.homedir(),
  'Desktop',
  '\u4e00\u952e\u51fa\u7a3f\u8bb0\u5f55\u5bfc\u51fa\u5217\u8868.xlsx'
);

const BALANCE_FIELD_HEADERS = {
  date: ['\u521b\u5efa\u65f6\u95f4', '\u6d88\u8017\u65f6\u95f4', '\u652f\u4ed8\u65f6\u95f4', '\u65f6\u95f4', '\u65e5\u671f', 'date'],
  account: ['\u5b9d\u8d1d\u540d\u79f0', '\u8d26\u6237', '\u8d26\u53f7', 'account'],
  amount: ['\u5b9e\u9645\u652f\u4ed8\u91d1\u989d', '\u91d1\u989d', '\u652f\u4ed8\u91d1\u989d', 'amount'],
};

const FIELD_HEADERS = new Map([
  ['date', ['\u65f6\u95f4', '\u65e5\u671f', 'date']],
  ['platform', ['\u6e20\u9053', '\u6765\u6e90', 'platform']],
  ['account', ['\u8d26\u6237', '\u8d26\u53f7', 'account']],
  ['impressions', ['\u5c55\u73b0', 'impressions']],
  ['clicks', ['\u70b9\u51fb', 'clicks']],
  ['cost', ['\u8d26\u9762\u6d88\u8017', '\u6d88\u8017', 'cost']],
  ['conversions', ['\u670d\u52a1\u8d2d\u4e70\u6210\u529f', '\u8f6c\u5316', 'conversions']],
]);

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, ''));
}

function crc32(buffer) {
  const table = crc32.table || (crc32.table = Array.from({ length: 256 }, (_, index) => {
    let value = index;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? 0xEDB88320 ^ (value >>> 1) : value >>> 1;
    }
    return value >>> 0;
  }));
  let crc = 0 ^ -1;
  for (const byte of buffer) {
    crc = (crc >>> 8) ^ table[(crc ^ byte) & 0xFF];
  }
  return (crc ^ -1) >>> 0;
}

function parseZipEntries(filePath) {
  const buffer = fs.readFileSync(filePath);
  const entries = new Map();
  let offset = 0;
  while (offset < buffer.length - 46) {
    const signature = buffer.readUInt32LE(offset);
    if (signature !== 0x02014b50) {
      offset += 1;
      continue;
    }

    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const fileName = buffer.slice(offset + 46, offset + 46 + fileNameLength).toString('utf8');

    if (buffer.readUInt32LE(localHeaderOffset) !== 0x04034b50) {
      throw new Error('Invalid XLSX local header: ' + fileName);
    }
    const localFileNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localFileNameLength + localExtraLength;
    const dataEnd = dataStart + compressedSize;
    const data = buffer.slice(dataStart, dataEnd);
    if (method === 0) {
      entries.set(fileName, data);
    } else if (method === 8) {
      entries.set(fileName, zlib.inflateRawSync(data));
    } else {
      throw new Error('Unsupported XLSX compression method ' + method + ': ' + fileName);
    }
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function getXmlAttr(text, name) {
  const match = new RegExp(`${name}="([^"]*)"`).exec(text);
  return match ? match[1] : '';
}

function stripXmlTags(text) {
  return text.replace(/<[^>]+>/g, '');
}

function decodeXmlText(text) {
  return String(text || '')
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, decimal) => String.fromCodePoint(parseInt(decimal, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function readXlsxRows(filePath) {
  const entries = parseZipEntries(filePath);
  const sharedXml = entries.get('xl/sharedStrings.xml');
  const shared = [];
  if (sharedXml) {
    const sharedText = sharedXml.toString('utf8');
    for (const match of sharedText.matchAll(/<si\b[\s\S]*?<\/si>/g)) {
      const values = [...match[0].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
        .map((item) => decodeXmlText(item[1]));
      shared.push(values.join(''));
    }
  }

  const workbook = (entries.get('xl/workbook.xml') || Buffer.from('')).toString('utf8');
  const rels = (entries.get('xl/_rels/workbook.xml.rels') || Buffer.from('')).toString('utf8');
  const relMap = new Map();
  for (const match of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    relMap.set(getXmlAttr(match[0], 'Id'), getXmlAttr(match[0], 'Target'));
  }

  const sheetMatch = /<sheet\b[^>]*>/g.exec(workbook);
  if (!sheetMatch) return [];
  const rid = getXmlAttr(sheetMatch[0], 'r:id');
  let target = relMap.get(rid) || 'worksheets/sheet1.xml';
  if (!target.startsWith('xl/')) target = 'xl/' + target.replace(/^\/+/, '');
  const sheetXml = entries.get(target);
  if (!sheetXml) return [];

  const rows = [];
  const sheetText = sheetXml.toString('utf8');
  for (const rowMatch of sheetText.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
    const cells = [];
    for (const cellMatch of rowMatch[1].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const attrs = cellMatch[1];
      const body = cellMatch[2];
      const type = getXmlAttr(attrs, 't');
      let value = '';
      if (type === 'inlineStr') {
        value = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((item) => decodeXmlText(item[1])).join('');
      } else {
        const valueMatch = /<v>([\s\S]*?)<\/v>/.exec(body);
        value = valueMatch ? decodeXmlText(valueMatch[1]) : '';
        if (type === 's' && value !== '') value = shared[Number(value)] || '';
      }
      cells.push(value);
    }
    if (cells.some((cell) => normalize(cell))) rows.push(cells);
  }
  return rows;
}

function headerIndex(cells, names) {
  const wanted = new Set(names.map((name) => normalizeHeader(name)));
  return cells.findIndex((cell) => wanted.has(normalizeHeader(cell)));
}

function preferredHeaderIndex(cells, names) {
  for (const name of names) {
    const wanted = normalizeHeader(name);
    const index = cells.findIndex((cell) => normalizeHeader(cell) === wanted);
    if (index >= 0) return index;
  }
  return -1;
}

function parseNumber(value) {
  const text = normalize(value).replace(/,/g, '');
  if (!text) return 0;
  const match = /-?\d+(?:\.\d+)?/.exec(text);
  return match ? Number(match[0]) : 0;
}

function formatNumber(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return '0';
  return String(Math.trunc(number));
}

function findBalanceHeaderRow(rows) {
  for (let index = 0; index < rows.length; index++) {
    const cells = rows[index];
    if (
      preferredHeaderIndex(cells, BALANCE_FIELD_HEADERS.date) >= 0 &&
      preferredHeaderIndex(cells, BALANCE_FIELD_HEADERS.account) >= 0 &&
      preferredHeaderIndex(cells, BALANCE_FIELD_HEADERS.amount) >= 0
    ) {
      return index;
    }
  }
  return -1;
}

function loadBalanceReconcile(config, baseDir) {
  const options = config.balanceReconcile || {};
  const enabled = options.enabled !== false;
  if (!enabled) return { enabled: false, byKey: new Map(), filePath: '' };

  const configuredPath = normalize(options.filePath || '');
  const filePath = configuredPath
    ? (path.isAbsolute(configuredPath) ? configuredPath : path.resolve(baseDir, configuredPath))
    : DEFAULT_BALANCE_RECONCILE_FILE;

  if (!fs.existsSync(filePath)) {
    console.log(`Balance reconcile file not found, skip: ${filePath}`);
    return { enabled: true, byKey: new Map(), filePath };
  }

  const rows = readXlsxRows(filePath);
  const headerRow = findBalanceHeaderRow(rows);
  if (headerRow < 0) {
    console.log(`Balance reconcile headers not found, skip: ${filePath}`);
    return { enabled: true, byKey: new Map(), filePath };
  }

  const headers = rows[headerRow];
  const dateIndex = preferredHeaderIndex(headers, BALANCE_FIELD_HEADERS.date);
  const accountIndex = preferredHeaderIndex(headers, BALANCE_FIELD_HEADERS.account);
  const amountIndex = preferredHeaderIndex(headers, BALANCE_FIELD_HEADERS.amount);
  const byKey = new Map();

  for (const cells of rows.slice(headerRow + 1)) {
    const date = normalizeDate(cells[dateIndex] || '');
    const account = normalizeAccount(cells[accountIndex] || '');
    if (!date || !account) continue;
    const key = `${account}|${date}`;
    const current = byKey.get(key) || { orders: 0, amount: 0 };
    current.orders += 1;
    current.amount += parseNumber(cells[amountIndex] || '');
    byKey.set(key, current);
  }

  console.log(`Loaded balance reconcile rows: ${byKey.size} account-date groups from ${filePath}`);
  return { enabled: true, byKey, filePath };
}

function applyBalanceReconcile(rows, balance) {
  if (!balance || !balance.enabled || !balance.byKey || balance.byKey.size === 0) return rows;
  let matched = 0;
  const result = rows.map((row) => {
    const canonical = sourceToCanonical(row);
    const key = `${normalizeAccount(canonical.account)}|${normalizeDate(canonical.date)}`;
    const item = balance.byKey.get(key);
    if (!item) return canonical;
    matched += 1;
    return {
      ...canonical,
      balanceOrders: String(item.orders),
      balanceAmount: formatNumber(item.amount),
    };
  });
  console.log(`Matched balance reconcile groups to source rows: ${matched}/${rows.length}`);
  return result;
}

function getArg(name, defaultValue = '') {
  const prefix = `--${name}=`;
  const value = process.argv.find((arg) => arg.startsWith(prefix));
  return value ? value.slice(prefix.length) : defaultValue;
}

function normalize(value) {
  return String(value ?? '').trim();
}

function normalizeAccount(value) {
  return normalize(value).replace(/\s+/g, ' ');
}

function normalizeHeader(value) {
  return normalize(value).replace(/\s+/g, '').toLowerCase();
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ',') {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else if (ch !== '\r') {
      field += ch;
    }
  }
  row.push(field);
  rows.push(row);
  return rows.filter((line) => line.some((value) => normalize(value) !== ''));
}

function readCsvObjects(filePath) {
  if (!fs.existsSync(filePath)) return [];
  const rows = parseCsv(fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, ''));
  if (rows.length < 2) return [];
  const headers = rows[0].map((header) => normalize(header).toLowerCase());
  return rows.slice(1).map((line) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = line[index] || '';
    });
    return sourceToCanonical(item);
  });
}

function canonicalHeader(header) {
  const text = normalize(header).toLowerCase();
  for (const [field, names] of FIELD_HEADERS.entries()) {
    if (names.map((name) => name.toLowerCase()).includes(text)) return field;
  }
  return '';
}

function sourceToCanonical(row) {
  return {
    date: normalizeDate(row.date),
    platform: normalize(row.platform),
    account: normalizeAccount(row.account),
    impressions: normalize(row.impressions),
    clicks: normalize(row.clicks),
    cost: normalize(row.cost),
    conversions: normalize(row.conversions),
    balanceOrders: normalize(row.balanceOrders),
    balanceAmount: normalize(row.balanceAmount),
  };
}

function normalizeDate(value) {
  const text = normalize(value);
  if (!text) return '';

  const ymd = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/.exec(text);
  if (ymd) {
    return [
      ymd[1],
      ymd[2].padStart(2, '0'),
      ymd[3].padStart(2, '0'),
    ].join('-');
  }

  const cn = /^(\d{4})\u5e74(\d{1,2})\u6708(\d{1,2})\u65e5/.exec(text);
  if (cn) {
    return [
      cn[1],
      cn[2].padStart(2, '0'),
      cn[3].padStart(2, '0'),
    ].join('-');
  }

  return text;
}

function formatDateForSheet(value) {
  const date = normalizeDate(value);
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  return match ? `${match[1]}/${match[2]}/${match[3]}` : normalize(value);
}

function addDays(dateText, days) {
  const date = normalizeDate(dateText);
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  if (!match) return '';
  const d = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  d.setUTCDate(d.getUTCDate() + days);
  const year = d.getUTCFullYear();
  const month = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function yesterdayLocalDate() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function todayLocalDate() {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function buildDateRange(startDate, endDate) {
  const start = normalizeDate(startDate);
  const end = normalizeDate(endDate);
  if (!start || !end) return [];
  const startTime = dateTime(start);
  const endTime = dateTime(end);
  if (startTime === null || endTime === null) return [];
  if (startTime > endTime) {
    throw new Error(`Start date must be earlier than or equal to end date: ${start} > ${end}`);
  }

  const dates = [];
  for (let cursor = start; dateTime(cursor) <= endTime; cursor = addDays(cursor, 1)) {
    dates.push(cursor);
  }
  return dates;
}

function buildTargetDates() {
  const targetDate = normalizeDate(getArg('date', ''));
  if (targetDate) return [targetDate];

  const days = Number(getArg('days', '0') || 0);
  const explicitStart = normalizeDate(getArg('start-date', ''));
  const explicitEnd = normalizeDate(getArg('end-date', ''));
  if (days > 0) {
    const end = explicitEnd || todayLocalDate();
    const start = addDays(end, -1 * (days - 1));
    return buildDateRange(start, end);
  }

  if (explicitStart || explicitEnd) {
    const start = explicitStart || explicitEnd;
    const end = explicitEnd || explicitStart;
    return buildDateRange(start, end);
  }

  return [];
}

function filterRowsByDate(rows, dateText) {
  const target = normalizeDate(dateText);
  if (!target) return rows;
  return rows.filter((row) => normalizeDate(row.date) === target);
}

function rowKey(row) {
  return `${normalizeAccount(row.account)}|${normalizeDate(row.date)}`;
}

function sortSourceRows(rows) {
  return [...rows].sort((a, b) => {
    const accountCompare = normalizeAccount(a.account).localeCompare(normalizeAccount(b.account), 'zh-Hans-CN');
    if (accountCompare !== 0) return accountCompare;
    return normalizeDate(a.date).localeCompare(normalizeDate(b.date));
  });
}

function splitCopiedSheet(text) {
  const normalized = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n');
  while (lines.length > 0 && lines[lines.length - 1] === '') {
    lines.pop();
  }
  return lines;
}

function standardFromCells(headers, cells) {
  const standard = {};
  headers.forEach((header, index) => {
    const field = canonicalHeader(header);
    if (field) standard[field] = cells[index] || '';
  });
  if (!standard.date && cells.length >= 1) standard.date = cells[0] || '';
  if (!standard.account && cells.length >= 3) standard.account = cells[2] || '';
  return sourceToCanonical(standard);
}

function parseSheetText(text) {
  const lines = splitCopiedSheet(text);
  if (lines.length === 0) {
    const parsed = { headers: [], headerFields: new Map(), zeroColumns: [], clearColumns: [], rows: [], rawLineCount: 0 };
    rebuildIndexes(parsed);
    return parsed;
  }

  let headerLineIndex = 0;
  for (let i = 0; i < lines.length; i++) {
    const candidateHeaders = lines[i].split('\t').map((header) => normalize(header));
    const candidateFields = buildHeaderFields(candidateHeaders);
    if (REQUIRED_FIELDS.every((field) => candidateFields.has(field))) {
      headerLineIndex = i;
      break;
    }
  }

  const headers = lines[headerLineIndex].split('\t').map((header) => normalize(header));
  const headerFields = buildHeaderFields(headers);
  const rows = [];
  for (let i = headerLineIndex + 1; i < lines.length; i++) {
    const cells = lines[i].split('\t');
    rows.push({
      cells,
      sheetRow: i - headerLineIndex + 1,
      standard: standardFromCells(headers, cells),
    });
  }
  const parsed = {
    headers,
    headerFields,
    zeroColumns: buildColumnsByHeaders(headers, ZERO_AFTER_COPY_HEADERS),
    balanceOrderColumns: buildColumnsByHeaders(headers, BALANCE_ORDER_HEADERS),
    balanceAmountColumns: buildColumnsByHeaders(headers, BALANCE_AMOUNT_HEADERS),
    rebateColumns: buildColumnsByHeaders(headers, REBATE_HEADERS),
    financeCostColumns: buildColumnsByHeaders(headers, FINANCE_COST_HEADERS),
    profitColumns: buildColumnsByHeaders(headers, PROFIT_HEADERS),
    roiColumns: buildColumnsByHeaders(headers, ROI_HEADERS),
    clearColumns: buildColumnsByHeaders(headers, CLEAR_AFTER_COPY_HEADERS),
    rows,
    rawLineCount: lines.length,
  };
  rebuildIndexes(parsed);
  return parsed;
}

function buildHeaderFields(headers) {
  const result = new Map();
  headers.forEach((header, index) => {
    const field = canonicalHeader(header);
    if (field && !result.has(field)) result.set(field, index + 1);
  });
  return result;
}

function buildColumnsByHeaders(headers, wantedHeaders) {
  const wanted = new Set(wantedHeaders.map((header) => normalize(header)));
  const columns = [];
  headers.forEach((header, index) => {
    if (wanted.has(normalize(header))) columns.push(index + 1);
  });
  return columns;
}

function rebuildIndexes(parsed) {
  parsed.byAccountDate = new Map();
  parsed.byAccount = new Map();
  for (const row of parsed.rows) {
    const account = normalizeAccount(row.standard.account);
    const key = rowKey(row.standard);
    if (key.replace('|', '') && !parsed.byAccountDate.has(key)) {
      parsed.byAccountDate.set(key, row);
    }
    if (account) {
      if (!parsed.byAccount.has(account)) parsed.byAccount.set(account, []);
      parsed.byAccount.get(account).push(row);
    }
  }
}

function assertRequiredHeaders(parsed, sheetName) {
  const missing = REQUIRED_FIELDS.filter((field) => !parsed.headerFields.has(field));
  if (missing.length > 0) {
    throw new Error(`${sheetName}: missing headers for fields: ${missing.join(', ')}`);
  }
}

function findExisting(parsed, sourceRow) {
  return parsed.byAccountDate.get(rowKey(sourceRow)) || null;
}

function dateTime(value) {
  const date = normalizeDate(value);
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  if (!match) return null;
  return Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function findClosestAccountDate(parsed, sourceRow) {
  const accountRows = parsed.byAccount.get(normalizeAccount(sourceRow.account)) || [];
  const targetTime = dateTime(sourceRow.date);
  if (targetTime === null) return null;

  let best = null;
  for (const row of accountRows) {
    const rowTime = dateTime(row.standard.date);
    if (rowTime === null) continue;
    const distance = Math.abs(rowTime - targetTime);
    const isBeforeOrSame = rowTime <= targetTime;
    if (
      !best ||
      distance < best.distance ||
      (distance === best.distance && isBeforeOrSame && !best.isBeforeOrSame)
    ) {
      best = { row, distance, isBeforeOrSame };
    }
  }
  return best ? best.row : null;
}

function findInsertTemplateRow(parsed, sourceRow) {
  const accountRows = parsed.byAccount.get(normalizeAccount(sourceRow.account)) || [];
  const targetTime = dateTime(sourceRow.date);
  if (targetTime === null) return null;

  let previous = null;
  let next = null;
  for (const row of accountRows) {
    const rowTime = dateTime(row.standard.date);
    if (rowTime === null) continue;
    if (rowTime <= targetTime && (!previous || rowTime > previous.time)) {
      previous = { row, time: rowTime };
    }
    if (rowTime > targetTime && (!next || rowTime < next.time)) {
      next = { row, time: rowTime };
    }
  }

  if (previous) return previous.row;
  return next ? next.row : null;
}

function describeAccountRows(parsed, account, limit = 12) {
  const rows = parsed.byAccount.get(normalizeAccount(account)) || [];
  return rows
    .slice(-limit)
    .map((row) => `row ${row.sheetRow}: ${row.standard.account} ${row.standard.date}`)
    .join('; ');
}

function rememberInsertedRow(parsed, previousRow, sourceRow) {
  const insertAtSheetRow = previousRow.sheetRow + 1;
  const clonedCells = buildInsertedCells(parsed, previousRow, sourceRow, insertAtSheetRow);
  const inserted = {
    cells: clonedCells,
    sheetRow: insertAtSheetRow,
    standard: sourceToCanonical(sourceRow),
  };

  for (const row of parsed.rows) {
    if (row.sheetRow >= insertAtSheetRow) row.sheetRow += 1;
  }

  const previousIndex = parsed.rows.indexOf(previousRow);
  parsed.rows.splice(previousIndex + 1, 0, inserted);
  parsed.rawLineCount += 1;
  rebuildIndexes(parsed);
  return inserted;
}

function buildInsertedCells(parsed, previousRow, sourceRow, sheetRow) {
  const clonedCells = [...previousRow.cells];
  for (const field of REQUIRED_FIELDS) {
    const column = parsed.headerFields.get(field);
    if (!column) continue;
    clonedCells[column - 1] = formatInsertedFieldValue(field, sourceRow[field], sourceRow.date);
  }
  for (const column of parsed.zeroColumns || []) {
    clonedCells[column - 1] = '0';
  }
  for (const column of parsed.balanceOrderColumns || []) {
    clonedCells[column - 1] = sourceRow.balanceOrders || '0';
  }
  for (const column of parsed.balanceAmountColumns || []) {
    clonedCells[column - 1] = sourceRow.balanceAmount || '0';
  }
  for (const column of parsed.clearColumns || []) {
    clonedCells[column - 1] = '';
  }
  applyFinanceCostFormulas(parsed, clonedCells, sheetRow);
  applyProfitFormulas(parsed, clonedCells, sheetRow);
  applyRoiFormulas(parsed, clonedCells, sheetRow);
  return clonedCells;
}

function formatInsertedFieldValue(field, value, dateValue) {
  if (field === 'date') return formatDateForSheet(dateValue);
  if (['impressions', 'clicks', 'cost', 'conversions'].includes(field)) {
    return formatNumber(value);
  }
  return value || '';
}

function applyFinanceCostFormulas(parsed, cells, sheetRow) {
  const bookCostColumn = parsed.headerFields.get('cost');
  const rebateColumn = (parsed.rebateColumns || [])[0];
  if (!bookCostColumn || !rebateColumn || !sheetRow) return;

  const bookCostRef = columnName(bookCostColumn);
  const rebateRef = columnName(rebateColumn);
  for (const column of parsed.financeCostColumns || []) {
    cells[column - 1] = `=ROUNDDOWN(${bookCostRef}${sheetRow}/${rebateRef}${sheetRow},0)`;
  }
}

function applyProfitFormulas(parsed, cells, sheetRow) {
  const outputColumn = (parsed.balanceAmountColumns || [])[0];
  const financeCostColumn = (parsed.financeCostColumns || [])[0];
  if (!outputColumn || !financeCostColumn || !sheetRow) return;

  const outputRef = columnName(outputColumn);
  const financeCostRef = columnName(financeCostColumn);
  for (const column of parsed.profitColumns || []) {
    cells[column - 1] = `=ROUNDDOWN(${outputRef}${sheetRow}-${financeCostRef}${sheetRow},0)`;
  }
}

function applyRoiFormulas(parsed, cells, sheetRow) {
  const profitColumn = (parsed.profitColumns || [])[0];
  const outputColumn = (parsed.balanceAmountColumns || [])[0];
  if (!profitColumn || !outputColumn || !sheetRow) return;

  const profitRef = columnName(profitColumn);
  const outputRef = columnName(outputColumn);
  for (const column of parsed.roiColumns || []) {
    cells[column - 1] = `=IFERROR(${profitRef}${sheetRow}/${outputRef}${sheetRow},0)`;
  }
}

function buildSheetRows(config, allRows) {
  const accountSheets = buildAccountSheetMap(config);

  const seen = new Set();
  const sheetRows = new Map();
  for (const row of allRows) {
    const account = normalizeAccount(row.account);
    const sheetName = accountSheets.get(account);
    if (!sheetName) continue;
    const canonical = sourceToCanonical(row);
    const key = `${sheetName}|${rowKey(canonical)}`;
    if (seen.has(key)) continue;
    seen.add(key);
    if (!sheetRows.has(sheetName)) sheetRows.set(sheetName, []);
    sheetRows.get(sheetName).push(canonical);
  }

  for (const [sheetName, rows] of sheetRows.entries()) {
    sheetRows.set(sheetName, sortSourceRows(rows));
  }
  return sheetRows;
}

function buildAccountSheetMap(config) {
  const accountSheets = new Map();
  for (const [account, sheetName] of Object.entries(config.accountSheets || {})) {
    accountSheets.set(normalizeAccount(account), sheetName);
  }
  return accountSheets;
}

function buildSheetAccounts(config) {
  const sheetAccounts = new Map();
  for (const [account, sheetName] of Object.entries(config.accountSheets || {})) {
    const normalized = normalizeAccount(account);
    if (!normalized || !sheetName) continue;
    if (!sheetAccounts.has(sheetName)) sheetAccounts.set(sheetName, []);
    if (!sheetAccounts.get(sheetName).some((item) => normalizeAccount(item) === normalized)) {
      sheetAccounts.get(sheetName).push(account);
    }
  }
  return sheetAccounts;
}

function existingDatesForAccount(parsed, account) {
  const rows = parsed.byAccount.get(normalizeAccount(account)) || [];
  return new Set(rows.map((row) => normalizeDate(row.standard.date)).filter(Boolean));
}

function latestDate(dates) {
  return [...dates].sort().slice(-1)[0] || '';
}

function earliestDate(dates) {
  return [...dates].sort()[0] || '';
}

function buildDefaultMissingTargetDates(existing) {
  const start = earliestDate(existing);
  const end = yesterdayLocalDate();
  if (!start || dateTime(start) === null || dateTime(end) === null || dateTime(start) > dateTime(end)) {
    return [];
  }
  return buildDateRange(start, end);
}

async function scanMissingDates(page, sync, targetDates) {
  const sheetAccounts = buildSheetAccounts(sync);
  const plan = {
    documentUrl: sync.documentUrl || '',
    targetDates,
    generatedAt: new Date().toISOString(),
    items: [],
  };

  for (const [sheetName, accounts] of sheetAccounts.entries()) {
    await clickSheet(page, sheetName);
    const parsed = await reparseCurrentSheet(page, sync);
    if (parsed.headers.length > 0) {
      assertRequiredHeaders(parsed, sheetName);
    }

    for (const account of accounts) {
      const normalizedAccount = normalizeAccount(account);
      const accountRows = parsed.byAccount.get(normalizedAccount) || [];
      if (accountRows.length === 0) {
        plan.items.push({
          sheetName,
          account: normalizedAccount,
          matched: false,
          existingCount: 0,
          latestDate: '',
          missingDates: [],
          reason: 'account_not_found_in_column_c',
        });
        console.log(`${sheetName}: account not found in column C: ${normalizedAccount}`);
        continue;
      }

      const existing = existingDatesForAccount(parsed, normalizedAccount);
      const accountTargetDates = targetDates.length > 0 ? targetDates : buildDefaultMissingTargetDates(existing);
      const missingDates = accountTargetDates.filter((date) => !existing.has(normalizeDate(date)));
      plan.items.push({
        sheetName,
        account: normalizedAccount,
        matched: true,
        existingCount: accountRows.length,
        targetDates: accountTargetDates,
        latestDate: latestDate(existing),
        missingDates,
        reason: missingDates.length > 0 ? 'missing' : 'complete',
      });

      if (missingDates.length > 0) {
        console.log(`${sheetName}: missing ${normalizedAccount}: ${missingDates.join(', ')}`);
      } else {
        const dateScope = accountTargetDates.length > 0 ? accountTargetDates.join(', ') : 'no target dates';
        console.log(`${sheetName}: complete ${normalizedAccount} for ${dateScope}`);
      }
    }
  }

  return plan;
}

function defaultHeaders() {
  return [
    '\u65f6\u95f4',
    '\u6e20\u9053',
    '\u8d26\u6237',
    '\u5c55\u73b0',
    '\u70b9\u51fb',
    '\u8d26\u9762\u6d88\u8017',
    '\u670d\u52a1\u8d2d\u4e70\u6210\u529f',
  ];
}

function waitForEnter(message) {
  return new Promise((resolve) => {
    process.stdout.write(`${message}\n`);
    process.stdin.resume();
    process.stdin.once('data', () => {
      process.stdin.pause();
      resolve();
    });
  });
}

async function readClipboard(page) {
  try {
    return await page.evaluate(() => navigator.clipboard.readText());
  } catch {
    return '';
  }
}

async function writeClipboard(page, text) {
  await page.evaluate(async (value) => {
    await navigator.clipboard.writeText(value);
  }, String(text ?? ''));
}

async function clickSheet(page, sheetName) {
  const timeoutMs = 30000;
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const exact = page.getByText(sheetName, { exact: true });
    if (await exact.count()) {
      await exact.last().click({ timeout: 10000 });
      await page.waitForTimeout(1200);
      return;
    }

    const fuzzy = page.getByText(sheetName);
    if (await fuzzy.count()) {
      await fuzzy.last().click({ timeout: 10000 });
      await page.waitForTimeout(1200);
      return;
    }
    await page.waitForTimeout(1000);
  }

  await saveShimoDebug(page, `missing-sheet-${sheetName}`);
  throw new Error(`Cannot find Shimo sheet tab: ${sheetName}`);
}

async function openShimoDocument(page, documentUrl) {
  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      await page.goto(documentUrl, { waitUntil: 'domcontentloaded', timeout: 120000 });
      return;
    } catch (error) {
      lastError = error;
      const currentUrl = page.url();
      if (currentUrl && currentUrl.startsWith(documentUrl)) {
        await page.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => {});
        return;
      }
      console.log(`Open Shimo document failed, retry ${attempt}/3: ${error.message}`);
      await page.waitForTimeout(3000);
    }
  }
  throw lastError;
}

async function saveShimoDebug(page, label) {
  try {
    const safeLabel = String(label || 'debug').replace(/[^\w.-]+/g, '_');
    const debugDir = path.resolve(__dirname, '..', 'data');
    fs.mkdirSync(debugDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[-:T.Z]/g, '').slice(0, 14);
    const base = path.join(debugDir, `shimo_${safeLabel}_${stamp}`);
    await page.screenshot({ path: `${base}.png`, fullPage: true });
    const text = await page.locator('body').innerText({ timeout: 5000 }).catch(() => '');
    fs.writeFileSync(`${base}.txt`, text, 'utf8');
    console.log(`Saved Shimo debug files: ${base}.png, ${base}.txt`);
  } catch (error) {
    console.log(`Failed to save Shimo debug files: ${error.message}`);
  }
}

async function focusGrid(page, options) {
  await page.mouse.click(Number(options.clickX || 280), Number(options.clickY || 260));
  await page.waitForTimeout(300);
}

async function jumpToCell(page, row, column) {
  const address = `${columnName(column)}${row}`;
  const quickJump = page.locator('input.sm-quick-jump-input');
  if (await quickJump.count()) {
    await quickJump.first().click({ timeout: 5000 });
    await page.keyboard.press('Control+A');
    await page.keyboard.type(address);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(500);
    return true;
  }
  return false;
}

function columnName(column) {
  let value = Number(column);
  let name = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name || 'A';
}

async function copySheetText(page, options) {
  await focusGrid(page, options);
  await page.keyboard.press('Control+A');
  await page.waitForTimeout(250);
  await page.keyboard.press('Control+C');
  await page.waitForTimeout(Number(options.waitAfterCopySeconds || 1) * 1000);
  return readClipboard(page);
}

async function pressRepeated(page, key, count, delayMs = 20) {
  for (let i = 0; i < count; i++) {
    await page.keyboard.press(key);
    if (delayMs > 0) await page.waitForTimeout(delayMs);
  }
}

async function moveToCell(page, options, cursor, row, column) {
  if (await jumpToCell(page, row, column)) {
    cursor.row = row;
    cursor.column = column;
    return;
  }

  if (!cursor.row || !cursor.column) {
    await focusGrid(page, options);
    await page.keyboard.press('Control+Home');
    await page.waitForTimeout(120);
    cursor.row = 1;
    cursor.column = 1;
  }

  const rowDelta = row - cursor.row;
  const keyPressDelayMs = Number(options.keyPressDelayMs || 25);
  if (rowDelta > 0) await pressRepeated(page, 'ArrowDown', rowDelta, keyPressDelayMs);
  if (rowDelta < 0) await pressRepeated(page, 'ArrowUp', -rowDelta, keyPressDelayMs);

  const columnDelta = column - cursor.column;
  if (columnDelta > 0) await pressRepeated(page, 'ArrowRight', columnDelta, keyPressDelayMs);
  if (columnDelta < 0) await pressRepeated(page, 'ArrowLeft', -columnDelta, keyPressDelayMs);

  cursor.row = row;
  cursor.column = column;
  await page.waitForTimeout(120);
}

async function selectWholeRow(page, options, cursor, row) {
  await moveToCell(page, options, cursor, row, 1);
  await page.keyboard.press('Shift+Space');
  await page.waitForTimeout(250);
}

async function copySelectedRowText(page, options) {
  await page.keyboard.press('Control+C');
  await page.waitForTimeout(Number(options.waitAfterCopySeconds || 1) * 1000);
  return readClipboard(page);
}

function copiedRowLooksLikeTemplate(text, previousRow) {
  const rowText = String(text || '');
  const account = normalizeAccount(previousRow.standard.account);
  const date = normalizeDate(previousRow.standard.date);
  const cells = rowText.split(/\t|\r?\n/).map((cell) => normalize(cell));
  const hasAccount = cells.some((cell) => normalizeAccount(cell) === account) || rowText.includes(account);
  const hasDate = cells.some((cell) => normalizeDate(cell) === date);
  return hasAccount && hasDate;
}

function previewCopiedRow(text) {
  return String(text || '')
    .replace(/\r?\n/g, ' | ')
    .split('\t')
    .slice(0, 12)
    .join('\t')
    .slice(0, 500);
}

async function duplicateRowBelow(page, options, cursor, previousRow) {
  const previousSheetRow = previousRow.sheetRow;
  const candidateRows = [previousSheetRow];
  for (let offset = 1; offset <= Number(options.templateRowSearchWindow || 80); offset++) {
    candidateRows.push(previousSheetRow + offset);
    if (previousSheetRow - offset > 1) candidateRows.push(previousSheetRow - offset);
  }

  let copiedText = '';
  let actualTemplateRow = null;
  const previews = [];
  for (const candidateRow of candidateRows) {
    await selectWholeRow(page, options, cursor, candidateRow);
    copiedText = await copySelectedRowText(page, options);
    if (copiedRowLooksLikeTemplate(copiedText, previousRow)) {
      actualTemplateRow = candidateRow;
      break;
    }
    previews.push(`row ${candidateRow}: ${previewCopiedRow(copiedText)}`);
  }

  if (!actualTemplateRow) {
    throw new Error(`Copied row does not match template row ${previousSheetRow}: ${previews.slice(0, 6).join(' || ')}`);
  }
  if (actualTemplateRow !== previousSheetRow) {
    console.log(`Adjusted template row from parsed row ${previousSheetRow} to sheet row ${actualTemplateRow}`);
  }

  const targetSheetRow = actualTemplateRow + 1;
  await selectWholeRow(page, options, cursor, targetSheetRow);
  await page.keyboard.press('Control+Shift+=');
  await page.waitForTimeout(Number(options.waitAfterInsertSeconds || 2) * 1000);

  await selectWholeRow(page, options, cursor, targetSheetRow);
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(Number(options.waitAfterPasteSeconds || 4) * 1000);
  cursor.row = null;
  cursor.column = null;
  return targetSheetRow;
}

async function pasteFullRow(page, options, cursor, row, cells) {
  console.log(`Paste full row at sheet row ${row}`);
  await moveToCell(page, options, cursor, row, 1);
  await writeClipboard(page, cells.map((cell) => String(cell ?? '')).join('\t'));
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(Number(options.waitAfterPasteSeconds || 4) * 1000);
}

async function waitForInsertedRow(page, options, sourceRow, timeoutMs = 15000) {
  const started = Date.now();
  let parsed = null;
  while (Date.now() - started < timeoutMs) {
    parsed = await reparseCurrentSheet(page, options);
    const existing = findExisting(parsed, sourceRow);
    if (existing) return { parsed, row: existing };
    await page.waitForTimeout(1000);
  }
  return { parsed, row: null };
}

async function pasteValueAtCell(page, options, cursor, row, column, value) {
  await moveToCell(page, options, cursor, row, column);
  if (value === null) {
    await page.keyboard.press('Delete');
  } else {
    await typeCellValue(page, value);
  }
  await page.waitForTimeout(Number(options.waitAfterCellPasteMs || 250));
}

async function typeCellValue(page, value) {
  const text = String(value ?? '');
  await page.keyboard.press('Delete');
  if (text) {
    await page.keyboard.type(text, { delay: 5 });
  }
  await page.keyboard.press('Enter');
}

async function moveToColumnOnCurrentRow(page, cursor, column) {
  if (!cursor.column) {
    await page.keyboard.press('Home');
    await page.waitForTimeout(120);
    cursor.column = 1;
  }
  const columnDelta = column - cursor.column;
  if (columnDelta > 0) await pressRepeated(page, 'ArrowRight', columnDelta, 25);
  if (columnDelta < 0) await pressRepeated(page, 'ArrowLeft', -columnDelta, 25);
  cursor.column = column;
  await page.waitForTimeout(80);
}

async function pasteValueAtCurrentRow(page, options, cursor, column, value) {
  await moveToColumnOnCurrentRow(page, cursor, column);
  await typeCellValue(page, value);
  await page.waitForTimeout(Number(options.waitAfterCellPasteMs || 250));
}

async function updateMappedFields(page, options, parsed, cursor, row, sourceRow) {
  const actions = buildMappedFieldActions(parsed, sourceRow);
  await page.keyboard.press('Escape');
  await page.waitForTimeout(120);
  for (const action of actions) {
    await pasteValueAtCell(page, options, cursor, row, action.column, action.value);
  }
}

function buildMappedFieldActions(parsed, sourceRow) {
  const actions = [];
  for (const field of REQUIRED_FIELDS) {
    if (COPIED_TEMPLATE_FIELDS.has(field)) continue;
    const column = parsed.headerFields.get(field);
    if (!column) continue;
    const value = field === 'date' ? formatDateForSheet(sourceRow.date) : sourceRow[field] || '';
    actions.push({ column, value });
  }
  for (const column of parsed.zeroColumns || []) {
    actions.push({ column, value: '0' });
  }
  for (const column of parsed.balanceOrderColumns || []) {
    actions.push({ column, value: sourceRow.balanceOrders || '0' });
  }
  for (const column of parsed.balanceAmountColumns || []) {
    actions.push({ column, value: sourceRow.balanceAmount || '0' });
  }
  for (const column of parsed.clearColumns || []) {
    actions.push({ column, value: null });
  }
  actions.sort((a, b) => a.column - b.column);
  return actions;
}

function hasBalanceReconcileValues(sourceRow) {
  return normalize(sourceRow.balanceOrders) !== '' || normalize(sourceRow.balanceAmount) !== '';
}

function buildBalanceFieldActions(parsed, sourceRow) {
  if (!hasBalanceReconcileValues(sourceRow)) return [];
  const actions = [];
  for (const column of parsed.balanceOrderColumns || []) {
    actions.push({ column, value: sourceRow.balanceOrders || '0' });
  }
  for (const column of parsed.balanceAmountColumns || []) {
    actions.push({ column, value: sourceRow.balanceAmount || '0' });
  }
  actions.sort((a, b) => a.column - b.column);
  return actions;
}

async function updateBalanceFields(page, options, parsed, cursor, row, sourceRow) {
  const actions = buildBalanceFieldActions(parsed, sourceRow);
  const formulaActions = buildFormulaFieldActions(parsed, row);
  if (actions.length === 0 && formulaActions.length === 0) return false;
  await page.keyboard.press('Escape');
  await page.waitForTimeout(120);
  for (const action of actions) {
    await pasteValueAtCell(page, options, cursor, row, action.column, action.value);
  }
  for (const action of formulaActions) {
    await pasteValueAtCell(page, options, cursor, row, action.column, action.value);
  }
  return true;
}

function buildFormulaFieldActions(parsed, row) {
  const cells = [];
  const financeCostFormula = buildFormulaCellValue(parsed, row, 'financeCost');
  for (const column of parsed.financeCostColumns || []) {
    if (financeCostFormula) cells.push({ column, value: financeCostFormula });
  }

  const profitFormula = buildFormulaCellValue(parsed, row, 'profit');
  for (const column of parsed.profitColumns || []) {
    if (profitFormula) cells.push({ column, value: profitFormula });
  }

  const roiFormula = buildFormulaCellValue(parsed, row, 'roi');
  for (const column of parsed.roiColumns || []) {
    if (roiFormula) cells.push({ column, value: roiFormula });
  }

  return cells.sort((a, b) => a.column - b.column);
}

function buildFormulaCellValue(parsed, row, kind) {
  const bookCostColumn = parsed.headerFields.get('cost');
  const rebateColumn = (parsed.rebateColumns || [])[0];
  const outputColumn = (parsed.balanceAmountColumns || [])[0];
  const financeCostColumn = (parsed.financeCostColumns || [])[0];
  const profitColumn = (parsed.profitColumns || [])[0];

  if (kind === 'financeCost') {
    if (!bookCostColumn || !rebateColumn) return '';
    return `=ROUNDDOWN(${columnName(bookCostColumn)}${row}/${columnName(rebateColumn)}${row},0)`;
  }
  if (kind === 'profit') {
    if (!outputColumn || !financeCostColumn) return '';
    return `=ROUNDDOWN(${columnName(outputColumn)}${row}-${columnName(financeCostColumn)}${row},0)`;
  }
  if (kind === 'roi') {
    if (!profitColumn || !outputColumn) return '';
    return `=IFERROR(${columnName(profitColumn)}${row}/${columnName(outputColumn)}${row},0)`;
  }
  return '';
}

async function updateMappedFieldsAtCurrentRow(page, options, parsed, sourceRow) {
  const cursor = { column: null };
  const actions = buildMappedFieldActions(parsed, sourceRow);
  await page.keyboard.press('Escape');
  await page.waitForTimeout(120);
  for (const action of actions) {
    await pasteValueAtCurrentRow(page, options, cursor, action.column, action.value);
  }
}

async function reparseCurrentSheet(page, options) {
  const text = await copySheetText(page, options);
  return parseSheetText(text);
}

function summarizePlan(sheetName, parsed, rows) {
  let exists = 0;
  let missingWithTemplate = 0;
  let missingNoTemplate = 0;
  for (const row of rows) {
    if (findExisting(parsed, row)) {
      exists += 1;
      console.log(`${sheetName}: dry-run exists ${row.account} ${row.date}`);
    } else if (findInsertTemplateRow(parsed, row)) {
      missingWithTemplate += 1;
      const template = findInsertTemplateRow(parsed, row);
      console.log(`${sheetName}: dry-run insertable ${row.account} ${row.date} after row ${template.sheetRow} (${template.standard.date})`);
    } else {
      missingNoTemplate += 1;
      console.log(`${sheetName}: dry-run no account template row ${row.account} ${row.date}`);
    }
  }
  console.log(`${sheetName}: source=${rows.length}, exists=${exists}, insertable=${missingWithTemplate}, no_account_template=${missingNoTemplate}`);
}

function planInsertRows(parsed, rows) {
  return rows.map((sourceRow, index) => {
    const existing = findExisting(parsed, sourceRow);
    const templateRow = existing ? null : findInsertTemplateRow(parsed, sourceRow);
    return { sourceRow, existing, templateRow, index };
  }).sort((a, b) => {
    const rowA = a.templateRow ? a.templateRow.sheetRow : -1;
    const rowB = b.templateRow ? b.templateRow.sheetRow : -1;
    if (rowA !== rowB) return rowB - rowA;
    const dateCompare = normalizeDate(b.sourceRow.date).localeCompare(normalizeDate(a.sourceRow.date));
    if (dateCompare !== 0) return dateCompare;
    return a.index - b.index;
  });
}

async function syncSheet(page, sync, sheetName, rows, dryRun) {
  await clickSheet(page, sheetName);
  let parsed = await reparseCurrentSheet(page, sync);
  if (parsed.headers.length === 0) {
    parsed = parseSheetText(defaultHeaders().join('\t'));
  }
  assertRequiredHeaders(parsed, sheetName);

  if (dryRun) {
    summarizePlan(sheetName, parsed, rows);
    return;
  }

  let inserted = 0;
  let skipped = 0;
  let updated = 0;
  let noTemplate = 0;
  const insertedRows = [];
  const cursor = { row: null, column: null };

  for (const plan of planInsertRows(parsed, rows)) {
    const sourceRow = plan.sourceRow;
    if (plan.existing) {
      skipped += 1;
      if (await updateBalanceFields(page, sync, parsed, cursor, plan.existing.sheetRow, sourceRow)) {
        updated += 1;
        console.log(`${sheetName}: update balance ${sourceRow.account} ${sourceRow.date} orders=${sourceRow.balanceOrders || '0'} amount=${sourceRow.balanceAmount || '0'}`);
      } else {
        console.log(`${sheetName}: skip existing ${sourceRow.account} ${sourceRow.date}`);
      }
      continue;
    }

    const templateRow = plan.templateRow;
    if (!templateRow) {
      noTemplate += 1;
      console.log(`${sheetName}: no account template row, skip ${sourceRow.account} ${sourceRow.date}`);
      console.log(`${sheetName}: account rows ${sourceRow.account}: ${describeAccountRows(parsed, sourceRow.account) || 'none'}`);
      continue;
    }

    console.log(`${sheetName}: insert ${sourceRow.account} ${sourceRow.date} using template row ${templateRow.sheetRow} (${templateRow.standard.account} ${templateRow.standard.date})`);
    const targetRow = await duplicateRowBelow(page, sync, cursor, templateRow);
    await pasteFullRow(page, sync, cursor, targetRow, buildInsertedCells(parsed, templateRow, sourceRow, targetRow));
    const verifiedInsert = await waitForInsertedRow(page, sync, sourceRow, Number(sync.insertVerifyTimeoutMs || 15000));
    if (!verifiedInsert.row) {
      console.log(`${sheetName}: verify account rows ${sourceRow.account}: ${describeAccountRows(verifiedInsert.parsed || parsed, sourceRow.account)}`);
      throw new Error(`${sheetName}: inserted row not found after update: ${sourceRow.account} ${sourceRow.date}`);
    }
    parsed = verifiedInsert.parsed;
    insertedRows.push(sourceRow);
    inserted += 1;
  }

  if (insertedRows.length > 0 || updated > 0) {
    await page.waitForTimeout(Number(sync.waitAfterPasteSeconds || 4) * 1000);
  }

  if (insertedRows.length > 0) {
    const verified = await reparseCurrentSheet(page, sync);
    const missing = insertedRows.filter((row) => !findExisting(verified, row));
    if (missing.length > 0) {
      for (const row of missing) {
        console.log(`${sheetName}: verify account rows ${row.account}: ${describeAccountRows(verified, row.account)}`);
      }
      throw new Error(`${sheetName}: inserted rows not found after verification: ${missing.map((row) => `${row.account} ${row.date}`).join('; ')}`);
    }
  }

  console.log(`${sheetName}: done inserted=${inserted}, updated=${updated}, skipped=${skipped}, no_account_template=${noTemplate}`);
}

async function main() {
  const configPath = getArg('config');
  if (!configPath) throw new Error('Missing --config');

  const rootConfig = readJson(configPath);
  const sync = rootConfig.shimoAccountSheets || {};
  if (!sync.enabled) {
    console.log('Shimo account-sheet sync is disabled.');
    return;
  }
  if (!sync.documentUrl) throw new Error('Missing shimoAccountSheets.documentUrl');

  const dryRun = process.argv.includes('--dry-run');
  const setup = process.argv.includes('--setup');
  const scanMissing = process.argv.includes('--scan-missing');
  const targetDate = normalizeDate(getArg('date', ''));
  const targetAccount = normalizeAccount(getArg('account', ''));
  const baseDir = path.dirname(configPath);
  const allRows = [];
  let sheetRows = new Map();
  if (!scanMissing) {
    const balance = loadBalanceReconcile(sync, baseDir);
    for (const source of sync.sources || []) {
      const csvPath = path.isAbsolute(source.csvPath)
        ? source.csvPath
        : path.resolve(baseDir, source.csvPath);
      const rows = filterRowsByDate(readCsvObjects(csvPath), targetDate)
        .filter((row) => !targetAccount || normalizeAccount(row.account) === targetAccount);
      allRows.push(...rows);
    }
    sheetRows = buildSheetRows(sync, applyBalanceReconcile(allRows, balance));
  }
  if (targetDate) {
    console.log(`Target date: ${targetDate}`);
  }
  if (targetAccount) {
    console.log(`Target account: ${targetAccount}`);
  }

  const chromePath = sync.chromePath || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
  const userDataDir = sync.userDataDir || path.join(baseDir, 'data', 'chrome-shimo-profile');
  fs.mkdirSync(userDataDir, { recursive: true });

  const context = await chromium.launchPersistentContext(userDataDir, {
    executablePath: chromePath,
    headless: false,
    viewport: null,
    args: ['--start-maximized'],
    permissions: ['clipboard-read', 'clipboard-write'],
  });
  await context.grantPermissions(['clipboard-read', 'clipboard-write'], {
    origin: new URL(sync.documentUrl).origin,
  });

  const page = context.pages()[0] || (await context.newPage());
  await openShimoDocument(page, sync.documentUrl);

  if (setup) {
    await waitForEnter(
      'Please log in to Shimo in the opened Chrome window. Open the spreadsheet, then press Enter here to sync.'
    );
  } else {
    await page.waitForTimeout(Number(sync.waitAfterOpenSeconds || 8) * 1000);
  }

  try {
    if (scanMissing) {
      const targetDates = buildTargetDates();
      const plan = await scanMissingDates(page, sync, targetDates);
      const outputJson = getArg('output-json', '');
      if (outputJson) {
        fs.mkdirSync(path.dirname(outputJson), { recursive: true });
        fs.writeFileSync(outputJson, JSON.stringify(plan, null, 2), 'utf8');
        console.log(`Wrote missing plan: ${outputJson}`);
      } else {
        console.log(JSON.stringify(plan, null, 2));
      }
    } else {
      for (const [sheetName, rows] of sheetRows.entries()) {
        await syncSheet(page, sync, sheetName, rows, dryRun);
      }
    }
  } finally {
    if (!setup) {
      await context.close().catch(() => {});
    }
  }
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
