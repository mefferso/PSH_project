import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { inflateRawSync } from 'node:zlib';

const FIXTURES = [
  'fixtures/PSHLIX_testingspreadsheet.xlsx',
  'fixtures/PSHLIX_2024AL06_Francine_Data.xlsx'
];
const TARGET_SHEET = 'Wind and Pressure';
const WEATHERSTEM_IDS = [
  'WSEBRAlexBox',
  'WSEBRTigerStadium',
  'WSNOLakefront',
  'WSNOMidCity',
  'WSNOMidCIty',
  'WSNOBayouSauvage',
  'WSSCEOC',
  'WSSCLuling'
];

const WEATHERSTEM_REFERENCE_NAMES = {
  WSEBRAlexBox: 'LSU Alex Box Stadium',
  WSEBRTigerStadium: 'LSU Tiger Stadium',
  WSNOLakefront: 'Municipal Yacht Harbor',
  WSNOMidCIty: 'Mid City New Orleasn',
  WSNOBayouSauvage: 'Bayou Sauvage',
  WSSCEOC: 'St. Charles Parish EOC',
  WSSCLuling: 'Luling'
};

const TESTING_FIXTURE = 'fixtures/PSHLIX_testingspreadsheet.xlsx';
const FRANCINE_FIXTURE = 'fixtures/PSHLIX_2024AL06_Francine_Data.xlsx';

function die(message) {
  console.error(`FAIL: ${message}`);
  process.exitCode = 1;
}

function u16(buf, off) { return buf.readUInt16LE(off); }
function u32(buf, off) { return buf.readUInt32LE(off); }

function unzipEntries(buf) {
  const EOCD = 0x06054b50;
  const CEN = 0x02014b50;
  const LOC = 0x04034b50;
  let eocd = -1;
  for (let i = Math.max(0, buf.length - 0x10000 - 22); i <= buf.length - 22; i++) {
    if (u32(buf, i) === EOCD) eocd = i;
  }
  if (eocd < 0) throw new Error('ZIP end-of-central-directory not found');
  const entries = new Map();
  const count = u16(buf, eocd + 10);
  let p = u32(buf, eocd + 16);
  for (let i = 0; i < count; i++) {
    if (u32(buf, p) !== CEN) throw new Error(`Bad central directory signature at ${p}`);
    const method = u16(buf, p + 10);
    const compressedSize = u32(buf, p + 20);
    const uncompressedSize = u32(buf, p + 24);
    const nameLen = u16(buf, p + 28);
    const extraLen = u16(buf, p + 30);
    const commentLen = u16(buf, p + 32);
    const localOffset = u32(buf, p + 42);
    const name = buf.subarray(p + 46, p + 46 + nameLen).toString('utf8');
    if (u32(buf, localOffset) !== LOC) throw new Error(`Bad local header for ${name}`);
    const localNameLen = u16(buf, localOffset + 26);
    const localExtraLen = u16(buf, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const compressed = buf.subarray(dataStart, dataStart + compressedSize);
    let data;
    if (method === 0) data = compressed;
    else if (method === 8) data = inflateRawSync(compressed);
    else throw new Error(`Unsupported ZIP method ${method} for ${name}`);
    if (data.length !== uncompressedSize) throw new Error(`Size mismatch for ${name}`);
    entries.set(name, data);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

function xmlText(entries, name, required = true) {
  const b = entries.get(name);
  if (!b) {
    if (required) throw new Error(`Missing XLSX part ${name}`);
    return '';
  }
  return b.toString('utf8');
}

function decodeXml(s) {
  return String(s ?? '')
    .replace(/&#(x?[0-9a-f]+);/gi, (_, n) => String.fromCodePoint(
      n[0].toLowerCase() === 'x' ? parseInt(n.slice(1), 16) : parseInt(n, 10)
    ))
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function attr(tag, name) {
  const m = tag.match(new RegExp(`\\b${name}="([^"]*)"`));
  return m ? decodeXml(m[1]) : null;
}

function parseSharedStrings(entries) {
  const xml = xmlText(entries, 'xl/sharedStrings.xml', false);
  if (!xml) return [];
  const out = [];
  for (const m of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    const pieces = [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x => decodeXml(x[1]));
    out.push(pieces.join(''));
  }
  return out;
}

function normalizeTarget(target) {
  const clean = target.replace(/^\//, '');
  return clean.startsWith('xl/') ? clean : `xl/${clean}`;
}

function parseSheets(entries) {
  const workbook = xmlText(entries, 'xl/workbook.xml');
  const rels = xmlText(entries, 'xl/_rels/workbook.xml.rels');
  const relMap = new Map();
  for (const m of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = attr(m[0], 'Id');
    const target = attr(m[0], 'Target');
    if (id && target) relMap.set(id, normalizeTarget(target));
  }
  const out = [];
  for (const m of workbook.matchAll(/<sheet\b[^>]*\/>/g)) {
    const name = attr(m[0], 'name');
    const rid = attr(m[0], 'r:id');
    if (name && rid && relMap.has(rid)) out.push({ name, path: relMap.get(rid) });
  }
  return out;
}

function cellValue(cellTag, body, shared) {
  const t = attr(cellTag, 't');
  if (t === 'inlineStr') {
    return [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x => decodeXml(x[1])).join('');
  }
  const vm = body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
  if (!vm) return '';
  const raw = decodeXml(vm[1]);
  if (t === 's') return shared[Number(raw)] ?? raw;
  if (t === 'b') return raw === '1';
  if (t === 'str' || t === 'e') return raw;
  const n = Number(raw);
  return Number.isFinite(n) ? n : raw;
}

function parseSheetRows(xml, shared) {
  const rows = new Map();
  for (const m of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const tag = `<c${m[1]}>`;
    const ref = attr(tag, 'r');
    if (!ref) continue;
    const rm = ref.match(/^([A-Z]+)(\d+)$/i);
    if (!rm) continue;
    const col = rm[1].toUpperCase();
    const rowNum = Number(rm[2]);
    if (!rows.has(rowNum)) rows.set(rowNum, {});
    rows.get(rowNum)[col] = cellValue(tag, m[2], shared);
  }
  return rows;
}

function workbookSummary(buf) {
  const entries = unzipEntries(buf);
  const shared = parseSharedStrings(entries);
  const sheets = parseSheets(entries);
  const sheet = sheets.find(s => s.name === TARGET_SHEET);
  if (!sheet) throw new Error(`Sheet ${TARGET_SHEET} not found; found: ${sheets.map(x => x.name).join(', ')}`);
  const rows = parseSheetRows(xmlText(entries, sheet.path), shared);
  const found = [];
  for (const [rowNum, row] of rows.entries()) {
    // Google Sheets XLSX exports can leave a linked Site ID cell without a cached
    // display value even though the row is otherwise complete. Do not silently
    // lose WeatherSTEM rows just because column A is uncached in the export.
    const id = String(row.A ?? '').trim();
    const network = String(row.H ?? '').trim().toUpperCase();
    const isWeatherStem = network === 'WEATHERSTEM'
      || id.toUpperCase().startsWith('WS')
      || WEATHERSTEM_IDS.some(x => x.toUpperCase() === id.toUpperCase());
    if (isWeatherStem) {
      found.push({ row: rowNum, id: id || null, network: row.H || null, cells: row });
    }
  }
  return { sheets: sheets.map(x => x.name), weatherstemRows: found };
}

const manifest = JSON.parse(await readFile('fixtures/source_manifest.json', 'utf8'));
const francineExpected = JSON.parse(await readFile('tests/francine_expected.json', 'utf8'));
const manifestByPath = new Map(manifest.files.map(x => [x.committed_path || `fixtures/${x.intended_repo_name}`, x]));
const summariesByPath = new Map();

function weatherStemRowByName(summary, siteName) {
  return summary.weatherstemRows.find(row => String(row.cells.B ?? '').trim() === siteName) || null;
}

function weatherStemSustained(row) {
  // The exported Francine/test workbooks place sustained wind in J when an
  // anemometer-height value exists in I, and in I when that height is blank.
  const j = Number(row?.cells?.J);
  if (Number.isFinite(j)) return j;
  const i = Number(row?.cells?.I);
  return Number.isFinite(i) ? i : null;
}

function numericCell(row, col) {
  const n = Number(row?.cells?.[col]);
  return Number.isFinite(n) ? n : null;
}

function assertNear(label, actual, expected, tolerance) {
  if (!Number.isFinite(actual) || Math.abs(actual - expected) > tolerance) {
    die(`${label}: expected ${expected} ± ${tolerance}, got ${actual}`);
  }
}

for (const path of FIXTURES) {
  const buf = await readFile(path);
  const hash = createHash('sha256').update(buf).digest('hex');
  const item = manifestByPath.get(path);
  console.log(`FIXTURE ${path}`);
  console.log(`  size_bytes=${buf.length}`);
  console.log(`  sha256=${hash}`);
  if (!item) die(`${path} is not represented in fixtures/source_manifest.json`);
  else {
    if (item.sha256 !== hash) die(`${path} sha256 mismatch: manifest=${item.sha256} actual=${hash}`);
    if (item.size_bytes !== buf.length) die(`${path} size mismatch: manifest=${item.size_bytes} actual=${buf.length}`);
  }
  const summary = workbookSummary(buf);
  summariesByPath.set(path, summary);
  console.log(`  sheets=${summary.sheets.join(' | ')}`);
  console.log(`  WeatherSTEM rows from "${TARGET_SHEET}":`);
  if (!summary.weatherstemRows.length) {
    die(`${path} contained no detectable WeatherSTEM rows in "${TARGET_SHEET}"`);
  }
  for (const row of summary.weatherstemRows) console.log(JSON.stringify(row));
}

const issued = summariesByPath.get(FRANCINE_FIXTURE);
const testing = summariesByPath.get(TESTING_FIXTURE);
if (!issued || !testing) {
  die('WeatherSTEM fixture comparison could not load both workbook summaries.');
} else {
  console.log('WeatherSTEM Francine authority checks:');
  for (const [id, siteName] of Object.entries(WEATHERSTEM_REFERENCE_NAMES)) {
    const expected = francineExpected.wind_pressure?.[id];
    if (!expected) {
      die(`Missing tests/francine_expected.json entry for ${id}`);
      continue;
    }

    const issuedRow = weatherStemRowByName(issued, siteName);
    const testingRow = weatherStemRowByName(testing, siteName);
    if (!issuedRow || !testingRow) {
      die(`${id}: could not find "${siteName}" in both XLSX fixtures`);
      continue;
    }

    const issuedSustained = weatherStemSustained(issuedRow);
    const issuedGust = numericCell(issuedRow, 'Q');
    const issuedPressure = numericCell(issuedRow, 'W');
    assertNear(`${id} issued sustained`, issuedSustained, expected.sustained_kt, 0.01);
    assertNear(`${id} issued gust`, issuedGust, expected.gust_kt, 0.01);
    assertNear(`${id} issued pressure`, issuedPressure, expected.mslp_mb, 0.05);

    const testingSustained = weatherStemSustained(testingRow);
    const testingGust = numericCell(testingRow, 'Q');
    assertNear(`${id} v0.10 sustained snapshot`, testingSustained, expected.sustained_kt, 1);

    if (id === 'WSEBRTigerStadium') {
      const gustError = testingGust - expected.gust_kt;
      if (!(Math.abs(gustError) > 1)) {
        die(`${id}: known Tiger gust discrepancy unexpectedly disappeared; review the source/fixture before changing expectations`);
      }
      console.log(`  ${id}: sustained PASS; known gust mismatch ${testingGust.toFixed(2)} vs ${expected.gust_kt.toFixed(2)} kt retained`);
    } else {
      assertNear(`${id} v0.10 gust snapshot`, testingGust, expected.gust_kt, 1);
      console.log(`  ${id}: sustained/gust snapshot PASS`);
    }
  }
}

if (issued && testing) {
  let matched = 0;
  let sustainedAbsError = 0;
  let sustainedMaxError = 0;
  let gustAbsError = 0;
  let gustCount = 0;
  for (const issuedRow of issued.weatherstemRows) {
    const siteName = String(issuedRow.cells.B ?? '').trim();
    if (!siteName) continue;
    const testingRow = weatherStemRowByName(testing, siteName);
    if (!testingRow) {
      die(`WeatherSTEM network snapshot: testing fixture is missing issued station "${siteName}"`);
      continue;
    }

    const issuedSustained = weatherStemSustained(issuedRow);
    const testingSustained = weatherStemSustained(testingRow);
    const sustainedError = testingSustained - issuedSustained;
    assertNear(`${siteName} network sustained`, testingSustained, issuedSustained, 1);
    sustainedAbsError += Math.abs(sustainedError);
    sustainedMaxError = Math.max(sustainedMaxError, Math.abs(sustainedError));

    const issuedGust = numericCell(issuedRow, 'Q');
    const testingGust = numericCell(testingRow, 'Q');
    if (Number.isFinite(issuedGust) && Number.isFinite(testingGust)) {
      const gustError = testingGust - issuedGust;
      gustAbsError += Math.abs(gustError);
      gustCount++;
      if (siteName !== 'LSU Tiger Stadium') {
        assertNear(`${siteName} network gust`, testingGust, issuedGust, 1);
      }
    }
    matched++;
  }

  if (matched < 25) {
    die(`WeatherSTEM network snapshot matched only ${matched} issued stations; expected broad Francine coverage`);
  } else {
    console.log(
      `WeatherSTEM network snapshot: ${matched} matched; sustained MAE=${(sustainedAbsError / matched).toFixed(3)} kt; ` +
      `max abs error=${sustainedMaxError.toFixed(3)} kt; gust MAE=${(gustAbsError / Math.max(1, gustCount)).toFixed(3)} kt`
    );
  }
}
if (!process.exitCode) console.log('PASS: committed XLSX fixtures match manifest, issued WeatherSTEM references, and validated v0.10 snapshots.');
