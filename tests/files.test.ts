import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { newBook, validateBook } from '../src/model';
import { exportXlsx, importXlsx, importCsv, exportCsv } from '../src/files';

test('XLSX round trip preserves values, formulas, style, dimensions, merge and multiple sheets', async () => {
  const source = newBook(true), id = source.sheetOrder[0];
  source.sheets[id].mergeData = [{ startRow: 8, endRow: 8, startColumn: 0, endColumn: 2 }];
  source.sheets[id].rowData = { 0: { h: 44 } };
  source.sheets.second = { id: 'second', name: '두 번째', rowCount: 200, columnCount: 26, cellData: { 0: { 0: { v: '한글, "쉼표"와 줄바꿈\n다음 줄' }, 1: { v: '00123', t: 1 }, 2: { f: "='가계부'!D2", v: 3000000 } } } }; source.sheetOrder.push('second');
  const bytes = await exportXlsx(source); const loaded = await importXlsx(bytes, '왕복.xlsx');
  assert.equal(loaded.sheetOrder.length, 2); const sheet = loaded.sheets[loaded.sheetOrder[0]];
  assert.equal(sheet.cellData![1][3].v, 3000000); assert.equal(sheet.cellData![5][3].f, '=SUM(D2:D5)');
  assert.equal((sheet.cellData![0][0].s as { bl: number }).bl, 1);
  assert.equal((sheet.cellData![0][0].s as { bg: { rgb: string } }).bg.rgb.toLowerCase(), '#18794e');
  assert.equal(sheet.columnData![2].w, 180); assert.equal(sheet.rowData![0].h, 44); assert.deepEqual(sheet.mergeData, source.sheets[id].mergeData);
  const second = loaded.sheets[loaded.sheetOrder[1]]; assert.equal(second.cellData![0][1].v, '00123'); assert.equal(second.cellData![0][2].f, "='가계부'!D2");
  const independent = new ExcelJS.Workbook(); await independent.xlsx.load(bytes); assert.equal(independent.worksheets[0].getCell('D6').formula, 'SUM(D2:D5)');
});
test('CSV quotes and Korean multiline roundtrip; formula injection is escaped', () => {
  const book = importCsv('이름,설명,코드\r\n가나다,"첫 줄\n둘째 줄",00123\r\n=1+1,"따옴표 ""예시""",42', '샘플.csv');
  const sheet = book.sheets[book.sheetOrder[0]]; assert.equal(sheet.cellData![1][1].v, '첫 줄\n둘째 줄'); assert.equal(sheet.cellData![1][2].v, '00123'); assert.equal(sheet.cellData![2][0].f, undefined); assert.equal(sheet.cellData![2][2].v, 42);
  const csv = exportCsv(book, book.sheetOrder[0]); assert.ok(csv.includes("'=1+1")); const restored = importCsv(csv, '다시.csv'); assert.equal(restored.sheets[restored.sheetOrder[0]].cellData![1][1].v, '첫 줄\n둘째 줄');
});
test('invalid CSV and oversized workbook are rejected', () => {
  assert.throws(() => importCsv('"unclosed', 'bad.csv'));
  const book = newBook(); book.sheets[book.sheetOrder[0]].rowCount = 10001; assert.throws(() => validateBook(book));
  assert.throws(() => validateBook({}));
});
test('corrupt XLSX is rejected', async () => { await assert.rejects(importXlsx(new TextEncoder().encode('bad zip').buffer, 'bad.xlsx')); });

test('1904 workbook dates, string formula results, borders and empty-row formatting survive normalized export', async () => {
  const original = new ExcelJS.Workbook(); original.properties.date1904 = true;
  const ws = original.addWorksheet('날짜 확인');
  ws.getCell('A1').value = new Date('2026-11-06T12:00:00.000Z'); ws.getCell('A1').numFmt = 'yyyy-mm-dd hh:mm';
  ws.getCell('B1').value = { formula: '"한글 결과"', result: '한글 결과' };
  ws.getCell('C1').value = '=SUM(A1:A2)';
  ws.getCell('D1').value = 42; ws.getCell('D1').border = { top: { style: 'double', color: { argb: 'FF18794E' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } }; ws.getCell('D1').font = { strike: true };
  ws.getRow(3).height = 42; ws.getRow(3).hidden = true; ws.getColumn(2).hidden = true;
  const bytes = Uint8Array.from(await original.xlsx.writeBuffer() as unknown as Uint8Array).buffer;
  const book = await importXlsx(bytes, '날짜.xlsx'); const sheet = book.sheets[book.sheetOrder[0]];
  assert.equal(sheet.cellData![0][0].v, new Date('2026-11-06T12:00:00.000Z').getTime() / 86400000 + 25569);
  assert.equal(sheet.cellData![0][1].v, '한글 결과'); assert.equal(sheet.cellData![0][2].f, undefined);
  const restored = new ExcelJS.Workbook(); await restored.xlsx.load(await exportXlsx(book)); const target = restored.worksheets[0];
  assert.equal((target.getCell('A1').value as Date).toISOString(), '2026-11-06T12:00:00.000Z');
  assert.equal(target.getCell('B1').result, '한글 결과'); assert.equal(target.getCell('C1').value, '=SUM(A1:A2)');
  assert.equal(target.getCell('D1').border.top?.style, 'double'); assert.equal(target.getCell('D1').font.strike, true);
  assert.equal(target.getRow(3).height, 42); assert.equal(target.getRow(3).hidden, true); assert.equal(target.getColumn(2).hidden, true);
});

test('malformed snapshot addresses, prototypes, dimensions and merge areas are rejected before export', async () => {
  const cases: unknown[] = [];
  const duplicate = newBook(); duplicate.sheetOrder.push(duplicate.sheetOrder[0]); cases.push(duplicate);
  const malicious = JSON.parse(JSON.stringify(newBook())); malicious.sheets[malicious.sheetOrder[0]].cellData = JSON.parse('{"__proto__":{"0":{"v":"bad"}}}'); cases.push(malicious);
  const fractional = newBook(); fractional.sheets[fractional.sheetOrder[0]].cellData = { 0: { 0.5: { v: 'bad' } } }; cases.push(fractional);
  const negative = newBook(); negative.sheets[negative.sheetOrder[0]].rowCount = -1; cases.push(negative);
  const invalidValue = newBook(); (invalidValue.sheets[invalidValue.sheetOrder[0]].cellData as unknown as Record<string, unknown>)[0] = { 0: { v: { text: 'bad' } } }; cases.push(invalidValue);
  const enormous = newBook(); enormous.sheets[enormous.sheetOrder[0]].mergeData = [{ startRow: 0, endRow: 999999999, startColumn: 0, endColumn: 999999999 }]; cases.push(enormous);
  const overlapping = newBook(); overlapping.sheets[overlapping.sheetOrder[0]].mergeData = [{ startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }, { startRow: 2, endRow: 4, startColumn: 1, endColumn: 3 }]; cases.push(overlapping);
  for (const invalid of cases) { assert.throws(() => validateBook(invalid)); await assert.rejects(exportXlsx(invalid as ReturnType<typeof newBook>)); }
  const cyclic = newBook() as unknown as Record<string, unknown>; cyclic.extra = cyclic; assert.throws(() => validateBook(cyclic));
});

test('excessive XLSX declared expansion is rejected before decompression', async () => {
  const bytes = new Uint8Array(await exportXlsx(newBook(true))); const view = new DataView(bytes.buffer);
  let central = -1; for (let i = 0; i < bytes.length - 46; i++) if (view.getUint32(i, true) === 0x02014b50) { central = i; break; }
  assert.ok(central >= 0); view.setUint32(central + 24, 70 * 1024 * 1024, true);
  await assert.rejects(importXlsx(bytes.buffer, 'excessive.xlsx'), /64MB/);
});
