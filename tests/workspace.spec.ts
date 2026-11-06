import { test, expect, type Page, type TestInfo } from '@playwright/test';
import { createRequire } from 'node:module';
import { readFile, writeFile } from 'node:fs/promises';
const ExcelJS = createRequire(import.meta.url)('exceljs') as typeof import('exceljs');

async function ready(page: Page) {
  await expect(page.locator('#tab-0')).toBeVisible();
  await expect(page.locator('#loading')).toBeHidden();
  await expect(page.locator('#editor canvas:visible').first()).toBeVisible();
}
async function selectCell(page: Page, address: string) {
  await page.locator('#address').fill(address);
  await page.locator('#address').press('Enter');
  await expect(page.locator('#selection-status')).toContainText(address);
}
async function setCell(page: Page, address: string, value: string) {
  await selectCell(page, address);
  await page.locator('#cell-input').fill(value);
  await page.locator('#cell-form button[type="submit"]').click();
  await expect(page.locator('#save-state')).toHaveText('이 브라우저에 보관됨');
}
async function download(page: Page, selector: string, info: TestInfo, name: string) {
  const event = page.waitForEvent('download');
  await page.locator(selector).click();
  const file = await event;
  const path = info.outputPath(name);
  await file.saveAs(path);
  return path;
}
async function downloadedBook(page: Page, info: TestInfo, name: string) {
  const path = await download(page, '#export', info, name);
  const book = new ExcelJS.Workbook();
  await book.xlsx.readFile(path);
  return book;
}
async function importFile(page: Page, path: string) {
  await page.locator('#open').click();
  await expect(page.locator('#modal-title')).toHaveText('파일 가져오기');
  const chooser = page.waitForEvent('filechooser');
  await page.locator('#modal-ok').click();
  await (await chooser).setFiles(path);
  await expect(page.locator('#toast')).toContainText('파일을 열었습니다');
  await ready(page);
}

test('real example formulas, formatting, undo, download, browser save and reload', async ({ page }, info) => {
  const errors: string[] = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.goto('/'); await ready(page);
  await page.locator('#example').click();
  await expect(page.locator('#document-name')).toHaveText('예시 가계부');
  await expect(page.locator('#save-state')).toHaveText('이 브라우저에 보관됨');
  await selectCell(page, 'F6');
  await expect(page.locator('#cell-input')).toHaveValue('=D6-E6');
  await expect(page.locator('#selection-status')).toContainText('2,835,000');
  await setCell(page, 'D2', '4000000');
  await selectCell(page, 'F6'); await expect(page.locator('#cell-input')).toHaveValue('=D6-E6');
  await expect(page.locator('#selection-status')).toContainText('3,835,000');
  await setCell(page, 'G1', '=SUM(E2:E5)');
  await selectCell(page, 'G1'); await expect(page.locator('#cell-input')).toHaveValue('=SUM(E2:E5)');
  await expect(page.locator('#selection-status')).toContainText('165,000');
  await page.locator('#bold').click();
  await page.locator('#font-size').selectOption('14');
  await page.locator('#align-center').click();
  await page.locator('#border').click();
  await setCell(page, 'G2', 'temporary');
  await page.locator('#undo').click();
  await selectCell(page, 'G2'); await expect(page.locator('#cell-input')).toHaveValue('');
  await page.locator('#redo').click();
  await selectCell(page, 'G2'); await expect(page.locator('#cell-input')).toHaveValue('temporary');
  await page.locator('#save').click(); await expect(page.locator('#save-state')).toHaveText('이 브라우저에 보관됨');
  const book = await downloadedBook(page, info, 'edited-example.xlsx');
  const sheet = book.worksheets[0];
  expect(sheet.getCell('D2').value).toBe(4000000);
  expect(sheet.getCell('F6').result).toBe(3835000);
  expect(sheet.getCell('G1').formula).toBe('SUM(E2:E5)');
  expect(sheet.getCell('G1').result).toBe(165000);
  expect(sheet.getCell('G1').font.bold).toBe(true);
  expect(sheet.getCell('G1').font.size).toBe(14);
  expect(sheet.getCell('G1').alignment.horizontal).toBe('center');
  expect(sheet.getCell('G1').border.top?.style).toBe('thin');
  await page.reload(); await ready(page);
  await expect(page.locator('#document-name')).toHaveText('예시 가계부');
  await selectCell(page, 'G2'); await expect(page.locator('#cell-input')).toHaveValue('temporary');
  await page.locator('#search').fill('없는 문서'); await expect(page.locator('#library')).toContainText('검색 결과가 없어요');
  await page.locator('#search').fill('가계부'); await expect(page.locator('#library .book-open')).toHaveCount(1);
  await page.screenshot({ path: info.outputPath('desktop-editor.png') });
  expect(errors).toEqual([]);
});

test('XLSX upload and multi-sheet content, CSV and JSON backup roundtrip', async ({ page }, info) => {
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet('가져온 시트');
  sheet.addRow(['항목', '수량', '단가', '금액']); sheet.addRow(['한글 품목', 3, 12000, { formula: 'B2*C2', result: 36000 }]);
  source.addWorksheet('추가 자료').getCell('A1').value = '두 번째 시트';
  const path = info.outputPath('upload.xlsx'); await source.xlsx.writeFile(path);
  await page.goto('/'); await ready(page); await importFile(page, path);
  await selectCell(page, 'D2'); await expect(page.locator('#cell-input')).toHaveValue('=B2*C2');
  await page.locator('#tab-1').click(); await page.locator('#insert-sheet').click();
  await setCell(page, 'A1', '새 시트 입력');
  const multi = await downloadedBook(page, info, 'multiple-sheets.xlsx');
  expect(multi.worksheets).toHaveLength(3);
  expect(multi.getWorksheet('가져온 시트')!.getCell('D2').result).toBe(36000);
  expect(multi.worksheets.some(s => s.getCell('A1').value === '새 시트 입력')).toBe(true);
  const backup = await download(page, '#backup', info, 'backup.json');
  await page.locator('#new').click(); await expect(page.locator('#document-name')).toHaveText('새 통합문서');
  await importFile(page, backup);
  const restored = await downloadedBook(page, info, 'restored.xlsx');
  expect(restored.worksheets).toHaveLength(3); expect(restored.getWorksheet('추가 자료')!.getCell('A1').value).toBe('두 번째 시트');
  const csvPath = info.outputPath('korean.csv');
  await writeFile(csvPath, '\uFEFF이름,설명,코드\r\n가나다,"첫 줄\n둘째 줄",00123\r\n=1+1,수식 아닌 문자열,42');
  await importFile(page, csvPath);
  await selectCell(page, 'A3'); await expect(page.locator('#cell-input')).toHaveValue('=1+1');
  const csv = await download(page, '#csv', info, 'exported.csv');
  const text = await readFile(csv, 'utf8'); expect(text).toContain("'=1+1"); expect(text).toContain('00123'); expect(text).toContain('첫 줄\n둘째 줄');
  const csvXlsx = await downloadedBook(page, info, 'csv-as-xlsx.xlsx');
  expect(csvXlsx.worksheets[0].getCell('A3').value).toBe('=1+1'); expect(csvXlsx.worksheets[0].getCell('A3').formula).toBeUndefined();
  expect(csvXlsx.worksheets[0].getCell('C2').value).toBe('00123');
  await importFile(page, csv); await selectCell(page, 'B2'); await expect(page.locator('#cell-input')).toHaveValue('첫 줄\n둘째 줄');
  await page.locator('#cell-form button[type="submit"]').click();
  await expect(page.locator('#save-state')).toHaveText('이 브라우저에 보관됨');
  const multiline = await downloadedBook(page, info, 'multiline-reapplied.xlsx');
  expect(multiline.worksheets[0].getCell('B2').value).toBe('첫 줄\n둘째 줄');
  expect(errors).toEqual([]);
});

test('responsive editor and dialogs at 360, 390, 768 and 1440 pixels', async ({ page }, info) => {
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  await page.goto('/'); await ready(page); await page.locator('#example').click();
  await expect(page.locator('#document-name')).toHaveText('예시 가계부');
  for (const viewport of [{ width: 360, height: 800 }, { width: 390, height: 844 }, { width: 768, height: 1024 }, { width: 1440, height: 900 }]) {
    await page.setViewportSize(viewport);
    await expect.poll(() => page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    await selectCell(page, 'F6'); await expect(page.locator('#cell-input')).toHaveValue('=D6-E6');
    await expect(page.locator('#selection-status')).toContainText('2,835,000');
    await page.screenshot({ path: info.outputPath(`editor-${viewport.width}.png`) });
    if (viewport.width <= 800) {
      await page.locator('#library-toggle').click(); await expect(page.locator('#library-dialog')).toBeVisible();
      await page.keyboard.press('Shift+Tab');
      expect(await page.evaluate(() => Boolean(document.activeElement?.closest('#library-dialog')))).toBe(true);
      await page.screenshot({ path: info.outputPath(`library-${viewport.width}.png`) });
      await page.keyboard.press('Escape'); await expect(page.locator('#library-dialog')).toBeHidden();
      await expect(page.locator('#library-toggle')).toBeFocused();
    }
    await page.locator('#ai-open').click(); await expect(page.locator('#ai-dialog')).toBeVisible();
    await expect(page.locator('#ai-status')).toContainText('AI 연결 전');
    await page.screenshot({ path: info.outputPath(`ai-${viewport.width}.png`) });
    await page.keyboard.press('Escape'); await expect(page.locator('#ai-dialog')).toBeHidden();
    await expect(page.locator('#ai-open')).toBeFocused();
  }
  expect(errors).toEqual([]);
});

test('keyboard ribbon and personalization, honest disconnected AI without data transmission', async ({ page }, info) => {
  let aiRequests = 0; page.on('request', request => { if (new URL(request.url()).pathname === '/api/ai') aiRequests++; });
  await page.goto('/'); await ready(page);
  await page.locator('#tab-0').focus(); await page.keyboard.press('ArrowRight');
  await expect(page.locator('#tab-1')).toBeFocused(); await expect(page.locator('#panel-1')).toBeVisible();
  await page.keyboard.press('ArrowLeft'); await expect(page.locator('#panel-0')).toBeVisible();
  await page.locator('#tab-3').click(); await page.locator('#personalize').click();
  await page.locator('#accent').selectOption('blue'); await page.locator('#density').selectOption('compact'); await page.locator('#modal-ok').click();
  await expect(page.locator('html')).toHaveAttribute('data-accent', 'blue'); await expect(page.locator('html')).toHaveAttribute('data-density', 'compact');
  await page.reload(); await ready(page); await expect(page.locator('html')).toHaveAttribute('data-accent', 'blue');
  await page.locator('#ai-open').click(); await expect(page.locator('#ai-status')).toContainText('AI 연결 전');
  await expect(page.locator('#include-cells')).not.toBeChecked();
  await page.locator('#ai-prompt').fill('B2:B10 합계 수식을 만들어 줘'); await page.locator('#ai-send').click();
  await expect(page.locator('#toast')).toContainText('AI가 아직 연결되지 않았습니다');
  await expect(page.locator('#ai-result')).toBeHidden(); expect(aiRequests).toBe(0);
  await page.screenshot({ path: info.outputPath('disconnected-ai.png') });
});
