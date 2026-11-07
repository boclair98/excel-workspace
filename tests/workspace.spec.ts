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
  await page.locator('#hint-example').click();
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
  await page.locator('#library-toggle').click();
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
  await page.locator('#library-toggle').click();
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

test('AI edit reviews typed cell changes, applies them to XLSX and undoes them together', async ({ page }, info) => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet('AI 검증');
  sheet.getCell('A2').value = '범위 밖';
  sheet.getCell('B2').value = '  상품  ';
  sheet.getCell('B2').font = { bold: true };
  sheet.getCell('C2').value = '지울 값';
  sheet.getCell('D2').value = '코드';
  sheet.getCell('C3').value = true;
  sheet.getCell('D3').value = { formula: 'B3+2', result: 2 };
  const file = info.outputPath('ai-source.xlsx');
  await source.xlsx.writeFile(file);
  const requests: { mode: string; context: string }[] = [];
  const errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.route('**/api/status', route => route.fulfill({ json: { aiAvailable: true, provider: 'openai', managed: true, signedIn: true, openAIUserRef: '1'.repeat(32) } }));
  await page.route('**/api/ai', route => {
    requests.push(route.request().postDataJSON());
    return route.fulfill({ json: { explanation: '상품명과 코드, 빈칸을 정리합니다.', edits: [
      { address: 'B2', kind: 'set', value: '상품' },
      { address: 'C2', kind: 'clear', value: null },
      { address: 'D2', kind: 'set', value: '00123' },
      { address: 'B3', kind: 'set', value: 0 },
      { address: 'C3', kind: 'set', value: false },
    ] } });
  });
  await page.goto('/');
  await ready(page);
  await importFile(page, file);
  await selectCell(page, 'B2:D3');
  await page.locator('#ai-open').click();
  await expect(page.locator('#ai-status')).toHaveText('OpenAI 연결 설정됨');
  await page.getByRole('radio', { name: '표 바꾸기' }).check();
  await expect(page.locator('#ai-range-label')).toHaveText('선택 범위 B2:D3');
  await page.locator('#ai-prompt').fill('상품명과 코드를 정리하고 빈칸을 0으로 채워 줘');
  await page.locator('#ai-send').click();
  await expect(page.locator('#ai-error')).toContainText('전송에 동의');
  expect(requests).toHaveLength(0);
  await page.locator('#include-cells').check();
  await page.locator('#ai-send').click();
  await expect(page.locator('#ai-result-label')).toHaveText('B2:D3 · 5개 셀 변경안');
  await expect(page.locator('#ai-change-rows tr')).toHaveCount(5);
  await expect(page.locator('#ai-change-rows tr').first()).toContainText('“  상품  ”');
  await expect(page.locator('#ai-change-rows tr').last()).toContainText('FALSE');
  await expect(page.locator('#cell-input')).toHaveValue('  상품  ');
  expect(requests).toHaveLength(1);
  expect(requests[0].mode).toBe('edit');
  const sent = JSON.parse(requests[0].context);
  expect(sent.range).toBe('B2:D3');
  expect(sent.values).toEqual([['  상품  ', '지울 값', '코드'], [null, true, 2]]);
  expect(sent.formulas[1][2]).toBe('=B3+2');
  expect(requests[0].context).not.toContain('범위 밖');
  await page.locator('#ai-apply').click();
  await expect(page.locator('#ai-result-label')).toHaveText('5개 셀에 적용했습니다');
  await page.keyboard.press('Escape');
  await expect(page.locator('#ai-open')).toBeFocused();
  const applied = (await downloadedBook(page, info, 'ai-applied.xlsx')).worksheets[0];
  expect(applied.getCell('A2').value).toBe('범위 밖');
  expect(applied.getCell('B2').value).toBe('상품');
  expect(applied.getCell('B2').font.bold).toBe(true);
  expect(applied.getCell('C2').value).toBeNull();
  expect(applied.getCell('D2').value).toBe('00123');
  expect(applied.getCell('B3').value).toBe(0);
  expect(applied.getCell('C3').value).toBe(false);
  expect(applied.getCell('D3').formula).toBe('B3+2');
  await page.locator('#undo').click();
  const undone = (await downloadedBook(page, info, 'ai-undone.xlsx')).worksheets[0];
  expect(undone.getCell('B2').value).toBe('  상품  ');
  expect(undone.getCell('C2').value).toBe('지울 값');
  expect(undone.getCell('D2').value).toBe('코드');
  expect(undone.getCell('B3').value).toBeNull();
  expect(undone.getCell('C3').value).toBe(true);
  expect(errors).toEqual([]);
});

test('AI rejects out-of-range changes and fits the reviewed changes at four sizes', async ({ page }, info) => {
  let expired = true;
  let invalid = true;
  let redirects = 0;
  const errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.route('**/api/status', route => route.fulfill({ json: { aiAvailable: true, provider: 'claude', managed: true, signedIn: true, openAIUserRef: '2'.repeat(32) } }));
  await page.route('https://example.invalid/expired-login', route => {
    redirects++;
    return route.fulfill({ status: 200, body: 'Login' });
  });
  await page.route('**/api/ai', route => expired
    ? route.fulfill({ status: 307, headers: { Location: 'https://example.invalid/expired-login' } })
    : route.fulfill({ json: {
    explanation: '선택한 범위의 빈칸에 새 값을 입력합니다.',
    edits: [
      { address: 'A1', kind: 'set', value: 0 },
      { address: invalid ? 'D1' : 'B1', kind: 'set', value: '공급처 연락처 확인 후 재고 담당자에게 검토 요청' },
      { address: 'C1', kind: 'set', value: '=1+1' },
    ],
  } }));
  await page.goto('/');
  await ready(page);
  await selectCell(page, 'A1:C1');
  await page.locator('#ai-open').click();
  await expect(page.locator('#ai-status')).toContainText('관리형 AI');
  await page.getByRole('radio', { name: '표 바꾸기' }).check();
  await page.locator('#ai-prompt').fill('빈칸을 채워 줘');
  await page.locator('#include-cells').check();
  await page.locator('#ai-send').click();
  await expect(page.locator('#ai-error')).toContainText('로그인 상태');
  expect(redirects).toBe(0);
  expired = false;
  await page.locator('#ai-send').click();
  await expect(page.locator('#ai-error')).toContainText('범위 밖');
  await expect(page.locator('#ai-apply')).toBeHidden();
  await expect(page.locator('#cell-input')).toHaveValue('');
  invalid = false;
  await page.locator('#ai-send').click();
  await expect(page.locator('#ai-result-label')).toHaveText('A1:C1 · 3개 셀 변경안');
  await expect(page.locator('#ai-error')).toBeHidden();
  for (const viewport of [{ width: 360, height: 800 }, { width: 390, height: 844 }, { width: 768, height: 1024 }, { width: 1440, height: 900 }]) {
    await page.setViewportSize(viewport);
    await page.locator('#ai-apply').scrollIntoViewIfNeeded();
    await expect(page.locator('#ai-apply')).toBeInViewport();
    expect(await page.locator('#ai-apply').evaluate(button => button.getBoundingClientRect().height)).toBeGreaterThanOrEqual(44);
    expect(await page.evaluate(() => {
      const panel = document.querySelector('.ai-content')!;
      const preview = document.getElementById('ai-changes')!;
      return document.documentElement.scrollWidth <= innerWidth && panel.scrollWidth <= panel.clientWidth && preview.scrollWidth <= preview.clientWidth;
    })).toBe(true);
    await page.screenshot({ path: info.outputPath(`ai-preview-${viewport.width}.png`) });
  }
  await page.locator('#ai-apply').focus();
  await page.keyboard.press('Enter');
  await expect(page.locator('#ai-result-label')).toHaveText('3개 셀에 적용했습니다');
  await page.keyboard.press('Escape');
  await expect(page.locator('#ai-open')).toBeFocused();
  const applied = (await downloadedBook(page, info, 'ai-literal.xlsx')).worksheets[0];
  expect(applied.getCell('C1').value).toBe('=1+1');
  expect(applied.getCell('C1').formula).toBeUndefined();
  expect(errors).toEqual([]);
});

test('responsive editor and dialogs at 360, 390, 768 and 1440 pixels', async ({ page }, info) => {
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  await page.goto('/'); await ready(page); await page.locator('#hint-example').click();
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
  await expect(page.locator('#sidebar')).toBeHidden();
  const wideSheet = (await page.locator('.sheet-stage').boundingBox())!.width;
  await page.locator('#tab-0').focus(); await page.keyboard.press('ArrowRight');
  await expect(page.locator('#tab-1')).toBeFocused(); await expect(page.locator('#panel-1')).toBeVisible();
  await page.keyboard.press('ArrowLeft'); await expect(page.locator('#panel-0')).toBeVisible();
  await page.locator('#tab-3').click(); await page.locator('#personalize').click();
  await page.locator('#accent').selectOption('blue');
  await page.locator('#density').selectOption('compact');
  await page.locator('#layout').selectOption('library');
  await page.locator('#ai-width').selectOption('compact');
  await page.locator('#modal-ok').click();
  await expect(page.locator('html')).toHaveAttribute('data-accent', 'blue'); await expect(page.locator('html')).toHaveAttribute('data-density', 'compact');
  await page.reload(); await ready(page); await expect(page.locator('html')).toHaveAttribute('data-accent', 'blue');
  await expect(page.locator('html')).toHaveAttribute('data-ai-width', 'compact');
  await expect(page.locator('#sidebar')).toBeVisible();
  expect((await page.locator('.sheet-stage').boundingBox())!.width).toBeLessThan(wideSheet);
  await page.locator('#appearance').click();
  await expect(page.locator('#layout')).toHaveValue('library');
  await page.keyboard.press('Escape');
  await expect(page.locator('#appearance')).toBeFocused();
  await page.locator('#ai-open').click(); await expect(page.locator('#ai-status')).toContainText('AI 연결 전');
  await expect(page.locator('#include-cells')).not.toBeChecked();
  await page.locator('#ai-prompt').fill('B2:B10 합계 수식을 만들어 줘'); await page.locator('#ai-send').click();
  await expect(page.locator('#ai-error')).toContainText('AI가 아직 연결되지 않았습니다');
  await expect(page.locator('#ai-result')).toBeHidden(); expect(aiRequests).toBe(0);
  await page.screenshot({ path: info.outputPath('disconnected-ai.png') });
});
