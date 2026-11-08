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

test('guided notice uses plain chat, copies the draft and carries context into revisions', async ({ page, context }, info) => {
  const requests: { mode: string; prompt: string; history: unknown[] }[] = [];
  const draft = '경로당 식사 시간 변경 안내\n\n3월 15일부터 오전 11시에 식사를 제공합니다.';
  const revision = '3월 15일부터 경로당 식사는 오전 11시에 시작합니다.';
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.route('**/api/status', route => route.fulfill({ json: { aiAvailable: true, provider: 'openai', managed: true, signedIn: true, sharedOpenAI: true } }));
  await page.route('**/api/ai', route => {
    requests.push(route.request().postDataJSON());
    return route.fulfill({ json: { reply: requests.length === 1 ? draft : revision } });
  });
  await page.goto('/');
  await ready(page);
  await setCell(page, 'B2', '채팅에 보내지 않을 셀');
  await page.locator('#ai-open').focus();
  await page.keyboard.press('Enter');
  await expect(page.locator('#ai-status')).toHaveText('AI에 연결됐어요.');
  await page.getByRole('button', { name: /안내문/ }).click();
  await expect(page.getByLabel(/안내 제목/)).toBeFocused();
  await page.getByRole('button', { name: '초안 받기' }).click();
  await expect(page.locator('#ai-error')).toContainText('안내 제목');
  expect(requests).toHaveLength(0);
  await page.getByLabel(/안내 제목/).fill('경로당 식사 시간 변경');
  await page.getByLabel(/안내할 내용/).fill('3월 15일부터 오전 11시에 식사를 제공합니다.');
  await page.getByRole('button', { name: '초안 받기' }).click();
  const answers = page.locator('.ai-message[data-role="assistant"]');
  await expect(answers.first()).toContainText('3월 15일부터 오전 11시');
  expect(requests[0]).toEqual({
    mode: 'chat',
    prompt: '안내문 초안을 작성해 주세요.\n\n안내 제목: 경로당 식사 시간 변경\n안내할 내용: 3월 15일부터 오전 11시에 식사를 제공합니다.',
    history: [],
  });
  await answers.first().getByRole('button', { name: '복사', exact: true }).click();
  await expect(answers.first().getByRole('button', { name: '복사했어요' })).toBeVisible();
  const clipboard = await page.evaluate(() => navigator.clipboard.readText());
  expect(clipboard.replace(/\r\n/g, '\n')).toBe(draft);

  await page.getByRole('checkbox', { name: '셀 연결' }).uncheck();
  await page.getByLabel('메시지', { exact: true }).fill('한 문장으로 줄여 주세요.');
  await page.getByRole('button', { name: '보내기', exact: true }).click();
  await expect(answers.last()).toContainText(revision);
  expect(requests[1]).toEqual({
    mode: 'chat',
    prompt: '한 문장으로 줄여 주세요.',
    history: [{ role: 'user', content: requests[0].prompt }, { role: 'assistant', content: draft }],
  });
  await page.screenshot({ path: info.outputPath('ai-draft-desktop.png') });
  await page.keyboard.press('Escape');
  await expect(page.locator('#ai-open')).toBeFocused();
  await expect(page.locator('#cell-input')).toHaveValue(draft);
  await page.locator('#ai-open').click();
  await expect(answers.last()).toContainText(revision);
  await page.getByRole('button', { name: '새 대화' }).click();
  await expect(answers).toHaveCount(0);
});

test('one chat prompt automatically applies values and a formula to XLSX and undoes together', async ({ page }, info) => {
  await page.setViewportSize({ width: 390, height: 844 });
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet('채팅 적용 검증');
  sheet.getCell('A2').value = '범위 밖 유지';
  sheet.getCell('B2').value = 12;
  sheet.getCell('C2').value = '  접수  ';
  sheet.getCell('D3').value = { formula: 'B2*2', result: 24 };
  const file = info.outputPath('ai-source.xlsx');
  await source.xlsx.writeFile(file);
  const requests: { mode: string; context: string }[] = [];
  const edits = [
    { address: 'B3', kind: 'set', value: 8 },
    { address: 'C2', kind: 'set', value: '접수' },
    { address: 'C3', kind: 'formula', value: '=SUM(B2:B3)' },
  ];
  await page.route('**/api/status', route => route.fulfill({ json: { aiAvailable: true, provider: 'openai', managed: true, signedIn: true, sharedOpenAI: true } }));
  await page.route('**/api/ai', route => {
    requests.push(route.request().postDataJSON());
    return route.fulfill({ json: { reply: '세 셀을 수정할 수 있어요.', edits: requests.length === 1
      ? [...edits, { address: 'A2', kind: 'clear', value: null }] : edits } });
  });
  await page.goto('/');
  await ready(page);
  await importFile(page, file);
  await page.getByRole('button', { name: 'AI 채팅', exact: true }).click();
  await page.getByLabel('메시지', { exact: true }).fill('B2:D3에서 B3에 8, C2의 공백을 지우고 C3에 B2:B3 합계 수식을 넣어줘.');
  await expect(page.locator('#ai-range')).toHaveText('작업 범위 B2:D3');
  await page.getByLabel('메시지', { exact: true }).press('Enter');
  await expect(page.locator('#ai-error')).toContainText('범위 밖');
  await expect(page.getByRole('button', { name: '엑셀에 적용' })).toHaveCount(0);
  expect(requests[0].mode).toBe('chat');
  expect(JSON.parse(requests[0].context)).toEqual({ sheet: '채팅 적용 검증', range: 'B2:D3', values: [[12, '  접수  ', null], [null, null, 24]], formulas: [['', '', ''], ['', '', '=B2*2']] });
  await page.getByRole('button', { name: '보내기', exact: true }).click();
  await expect(page.locator('.ai-plan-summary')).toContainText('3개 셀을 적용했어요');
  await expect(page.getByRole('table')).toContainText('=SUM(B2:B3)');
  await page.screenshot({ path: info.outputPath('ai-sheet-preview.png') });
  await page.getByRole('button', { name: 'AI 채팅 닫기' }).click();
  await expect(page.locator('#ai-dialog')).toBeHidden();
  const updated = (await downloadedBook(page, info, 'ai-updated.xlsx')).getWorksheet(1)!;
  expect(updated.getCell('A2').value).toBe('범위 밖 유지');
  expect(updated.getCell('B3').value).toBe(8);
  expect(updated.getCell('C2').value).toBe('접수');
  expect(updated.getCell('C3').formula).toBe('SUM(B2:B3)');
  expect(updated.getCell('D3').formula).toBe('B2*2');
  await page.locator('#undo').click();
  const restored = (await downloadedBook(page, info, 'ai-undone.xlsx')).getWorksheet(1)!;
  expect(restored.getCell('B3').value).toBeNull();
  expect(restored.getCell('C2').value).toBe('  접수  ');
  expect(restored.getCell('C3').value).toBeNull();
});

test('mobile application fields and composer remain usable and retain input after failure', async ({ page }, info) => {
  await page.setViewportSize({ width: 390, height: 844 });
  const requests: { mode: string; prompt: string; history: unknown[] }[] = [];
  await page.route('**/api/status', route => route.fulfill({ json: { aiAvailable: true, provider: 'openai', managed: true, signedIn: true, sharedOpenAI: true } }));
  await page.route('**/api/ai', route => {
    requests.push(route.request().postDataJSON());
    return requests.length === 1
      ? route.fulfill({ status: 429, json: { error: '잠시 후 다시 보내 주세요.' } })
      : route.fulfill({ json: { reply: '스마트폰 교육 신청서\n\n신청인 [직접 입력]\n사진을 주고받는 방법을 배우고자 교육을 신청합니다.' } });
  });
  await page.goto('/');
  await ready(page);
  await page.getByRole('button', { name: 'AI 채팅', exact: true }).click();
  await page.getByRole('button', { name: /신청서/ }).click();
  await expect(page.getByLabel(/이름.*선택/)).toBeVisible();
  const topic = page.getByLabel(/신청 항목/);
  const reason = page.getByLabel(/신청 사유/);
  await topic.fill('주민센터 스마트폰 교육');
  await reason.fill('가족과 사진을 주고받는 방법을 배우고 싶어요.');
  await page.screenshot({ path: info.outputPath('ai-application-mobile.png') });
  await page.setViewportSize({ width: 390, height: 480 });
  const submit = page.getByRole('button', { name: '초안 받기' });
  await expect(submit).toBeInViewport({ ratio: 1 });
  await submit.click();
  await expect(page.locator('#ai-error')).toContainText('잠시 후 다시 보내 주세요.');
  await expect(topic).toHaveValue('주민센터 스마트폰 교육');
  await expect(reason).toHaveValue('가족과 사진을 주고받는 방법을 배우고 싶어요.');
  await submit.click();
  await expect(page.locator('.ai-message[data-role="assistant"]')).toContainText('스마트폰 교육 신청서');
  expect(requests[1]).toEqual({
    mode: 'chat',
    prompt: '신청서 초안을 작성해 주세요.\n\n신청 항목: 주민센터 스마트폰 교육\n신청 사유: 가족과 사진을 주고받는 방법을 배우고 싶어요.',
    history: [],
  });
  await expect(page.getByRole('button', { name: '보내기', exact: true })).toBeInViewport({ ratio: 1 });
});

test('responsive editor and dialogs at 360, 390, 768 and 1440 pixels', async ({ page }, info) => {
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  await page.route('**/api/status', route => route.fulfill({ json: { aiAvailable: true, provider: 'openai', managed: true, signedIn: true, sharedOpenAI: true } }));
  await page.goto('/'); await ready(page); await page.locator('#hint-example').click();
  await expect(page.locator('#document-name')).toHaveText('예시 가계부');
  for (const viewport of [{ width: 360, height: 800 }, { width: 390, height: 844 }, { width: 768, height: 1024 }, { width: 1440, height: 900 }]) {
    await page.setViewportSize(viewport);
    await expect.poll(() => page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    await selectCell(page, 'F6'); await expect(page.locator('#cell-input')).toHaveValue('=D6-E6');
    await expect(page.locator('#selection-status')).toContainText('2,835,000');
    if (viewport.width <= 520) {
      for (const label of ['#library-toggle > span', '#open > span', '#save > span', '#export .desktop-label'])
        await expect(page.locator(label)).toBeInViewport({ ratio: 1 });
    }
    await page.screenshot({ path: info.outputPath(`editor-${viewport.width}.png`) });
    if (viewport.width <= 800) {
      await page.locator('#library-toggle').click(); await expect(page.locator('#library-dialog')).toBeVisible();
      await page.keyboard.press('Shift+Tab');
      expect(await page.evaluate(() => Boolean(document.activeElement?.closest('#library-dialog')))).toBe(true);
      await page.screenshot({ path: info.outputPath(`library-${viewport.width}.png`) });
      await page.keyboard.press('Escape'); await expect(page.locator('#library-dialog')).toBeHidden();
      await expect(page.locator('#library-toggle')).toBeFocused();
    }
    await expect(page.locator('#ai-open > span')).toBeVisible();
    await page.locator('#ai-open').click(); await expect(page.locator('#ai-dialog')).toBeVisible();
    await expect(page.locator('#ai-status')).toHaveText('AI에 연결됐어요.');
    await expect(page.locator('#ai-send')).toBeInViewport({ ratio: 1 });
    expect(await page.locator('#ai-dialog').evaluate(dialog => dialog.scrollWidth <= dialog.clientWidth)).toBe(true);
    expect(await page.locator('#ai-thread').evaluate(thread => thread.scrollWidth <= thread.clientWidth)).toBe(true);
    if (viewport.width <= 800) expect((await page.locator('#ai-open').boundingBox())!.height).toBeGreaterThanOrEqual(44);
    for (const control of ['[data-template="notice"]', '[data-template="application"]', '#ai-send', '[data-close="ai-dialog"]'])
      expect((await page.locator(control).boundingBox())!.height).toBeGreaterThanOrEqual(44);
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
  await page.locator('#ai-open').click(); await expect(page.locator('#ai-status')).toContainText('AI에 연결할 수 없어요');
  await page.locator('#ai-prompt').fill('B2:B10 합계 수식을 만들어 줘');
  await expect(page.locator('#ai-send')).toBeDisabled();
  await expect(page.locator('.ai-message')).toHaveCount(0);
  expect(aiRequests).toBe(0);
  await page.screenshot({ path: info.outputPath('disconnected-ai.png') });
});

test('docked AI keeps sheet usable and text cursor stable while applying a table automatically', async ({ page }) => {
  await page.route('**/api/status', route => route.fulfill({ json: { aiAvailable: true, provider: 'openai', managed: false, signedIn: true } }));
  await page.route('**/api/ai', route => {
    const input = route.request().postDataJSON();
    expect(JSON.parse(input.context).range).toBe('A1:J20');
    return route.fulfill({ json: { reply: '명단을 작성합니다.', edits: [
      { address: 'A1', kind: 'set', value: '이름' },
      { address: 'B1', kind: 'set', value: '수량' },
      { address: 'B2', kind: 'set', value: 3 },
    ] } });
  });
  await page.goto('/'); await ready(page);
  await page.locator('#ai-open').click();
  const prompt = page.getByLabel('메시지', { exact: true });
  await expect(prompt).toBeFocused();
  await prompt.fill('명단을 만들어줘');
  await prompt.press('Home'); await prompt.press('ArrowRight');
  const cursor = await prompt.evaluate((el: HTMLTextAreaElement) => el.selectionStart);
  expect(cursor).toBe(1);
  await prompt.press('Control+a');
  await prompt.pressSequentially('table');
  await expect(prompt).toHaveValue('table');
  await prompt.fill('신청자 명단을 만들어줘');
  const sheet = (await page.locator('.sheet-stage').boundingBox())!;
  const chat = (await page.locator('#ai-dialog').boundingBox())!;
  expect(sheet.x + sheet.width).toBeLessThanOrEqual(chat.x + 1);
  await prompt.press('Enter');
  await expect(page.locator('.ai-plan-summary')).toContainText('3개 셀을 적용했어요');
  await expect(prompt).toBeFocused();
  await selectCell(page, 'B2'); await expect(page.locator('#cell-input')).toHaveValue('3');
  await page.locator('#undo').click();
  await selectCell(page, 'B2'); await expect(page.locator('#cell-input')).toHaveValue('');
  await page.locator('#ai-open').click(); await expect(prompt).toBeFocused();
  await expect(page.locator('#ai-send')).toBeEnabled();
});


test('side chat opens alongside the sheet, resizes and remembers closed preference', async ({ page }, info) => {
  await page.goto('/'); await ready(page);
  const chat = page.locator('#ai-dialog');
  const separator = page.getByRole('separator', {name: 'AI 채팅 너비 조절'});
  await expect(chat).toBeVisible();
  await expect(page.locator('#ai-prompt')).not.toBeFocused();
  const sheet = (await page.locator('.sheet-stage').boundingBox())!;
  const panel = (await chat.boundingBox())!;
  expect(sheet.x + sheet.width).toBeLessThan(panel.x);
  await separator.focus(); await separator.press('Home');
  expect((await chat.boundingBox())!.width).toBe(320);
  await separator.press('ArrowLeft');
  expect((await chat.boundingBox())!.width).toBe(330);
  const handle = (await separator.boundingBox())!;
  await page.mouse.move(handle.x + 4, handle.y + 80); await page.mouse.down();
  await page.mouse.move(1012, handle.y + 80); await page.mouse.up();
  expect((await chat.boundingBox())!.width).toBe(420);
  await page.reload(); await ready(page); await expect(chat).toBeVisible();
  expect((await chat.boundingBox())!.width).toBe(420);
  await page.screenshot({path: info.outputPath('split-desktop.png')});
  await page.getByRole('button', {name:'AI 채팅 닫기'}).click();
  await expect(chat).toBeHidden();
  expect((await page.locator('.sheet-stage').boundingBox())!.width).toBe(1440);
  await page.reload(); await ready(page); await expect(chat).toBeHidden();
  await page.setViewportSize({width:390,height:844});
  await page.locator('#ai-open').click(); await expect(chat).toBeVisible();
  await expect(separator).toBeHidden();
  await expect(page.locator('#ai-send')).toBeInViewport({ratio:1});
});
