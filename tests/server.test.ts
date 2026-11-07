import test from 'node:test';
import assert from 'node:assert/strict';
// @ts-expect-error Node server is intentionally dependency-free JavaScript.
import { createServer } from '../server.mjs';
// @ts-expect-error Shared server contract is JavaScript.
import { askAI, parseSuggestion, safeFormula, validateRequest } from '../ai.mjs';
test('AI validates inputs, gives honest missing-config status and strips unsafe formula', async () => {
  await assert.rejects(askAI({ prompt: '' }, undefined, {}));
  const result = await askAI({ prompt: '합계 알려줘' }, undefined, {}); assert.equal(result.status, 503);
  assert.equal(parseSuggestion('{"explanation":"합계","formula":"=SUM(A1:A5)"}').formula, '=SUM(A1:A5)');
  assert.equal(parseSuggestion('{"explanation":"bad","formula":"=WEBSERVICE(\"https://bad\")"}').formula, null);
});
test('AI selection schema is enforced server-side and only approved fields are forwarded', () => {
  const context = { sheet: '한글 시트', range: 'A1:B2', values: [[1, '한글'], [true, null]], formulas: [['', ''], ['', '']], extra: 'not shared' };
  const valid = validateRequest({ prompt: ' 요약해 주세요 ', context: JSON.stringify(context) });
  assert.equal(valid.prompt, '요약해 주세요'); assert.equal(JSON.parse(valid.context).extra, undefined);
  assert.throws(() => validateRequest({ prompt: '요약', context: 'null' }));
  assert.throws(() => validateRequest({ prompt: '요약', context: JSON.stringify({ ...context, values: [[{ secret: 'nested data' }]] }) }));
  assert.throws(() => validateRequest({ prompt: '요약', context: JSON.stringify({ ...context, values: [Array(201).fill(1)], formulas: [Array(201).fill('')] }) }));
  assert.throws(() => validateRequest({ prompt: '요약', context: JSON.stringify({ ...context, formulas: [['']] }) }));
});
test('formula review rejects external calls, custom functions and external workbooks', () => {
  for (const f of ['=WEBSERVICE("https://example.invalid")', '=HYPERLINK("url")', '=CALL("x")', '=EXEC("x")', "='[other.xlsx]Sheet1'!A1", "='C:\\folder\\[other.xlsx]Sheet1'!A1", '=CUSTOMFUNC(A1)', '=cmd|x!A1', '=SUM(A1:A2)\n']) assert.equal(safeFormula(f), null, f);
  assert.equal(safeFormula('=SUMIF(A1:A9,"완료",B1:B9)'), '=SUMIF(A1:A9,"완료",B1:B9)');
  assert.equal(safeFormula("='가계부'!A1+SUM(B2:B10)"), "='가계부'!A1+SUM(B2:B10)");
  assert.equal(safeFormula('=IF(A1="SUM(WEBSERVICE)",1,0)'), '=IF(A1="SUM(WEBSERVICE)",1,0)');
  for (const value of ['null', '[]', '42', '"string"']) assert.equal(parseSuggestion(value).formula, null);
  const blocked = parseSuggestion(JSON.stringify({ explanation: '원격 데이터를 사용합니다.', formula: '=WEBSERVICE("url")' })); assert.match(blocked.explanation, /자동 적용에서 제외/);
});
test('AI provider request forwards trusted identity, limits context and parses real response shape', async () => {
  let body: Record<string, unknown> = {}; let headers: Record<string, string> = {};
  const request = async (url: string, init: RequestInit) => { assert.equal(url, 'https://example.invalid/llm/v1/messages'); assert.equal(init.redirect, 'error'); body = JSON.parse(init.body as string); headers = init.headers as Record<string, string>; return new Response(JSON.stringify({ content: [{ type: 'text', text: '{"explanation":"합계 수식입니다.","formula":"=SUM(B2:B10)"}' }] }), { headers: { 'content-type': 'application/json' } }); };
  const result = await askAI({ prompt: 'B2:B10 합계', context: '' }, 'test-user', { ANTHROPIC_API_KEY: 'unit-test-placeholder', ANTHROPIC_BASE_URL: 'https://example.invalid/llm' }, request);
  assert.equal(result.status, 200); assert.equal(result.body.formula, '=SUM(B2:B10)'); assert.equal(headers['X-Coders-User'], 'test-user'); assert.equal(headers['User-Agent'], 'Excel-Workspace/0.3.1'); assert.equal(body.max_tokens, 1200);
});
test('AI timeout, provider failures and invalid response shape return useful safe errors', async () => {
  const config = { ANTHROPIC_API_KEY: 'unit-test-placeholder', ANTHROPIC_BASE_URL: 'https://example.invalid' };
  const input = { prompt: '요약해 줘' };
  assert.equal((await askAI(input, undefined, config, async () => { throw new DOMException('test', 'TimeoutError'); })).status, 504);
  const failure = await askAI(input, undefined, config, async () => { throw new Error('private detail'); }); assert.equal(failure.status, 502); assert.ok(!JSON.stringify(failure).includes('private detail'));
  assert.equal((await askAI(input, undefined, config, async () => new Response('provider private detail', { status: 429 }))).status, 429);
  assert.equal((await askAI(input, undefined, config, async () => new Response('not json'))).status, 502);
  assert.equal((await askAI(input, undefined, config, async () => new Response('{"content":{}}'))).status, 502);
  assert.equal((await askAI(input, undefined, { ...config, ANTHROPIC_BASE_URL: 'ftp://example.invalid' })).status, 503);
});
test('HTTP health, status, invalid requests and no-key path', async () => {
  const server = createServer({}); await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  try {
    assert.equal((await fetch(origin + '/healthz')).status, 200);
    assert.equal((await fetch(origin + '/healthz', { method: 'HEAD' })).status, 200);
    assert.equal((await (await fetch(origin + '/api/status')).json()).aiAvailable, false);
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ prompt: '수식 도움' }) })).status, 503);
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })).status, 400);
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { 'Content-Type': 'application/json', Origin: 'https://evil.invalid' }, body: '{}' })).status, 403);
    assert.equal((await (await fetch(origin + '/api/status', { headers: { 'X-Coders-User': 'client-spoof' } })).json()).signedIn, false);
    assert.equal((await fetch(origin + '/%E0%A4%A')).status, 400);
  } finally { await new Promise<void>(resolve => server.close(resolve)); }
});
test('HTTP native identity requirement, production origins and per-user rate limit', async () => {
  const server = createServer({ CODERS_NATIVE: 'true' }); await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const headers = { 'Content-Type': 'application/json', 'X-Coders-User': 'trusted-gate-user' };
  const body = JSON.stringify({ prompt: '합계' });
  try {
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body })).status, 401);
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { ...headers, Origin: 'http://127.0.0.1:7730' }, body })).status, 403);
    for (let i = 0; i < 6; i++) assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers, body })).status, 503);
    const limited = await fetch(origin + '/api/ai', { method: 'POST', headers, body }); assert.equal(limited.status, 429); assert.ok(Number(limited.headers.get('Retry-After')) > 0);
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { ...headers, 'X-Coders-User': 'different-user' }, body })).status, 503);
  } finally { await new Promise<void>(resolve => server.close(resolve)); }
});
test('HTTP accepts bounded Korean payload and rejects oversized bodies', async () => {
  const server = createServer({}); await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const context = JSON.stringify({ sheet: '시트1', range: 'A1', values: [['한'.repeat(14000)]], formulas: [['']] });
  try {
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ prompt: '한'.repeat(2000), context }) })).status, 503);
    assert.equal((await fetch(origin + '/api/ai', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ prompt: 'a'.repeat(70000) }) })).status, 413);
  } finally { await new Promise<void>(resolve => server.close(resolve)); }
});
