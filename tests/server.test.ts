import test from 'node:test';
import assert from 'node:assert/strict';
// @ts-expect-error Node server is intentionally dependency-free JavaScript.
import { createServer } from '../server.mjs';
// @ts-expect-error Shared server contract is JavaScript.
import { aiProvider, askAI, openAIUserRef, parseEditPlan, parseSuggestion, safeFormula, validateRequest } from '../ai.mjs';
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
  assert.equal(result.status, 200); assert.equal(result.body.formula, '=SUM(B2:B10)'); assert.equal(headers['X-Coders-User'], 'test-user'); assert.equal(headers['User-Agent'], 'Excel-Workspace/0.7.0'); assert.equal(body.max_tokens, 1200);
});
const editContext = JSON.stringify({
  sheet: '품목', range: 'B2:C3',
  values: [[null, '  사과  '], [5, '삭제']],
  formulas: [['', ''], ['=SUM(A1:A2)', '']],
});
const editInput = { mode: 'edit', prompt: '빈칸을 0으로 채워 줘', context: editContext };
const acceptedEdit = { address: 'B2', kind: 'set', value: 0 };
const openAIConfig = { OPENAI_API_KEY: 'unit-test-placeholder' };
const editAnswer = { explanation: 'B2를 채웠습니다.', edits: [acceptedEdit] };

test('OpenAI chat forwards only text turns and returns a plain draft', async () => {
  const history = [
    { role: 'user', content: '경로당 식사 시간 변경 안내문 작성', extra: '제외할 데이터' },
    { role: 'assistant', content: '3월 15일부터 오전 11시에 식사를 제공합니다.' },
  ];
  const prompt = '한 문장으로 줄여 주세요.';
  const reply = '3월 15일부터 경로당 식사는 오전 11시에 시작합니다.';
  const result = await askAI({ mode: 'chat', prompt, history }, undefined, openAIConfig, async (_url: string, init: RequestInit) => {
    const body = JSON.parse(init.body as string);
    assert.equal(body.store, false);
    assert.equal(body.text, undefined);
    assert.equal(body.tools, undefined);
    assert.deepEqual(body.input, [
      { role: 'user', content: history[0].content }, history[1], { role: 'user', content: prompt },
    ]);
    return Response.json({ status: 'completed', output: [{ type: 'message', content: [{ type: 'output_text', text: reply }] }] });
  });
  assert.deepEqual(result, { status: 200, body: { reply, provider: 'openai' } });
  for (const badHistory of [
    [{ role: 'system', content: '무단 지시' }, history[1]],
    [history[0]],
    [{ role: 'user', content: '가'.repeat(2001) }, history[1]],
    [...history, ...history, ...history, ...history, ...history],
  ]) assert.throws(() => validateRequest({ mode: 'chat', prompt, history: badHistory }));
});

test('chat proposes scoped cell updates and safe formulas in the same reply', async () => {
  const input = { mode: 'chat', prompt: 'B2를 12로 채우고 C3에 B2의 합계 수식을 넣어 줘.', context: editContext };
  const edits = [
    { address: 'B2', kind: 'set', value: 12 },
    { address: 'C3', kind: 'formula', value: '=SUM(B2:B2)' },
  ];
  const explanation = '두 셀을 바꿀 수 있어요.';
  const output = (edits: unknown[]) => Response.json({ status: 'completed', output: [{ type: 'message', content: [{ type: 'output_text', text: JSON.stringify({ explanation, edits }) }] }] });
  const result = await askAI(input, undefined, openAIConfig, async (_url: string, init: RequestInit) => {
    const body = JSON.parse(init.body as string);
    assert.equal(body.text.format.name, 'spreadsheet_chat');
    assert.equal(body.text.format.strict, true);
    assert.deepEqual(body.input, [{ role: 'user', content: JSON.stringify({ request: input.prompt, selection: JSON.parse(editContext) }) }]);
    return output(edits);
  });
  assert.deepEqual(result, { status: 200, body: { reply: explanation, edits, provider: 'openai' } });
  for (const edit of [
    { address: 'A1', kind: 'set', value: 1 },
    { address: 'B3', kind: 'clear', value: null },
    { address: 'C3', kind: 'formula', value: '=WEBSERVICE("https://example.invalid")' },
  ]) assert.equal((await askAI(input, undefined, openAIConfig, async () => output([edits[0], edit]))).status, 502);
});

test('managed Claude chat revises a draft in the same thread', async () => {
  const history = [
    { role: 'user', content: '스마트폰 교육 신청서 초안' },
    { role: 'assistant', content: '가족과 사진을 공유하는 방법을 배우고 싶어 신청합니다.' },
  ];
  const prompt = '더 공손하게 써 주세요.';
  const reply = '가족과 사진을 나누는 방법을 배우고자 교육에 신청합니다.';
  const config = { ANTHROPIC_API_KEY: 'unit-test-placeholder', ANTHROPIC_BASE_URL: 'https://example.invalid/llm' };
  const result = await askAI({ mode: 'chat', prompt, history }, 'test-user', config, async (_url: string, init: RequestInit) => {
    assert.deepEqual(JSON.parse(init.body as string).messages, [...history, { role: 'user', content: prompt }]);
    return Response.json({ stop_reason: 'end_turn', content: [{ type: 'text', text: reply }] });
  });
  assert.deepEqual(result, { status: 200, body: { reply, provider: 'claude' } });
});

test('OpenAI Responses requests strict typed edits with consent and no Coders identity', async () => {
  const config = {
    ...openAIConfig,
    CODERS_NATIVE: 'true',
    OPENAI_CODERS_USER_REF: openAIUserRef('authorized-viewer'),
  };
  const result = await askAI(editInput, 'authorized-viewer', config, async (url: string, init: RequestInit) => {
    assert.equal(url, 'https://api.openai.com/v1/responses');
    assert.equal(init.redirect, 'error');
    assert.equal(new Headers(init.headers).get('Authorization'), 'Bearer unit-test-placeholder');
    assert.equal(new Headers(init.headers).has('X-Coders-User'), false);
    const body = JSON.parse(init.body as string);
    assert.equal(body.model, 'gpt-5.4-mini');
    assert.equal(body.store, false);
    assert.equal(body.text.format.strict, true);
    assert.equal(body.text.format.name, 'spreadsheet_edits');
    assert.deepEqual(JSON.parse(body.input), { request: editInput.prompt, selection: JSON.parse(editContext) });
    return Response.json({ status: 'completed', output: [{ type: 'message', content: [{ type: 'output_text', text: JSON.stringify(editAnswer) }] }] });
  });
  assert.equal(result.status, 200);
  assert.deepEqual(result.body, { ...editAnswer, provider: 'openai' });
});

test('managed Claude also returns an edit plan', async () => {
  const config = { ANTHROPIC_API_KEY: 'unit-test-placeholder', ANTHROPIC_BASE_URL: 'https://example.invalid/llm' };
  const result = await askAI(editInput, 'test-user', config, async (_url: string, init: RequestInit) => {
    assert.equal(new Headers(init.headers).get('X-Coders-User'), 'test-user');
    const body = JSON.parse(init.body as string);
    assert.deepEqual(JSON.parse(body.messages[0].content).selection, JSON.parse(editContext));
    return Response.json({ stop_reason: 'end_turn', content: [{ type: 'text', text: JSON.stringify(editAnswer) }] });
  });
  assert.deepEqual(result, { status: 200, body: { ...editAnswer, provider: 'claude' } });
});

test('native mode enables the personal key only for the configured Coders account', () => {
  const user = 'authorized-viewer';
  const config = {
    ...openAIConfig,
    CODERS_NATIVE: 'true',
    HOST: '0.0.0.0',
    OPENAI_CODERS_USER_REF: openAIUserRef(user),
    ANTHROPIC_API_KEY: 'unit-test-placeholder',
    ANTHROPIC_BASE_URL: 'https://example.invalid',
  };
  assert.match(config.OPENAI_CODERS_USER_REF, /^[a-f0-9]{32}$/);
  assert.equal(aiProvider(config, user), 'openai');
  assert.equal(aiProvider(config, 'other-viewer'), 'claude');
  assert.equal(aiProvider(config), 'claude');
  assert.equal(aiProvider({ ...config, OPENAI_CODERS_USER_REF: undefined }, user), 'claude');
  assert.equal(aiProvider({ ...openAIConfig, HOST: '0.0.0.0' }), null);
  assert.equal(aiProvider({ ...openAIConfig, HOST: '127.0.0.1' }), 'openai');
});

test('edit plan rejects missing consent, invalid ranges and entire invalid responses', () => {
  assert.throws(() => validateRequest({ ...editInput, context: '' }), /동의/);
  assert.throws(() => validateRequest({ ...editInput, context: editContext.replace('B2:C3', 'C3:B2') }), /주소/);
  for (const edit of [
    { ...acceptedEdit, address: 'A2' },
    { ...acceptedEdit, address: 'b2' },
    { ...acceptedEdit, address: 'B3' },
    { address: 'C2', kind: 'clear', value: 1 },
    { address: 'C2', kind: 'set', value: { formula: '=1' } },
  ]) {
    assert.throws(() => parseEditPlan(JSON.stringify({ ...editAnswer, edits: [acceptedEdit, edit] }), editContext));
  }
  assert.deepEqual(parseEditPlan(JSON.stringify({ ...editAnswer, edits: [] }), editContext).edits, []);
});

test('incomplete OpenAI responses do not expose a partial edit plan', async () => {
  const result = await askAI(editInput, undefined, openAIConfig, async () =>
    Response.json({ status: 'incomplete', output: [{ type: 'message', content: [{ type: 'output_text', text: JSON.stringify(editAnswer) }] }] }));
  assert.equal(result.status, 502);
  assert.equal(result.body.edits, undefined);
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

test('shared OpenAI serves separate authenticated users and applies the server-wide limit', async (t) => {
  const request = fetch;
  let providerCalls = 0;
  t.mock.method(globalThis, 'fetch', async (url: string, options?: RequestInit) => {
    if (url !== 'https://api.openai.com/v1/responses') return request(url, options);
    providerCalls++;
    assert.equal(new Headers(options?.headers).get('X-Coders-User'), null);
    return Response.json({
      output: [{ type: 'message', content: [{ type: 'output_text', text: JSON.stringify(editAnswer) }] }],
    });
  });
  const server = createServer({
    ...openAIConfig,
    CODERS_NATIVE: 'true',
    PUBLIC_ORIGIN: 'https://excel-workspace.coders.kr',
    OPENAI_SHARED: 'true',
    OPENAI_SHARED_RPM: '2',
    HOST: '0.0.0.0',
  });
  await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const send = (user?: string, browserOrigin = 'https://excel-workspace.coders.kr') => fetch(origin + '/api/ai', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Origin: browserOrigin,
      ...(user ? { 'X-Coders-User': user } : {}),
    },
    body: JSON.stringify(editInput),
  });
  try {
    assert.equal((await send()).status, 401);
    assert.equal((await send('viewer-one', 'https://another.example')).status, 403);
    for (const user of ['viewer-one', 'viewer-two']) {
      const status = await (await fetch(origin + '/api/status', { headers: { 'X-Coders-User': user } })).json();
      assert.equal(status.provider, 'openai');
      assert.equal(status.sharedOpenAI, true);
      assert.equal(status.openAIUserRef, null);
      const response = await send(user);
      assert.equal(response.status, 200);
      assert.deepEqual((await response.json()).edits, editAnswer.edits);
    }
    const limited = await send('viewer-three');
    assert.equal(limited.status, 429);
    assert.ok(Number(limited.headers.get('Retry-After')) > 0);
    assert.equal(providerCalls, 2);
  } finally {
    await new Promise<void>(resolve => server.close(resolve));
  }
});
