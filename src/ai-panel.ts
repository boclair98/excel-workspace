import { applyAIChanges, captureAISelection, resolveAIRange, reviewAIChanges, type AISelection } from "./ai-edit";
import type { Editor } from "./editor";

type TemplateKey = "notice" | "application";
type Field = {
  name: string;
  label: string;
  example: string;
  limit: number;
  optional?: boolean;
  multiline?: boolean;
};

const templates: Record<TemplateKey, { title: string; fields: Field[] }> = {
  notice: {
    title: "안내문",
    fields: [
      { name: "topic", label: "안내 제목", example: "경로당 식사 시간 변경", limit: 100 },
      { name: "audience", label: "안내 대상", example: "경로당 이용자", limit: 120, optional: true },
      { name: "details", label: "안내할 내용", example: "3월 15일부터 오전 11시에 식사를 제공합니다.", limit: 1200, multiline: true },
    ],
  },
  application: {
    title: "신청서",
    fields: [
      { name: "topic", label: "신청 항목", example: "주민센터 스마트폰 교육", limit: 100 },
      { name: "reason", label: "신청 사유", example: "가족과 사진을 주고받는 방법을 배우고 싶습니다.", limit: 1200, multiline: true },
      { name: "applicant", label: "이름", example: "신청하는 사람의 이름", limit: 80, optional: true },
      { name: "recipient", label: "제출처", example: "○○동 주민센터", limit: 120, optional: true },
    ],
  },
};

export function aiPanelMarkup(icon: (name: string) => string) {
  return `<dialog id="ai-dialog" class="ai-dialog" aria-labelledby="ai-title">
    <header class="dialog-heading">
      <span class="ai-mark" aria-hidden="true">${icon("sparkles")}</span>
      <h2 id="ai-title">AI 채팅</h2>
      <button type="button" id="ai-new-chat" hidden>새 대화</button>
      <button type="button" class="icon-btn" data-close="ai-dialog" aria-label="AI 채팅 닫기">${icon("x")}</button>
    </header>
    <form id="ai-form" novalidate>
      <div id="ai-thread" class="ai-thread" tabindex="0" aria-label="문서 작성과 대화">
        <div id="ai-status" class="notice" role="status">AI에 연결하고 있어요.</div>
        <details class="ai-owner" id="ai-owner" hidden>
          <summary>개인 OpenAI 연결 코드</summary>
          <code id="ai-owner-code"></code>
          <button type="button" id="ai-copy-code">코드 복사</button>
        </details>
        <section id="ai-starters" class="ai-starters" aria-labelledby="ai-starters-title">
          <h3 id="ai-starters-title">엑셀에 바로 작성하세요</h3>
          <p>표 만들기, 계산, 내용 정리를 요청하세요. 결과는 시트에 바로 반영됩니다.</p>
          <div class="ai-template-buttons">
            <button type="button" data-template="notice"><strong>안내문</strong><span>일정, 공지, 알림</span></button>
            <button type="button" data-template="application"><strong>신청서</strong><span>교육, 행사, 이용 신청</span></button>
          </div>
        </section>
        <div id="ai-messages" class="ai-messages" role="log" aria-label="AI 대화"></div>
        <fieldset id="ai-template" class="ai-template" aria-describedby="ai-template-hint" hidden>
          <legend id="ai-template-title"></legend>
          <p id="ai-template-hint">필수 항목만 적어도 초안을 받을 수 있어요.</p>
          <div id="ai-template-fields" class="ai-template-fields"></div>
          <button type="button" id="ai-cancel-template">대화로 돌아가기</button>
        </fieldset>
      </div>
      <div class="ai-composer">
        <span id="ai-progress" class="sr-only" role="status"></span>
        <p id="ai-error" class="ai-error" role="alert" hidden></p>
        <div id="ai-scope" class="ai-scope">
          <label><input id="ai-link-sheet" type="checkbox" aria-describedby="ai-range ai-privacy" checked>셀 연결</label>
          <span id="ai-range"></span>
        </div>
        <div class="ai-composer-row">
          <div id="ai-message-entry" class="ai-message-entry">
            <label class="sr-only" for="ai-prompt">메시지</label>
            <textarea id="ai-prompt" rows="2" maxlength="2000" aria-describedby="ai-privacy" placeholder="예: 품목과 수량으로 표를 만들어줘"></textarea>
          </div>
          <button id="ai-send" class="primary" type="submit" disabled>보내기</button>
        </div>
        <p id="ai-privacy" class="ai-privacy">보낸 글과 최근 대화가 AI에 전달됩니다.</p>
      </div>
    </form>
  </dialog><div id="ai-resize" class="ai-resize" role="separator" aria-label="AI 채팅 너비 조절" aria-orientation="vertical" aria-controls="ai-dialog" aria-valuemin="320" aria-valuemax="560" aria-valuenow="320" tabindex="0" hidden title="드래그 또는 방향키로 너비 조절 · 더블 클릭으로 초기화"></div>`;
}

type Status = {
  aiAvailable: boolean;
  managed: boolean;
  signedIn: boolean;
  provider?: "openai" | "claude" | null;
  sharedOpenAI?: boolean;
  openAIUserRef?: string | null;
};
type Message = { role: "user" | "assistant"; content: string };
type Options = { editor: () => Editor | undefined; changed: () => void; toast: (text: string) => void };

function recentHistory(history: Message[]) {
  let start = history.length;
  let length = 0;
  while (start >= 2 && history.length - start < 8) {
    const pairLength = history[start - 2].content.length + history[start - 1].content.length;
    if (length + pairLength > 10000) break;
    length += pairLength;
    start -= 2;
  }
  return history.slice(start);
}

function responseError(status: number) {
  if (status === 429) return "AI 사용 한도에 도달했어요. 잠시 후 다시 보내 주세요.";
  if (status === 504) return "AI 응답 시간이 초과됐어요. 잠시 후 다시 보내 주세요.";
  return `AI 서버 응답을 확인할 수 없어요${status >= 400 ? ` (HTTP ${status})` : ""}. 잠시 후 다시 보내 주세요.`;
}

export function mountAiPanel(options: Options) {
  const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
    document.getElementById(id) as T;
  const dialog = $<HTMLDialogElement>("ai-dialog");
  const form = $<HTMLFormElement>("ai-form");
  const thread = $("ai-thread");
  const prompt = $<HTMLTextAreaElement>("ai-prompt");
  const send = $<HTMLButtonElement>("ai-send");
  const link = $<HTMLInputElement>("ai-link-sheet");
  const reset = $<HTMLButtonElement>("ai-new-chat");
  const cancel = $<HTMLButtonElement>("ai-cancel-template");
  const fields = $("ai-template-fields");
  const messages = $("ai-messages");
  const feedback = $("ai-error");
  const starters = Array.from(form.querySelectorAll<HTMLButtonElement>("[data-template]"));
  const history: Message[] = [];
  let selected: TemplateKey | null = null;
  let status: Status | undefined;
  let pending: AbortController | undefined;
  let opening = 0;
  const resize = $("ai-resize");
  const desktop = matchMedia("(min-width: 1000px)");
  const stored = (key: string) => { try { return localStorage.getItem(key); } catch { return null; } };
  const persist = (key: string, value: string) => { try { localStorage.setItem(key, value); } catch { /* Optional display preference. */ } };
  const widthKey = "excel-workspace-ai-width";
  let preferredWidth = Number(stored(widthKey)) || 0;
  let temporaryClose = false;
  const maxWidth = () => Math.max(320, Math.min(560, innerWidth - 656));
  function syncWidth() {
    if (preferredWidth) document.documentElement.style.setProperty("--ai-panel-width", `${Math.min(maxWidth(), Math.max(320, preferredWidth))}px`);
    resize.setAttribute("aria-valuemax", String(maxWidth()));
    resize.setAttribute("aria-valuenow", String(Math.round(dialog.getBoundingClientRect().width || 320)));
  }
  function setWidth(width: number) {
    preferredWidth = Math.min(maxWidth(), Math.max(320, width));
    persist(widthKey, String(preferredWidth));
    syncWidth();
    window.dispatchEvent(new Event("resize"));
  }
  function resetWidth() {
    preferredWidth = 0;
    persist(widthKey, "0");
    document.documentElement.style.removeProperty("--ai-panel-width");
    window.dispatchEvent(new Event("resize"));
  }
  resize.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    resize.focus();
    resize.setPointerCapture(event.pointerId);
  });
  resize.addEventListener("pointermove", (event) => {
    if (resize.hasPointerCapture(event.pointerId)) setWidth(innerWidth - event.clientX - 8);
  });
  resize.addEventListener("pointerup", (event) => {
    if (resize.hasPointerCapture(event.pointerId)) resize.releasePointerCapture(event.pointerId);
  });
  resize.addEventListener("keydown", (event) => {
    const width = dialog.getBoundingClientRect().width;
    const next = { ArrowLeft: width + 10, ArrowRight: width - 10, Home: 320, End: maxWidth() }[event.key];
    if (next == null) return;
    event.preventDefault(); event.stopPropagation(); setWidth(next);
  });
  resize.addEventListener("dblclick", resetWidth);
  document.addEventListener("ai-panel-width-reset", resetWidth);
  window.addEventListener("resize", syncWidth);
  desktop.addEventListener("change", () => {
    if (!desktop.matches && dialog.open) { temporaryClose = true; dialog.close(); }
  });
  syncWidth();

  function showError(message: string) {
    feedback.textContent = message;
    feedback.hidden = false;
  }

  function updateScope() {
    const provider = status?.provider === "openai" || status?.sharedOpenAI
      ? "OpenAI"
      : status?.provider === "claude" ? "Claude" : "AI";
    $("ai-scope").hidden = Boolean(selected);
    $("ai-range").textContent = "";
    $("ai-privacy").textContent = `보낸 글과 최근 대화가 ${provider}에 전달됩니다.`;
    const editor = options.editor();
    if (!link.checked || !editor) return;
    try {
      const { startRow, startColumn } = editor.range().getRange();
      const range = selected
        ? editor.current().getActiveSheet().getRange(startRow, startColumn)
        : resolveAIRange(editor, prompt.value);
      $("ai-range").textContent = `작업 범위 ${range.getA1Notation()}`;
      $("ai-privacy").textContent = selected
        ? `입력한 내용과 최근 대화가 ${provider}에 전달되며, 초안은 선택한 첫 셀에 작성됩니다.`
        : `보낼 때 이 범위의 셀과 최근 대화가 ${provider}에 전달됩니다.`;
      return { editor, range };
    } catch (error) {
      $("ai-range").textContent = error instanceof Error ? error.message : "작업 범위를 다시 선택해 주세요.";
    }
  }

  function updateComposer() {
    const busy = Boolean(pending);
    send.disabled = busy || !status?.aiAvailable || (status.managed && !status.signedIn);
    send.textContent = busy ? "작성 중…" : selected ? "초안 받기" : "보내기";
    link.disabled = busy;
    $("ai-progress").textContent = busy ? "AI가 답변을 작성하고 있어요." : "";
    $("ai-message-entry").hidden = Boolean(selected);
    $("ai-starters").hidden = Boolean(selected) || history.length > 0;
    $("ai-template").hidden = !selected;
    reset.hidden = !history.length;
    reset.disabled = cancel.disabled = busy;
    for (const button of starters) button.disabled = busy;
    prompt.readOnly = busy;
    for (const field of fields.querySelectorAll<HTMLInputElement | HTMLTextAreaElement>("input,textarea"))
      field.readOnly = busy;
    updateScope();
  }

  function selectTemplate(key: TemplateKey | null) {
    selected = key;
    feedback.hidden = true;
    fields.replaceChildren();
    if (key) {
      const template = templates[key];
      $("ai-template-title").textContent = `${template.title}에 들어갈 내용`;
      for (const field of template.fields) {
        const group = document.createElement("div");
        const label = document.createElement("label");
        const input = document.createElement(field.multiline ? "textarea" : "input");
        input.id = `ai-field-${field.name}`;
        input.name = field.name;
        input.required = !field.optional;
        input.maxLength = field.limit;
        input.placeholder = `예: ${field.example}`;
        if (input instanceof HTMLTextAreaElement) input.rows = 3;
        label.htmlFor = input.id;
        label.textContent = field.label + (field.optional ? " (선택)" : " (필수)");
        group.append(label, input);
        fields.append(group);
      }
    }
    updateComposer();
  }

  function readPrompt() {
    if (!selected) {
      const text = prompt.value.trim();
      if (text) return text;
      prompt.focus();
      throw new Error("보낼 메시지를 적어 주세요.");
    }
    const template = templates[selected];
    const lines = [];
    for (const field of template.fields) {
      const input = $<HTMLInputElement | HTMLTextAreaElement>(`ai-field-${field.name}`);
      const value = input.value.trim();
      if (!value && !field.optional) {
        input.setAttribute("aria-invalid", "true");
        input.focus();
        throw new Error(`‘${field.label}’ 항목을 적어 주세요.`);
      }
      if (value) lines.push(`${field.label}: ${value}`);
    }
    return `${template.title} 초안을 작성해 주세요.\n\n${lines.join("\n")}`;
  }

  function appendMessage(message: Message) {
    const article = document.createElement("article");
    article.className = "ai-message";
    article.dataset.role = message.role;
    const heading = document.createElement("h4");
    heading.textContent = message.role === "user" ? "나" : "AI 답변";
    const body = document.createElement("p");
    body.className = "ai-message-body";
    body.textContent = message.content;
    article.append(heading, body);
    if (message.role === "assistant") {
      const copy = document.createElement("button");
      copy.type = "button";
      copy.className = "ai-copy";
      copy.textContent = "복사";
      copy.addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(message.content);
          copy.textContent = "복사했어요";
        } catch {
          const range = document.createRange();
          range.selectNodeContents(body);
          const selection = window.getSelection();
          selection?.removeAllRanges();
          selection?.addRange(range);
          copy.textContent = "선택한 답변을 복사해 주세요";
        }
      });
      article.append(copy);
    }
    messages.append(article);
    return article;
  }

  function appendChanges(article: HTMLElement, selection: AISelection, changes: ReturnType<typeof reviewAIChanges>) {
    const panel = document.createElement("section");
    panel.className = "ai-plan";
    const summary = document.createElement("p");
    summary.className = "ai-plan-summary";
    summary.setAttribute("role", "status");
    summary.textContent = `${selection.label}에 ${changes.length}개 셀을 적용했어요. 실행 취소로 되돌릴 수 있어요.`;
    const scroll = document.createElement("div");
    scroll.className = "ai-plan-scroll";
    scroll.tabIndex = 0;
    scroll.setAttribute("aria-label", "셀 변경안 목록");
    const table = document.createElement("table");
    const heading = table.createTHead().insertRow();
    for (const label of ["셀", "바꾸기 전", "바꾼 뒤"]) {
      const cell = document.createElement("th");
      cell.scope = "col";
      cell.textContent = label;
      heading.append(cell);
    }
    const body = table.createTBody();
    for (const change of changes) {
      const row = body.insertRow();
      row.insertCell().textContent = change.address;
      row.insertCell().textContent = String(change.cell.before ?? "빈칸");
      row.insertCell().textContent = String(change.formula ?? change.after ?? "빈칸");
    }
    scroll.append(table);
    panel.append(summary, scroll);
    article.append(panel);
  }

  function showStatus() {
    const notice = $("ai-status");
    if (status?.managed && !status.signedIn) {
      const link = document.createElement("a");
      link.href = `https://mcp.coders.kr/sso/login?return_to=${encodeURIComponent(location.origin)}`;
      link.textContent = "로그인하기";
      notice.replaceChildren("AI 채팅을 쓰려면 로그인해 주세요.", link);
    } else {
      notice.textContent = status?.aiAvailable
        ? "AI에 연결됐어요."
        : "AI에 연결할 수 없어요. 잠시 후 채팅을 다시 열어 주세요.";
    }
    $("ai-owner").hidden = !status?.openAIUserRef;
    $("ai-owner-code").textContent = status?.openAIUserRef ?? "";
    updateComposer();
  }

  async function openPanel(focus = true) {
    if (dialog.open) { if (focus) prompt.focus(); return; }
    const previous = document.activeElement as HTMLElement | null;
    const id = ++opening;
    status = undefined;
    $("ai-owner").hidden = true;
    $("ai-status").textContent = "AI에 연결하고 있어요.";
    updateComposer();
    dialog.show();
    document.documentElement.classList.add("ai-panel-open");
    resize.hidden = false;
    $("ai-open").setAttribute("aria-expanded", "true");
    window.dispatchEvent(new Event("resize"));
    if (focus && !selected) prompt.focus();
    else if (!focus) {
      const target = previous && previous !== document.body ? previous : $("editor");
      if (target.id === "editor") target.tabIndex = -1;
      target.focus({ preventScroll: true });
    }
    try {
      const response = await fetch("/api/status", {
        cache: "no-store",
        signal: AbortSignal.timeout(8000),
      });
      if (!response.ok) throw new Error();
      const current = await response.json();
      if (dialog.open && id === opening) {
        status = current;
        showStatus();
      }
    } catch {
      if (dialog.open && id === opening) showStatus();
    }
  }
  $("ai-open").setAttribute("aria-controls", "ai-dialog");
  $("ai-open").setAttribute("aria-expanded", "false");
  $("ai-open").addEventListener("click", () => {
    persist("excel-workspace-ai-open", "true");
    void openPanel();
  });

  dialog.addEventListener("close", () => {
    pending?.abort();
    pending = undefined;
    updateComposer();
    document.documentElement.classList.remove("ai-panel-open");
    resize.hidden = true;
    $("ai-open").setAttribute("aria-expanded", "false");
    if (!temporaryClose) persist("excel-workspace-ai-open", "false");
    temporaryClose = false;
    window.dispatchEvent(new Event("resize"));
    $("ai-open").focus();
  });
  dialog.addEventListener("keydown", (event) => {
    event.stopPropagation();
    if (event.key === "Escape") { event.preventDefault(); dialog.close(); }
  });

  for (const button of starters) {
    button.addEventListener("click", () => {
      selectTemplate(button.dataset.template as TemplateKey);
      fields.querySelector("input")?.focus();
    });
  }
  cancel.addEventListener("click", () => {
    const previous = selected;
    selectTemplate(null);
    starters.find((button) => button.dataset.template === previous)?.focus();
  });
  reset.addEventListener("click", () => {
    history.length = 0;
    messages.replaceChildren();
    prompt.value = "";
    selectTemplate(null);
    thread.scrollTop = 0;
    prompt.focus();
  });
  form.addEventListener("input", (event) => {
    if (event.target instanceof HTMLElement) event.target.removeAttribute("aria-invalid");
    feedback.hidden = true;
    if (event.target === prompt || event.target === link) updateScope();
  });
  prompt.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" || event.shiftKey || event.isComposing || event.keyCode === 229 || send.disabled) return;
    if (matchMedia("(pointer: coarse)").matches && !event.ctrlKey && !event.metaKey) return;
    event.preventDefault();
    form.requestSubmit();
  });
  $("ai-copy-code").addEventListener("click", async () => {
    if (!status?.openAIUserRef) return;
    try {
      await navigator.clipboard.writeText(status.openAIUserRef);
      options.toast("연결 코드를 복사했습니다.");
    } catch {
      options.toast("표시된 연결 코드를 선택해 복사해 주세요.");
    }
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (send.disabled || pending || !dialog.open) return;
    feedback.hidden = true;
    let text: string;
    let selection: AISelection | undefined;
    try {
      text = readPrompt();
      const scope = updateScope();
      if (link.checked && !scope) throw new Error($("ai-range").textContent || "시트를 준비한 뒤 다시 보내 주세요.");
      if (scope) selection = captureAISelection(scope.editor, scope.range);
    } catch (error) {
      showError((error as Error).message);
      return;
    }
    const isTemplate = Boolean(selected);
    const request = new AbortController();
    pending = request;
    updateComposer();
    try {
      const input = { mode: "chat", prompt: text, history: recentHistory(history), context: isTemplate ? undefined : selection?.context };
      const encoder = new TextEncoder();
      let body = JSON.stringify(input);
      while (input.history.length && encoder.encode(body).length > 65536) {
        input.history = input.history.slice(2);
        body = JSON.stringify(input);
      }
      if (encoder.encode(body).length > 65536)
        throw new Error("보낼 내용이 너무 길어요. 셀 범위나 메시지를 줄여 주세요.");
      let response: Response;
      try {
        response = await fetch("/api/ai", {
          method: "POST",
          redirect: "manual",
          headers: { "Content-Type": "application/json" },
          body,
          signal: AbortSignal.any([request.signal, AbortSignal.timeout(45000)]),
        });
      } catch (error) {
        if (error instanceof TypeError) throw new Error("AI 서버에 연결하지 못했어요. 잠시 후 다시 보내 주세요.");
        throw error;
      }
      if (request !== pending || !dialog.open) return;
      if (response.type === "opaqueredirect" || response.status === 401) {
        if (status?.managed) {
          status = { ...status, signedIn: false };
          showStatus();
          $("ai-status").scrollIntoView({ block: "nearest" });
        }
        throw new Error("로그인한 뒤 다시 보내 주세요.");
      }
      const data = response.headers.get("content-type")?.includes("application/json")
        ? await response.json().catch(() => null)
        : null;
      if (request !== pending || !dialog.open) return;
      if (!response.ok)
        throw new Error(typeof data?.error === "string" ? data.error : responseError(response.status));
      if (!data) throw new Error(responseError(response.status));
      if (typeof data.reply !== "string" || !data.reply.trim())
        throw new Error("답변이 비어 있어요. 다시 보내 주세요.");
      const edits = isTemplate && selection
        ? [{ address: selection.cells.keys().next().value, kind: "set", value: data.reply }]
        : data.edits;
      const changes = selection ? reviewAIChanges(selection, edits) : [];
      if (selection && changes.length) {
        const editor = options.editor();
        if (!editor) throw new Error("시트를 준비한 뒤 다시 요청해 주세요.");
        applyAIChanges(editor, selection, changes);
        options.changed();
      }

      const user: Message = { role: "user", content: text };
      const answer: Message = { role: "assistant", content: data.reply };
      history.push(user, answer);
      appendMessage(user);
      const reply = appendMessage(answer);
      if (selection && changes.length) appendChanges(reply, selection, changes);
      if (!isTemplate) prompt.value = "";
      selectTemplate(null);
      reply.scrollIntoView({ block: "start" });
    } catch (error) {
      if (request !== pending || !dialog.open) return;
      showError(
        error instanceof DOMException && error.name === "TimeoutError"
          ? "답변이 늦어지고 있어요. 잠시 후 다시 보내 주세요."
          : error instanceof TypeError
            ? "변경안을 적용하는 중 문제가 생겼어요. 범위를 확인하고 다시 보내 주세요."
            : error instanceof Error ? error.message : "답변을 받지 못했어요.",
      );
    } finally {
      if (request === pending) {
        pending = undefined;
        updateComposer();
        if (dialog.open && (document.activeElement === send || document.activeElement === prompt || document.activeElement === document.body))
          prompt.focus({ preventScroll: true });
      }
    }
  });
  return {
    refreshScope() {
      if (dialog.open && !pending) updateScope();
    },
    openInitial() {
      if (desktop.matches && stored("excel-workspace-ai-open") !== "false") void openPanel(false);
    },
  };
}
