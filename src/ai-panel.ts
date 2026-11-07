import {
  applyAIChanges,
  captureAISelection,
  reviewAIChanges,
  type AISelection,
} from "./ai-edit";
import type { Editor } from "./editor";

export function aiPanelMarkup(icon: (name: string) => string) {
  return `<dialog id="ai-dialog" class="ai-dialog" aria-labelledby="ai-title">
    <div class="dialog-heading"><div><span class="eyebrow">SPREADSHEET ASSISTANT</span><h2 id="ai-title">${icon("sparkles")}AI 도우미</h2></div><button class="icon-btn" data-close="ai-dialog" aria-label="AI 도우미 닫기">${icon("x")}</button></div>
    <div class="ai-content">
      <div class="ai-intro"><span class="ai-mark">${icon("sparkles")}</span><h3>표를 말로 바꿔 보세요.</h3><p>셀을 선택하고 원하는 작업을 적으세요.<br>변경 전후를 확인한 뒤 적용할 수 있어요.</p></div>
      <div id="ai-status" class="notice" role="status">AI 연결 상태를 확인하고 있어요.</div>
      <details class="ai-owner" id="ai-owner" hidden><summary>개인 OpenAI 연결 코드</summary><p>내 Coders 계정에 개인 키를 연결할 때 쓰는 코드예요.</p><code id="ai-owner-code"></code><button id="ai-copy-code" type="button">코드 복사</button></details>
      <form id="ai-form">
        <fieldset class="ai-modes"><legend class="sr-only">AI 작업 선택</legend><label><input type="radio" name="ai-mode" value="question" checked>질문</label><label><input type="radio" name="ai-mode" value="edit">표 바꾸기</label></fieldset>
        <div class="suggestions"><button type="button" data-ai-mode="question" data-prompt="선택한 데이터의 합계를 구하는 수식을 만들어 줘">합계 수식 ${icon("plus")}</button><button type="button" data-ai-mode="question" data-prompt="선택 범위를 한국어로 요약해 줘">데이터 요약 ${icon("plus")}</button><button type="button" data-ai-mode="edit" data-prompt="빈칸을 0으로 채워 줘" hidden>빈칸 채우기 ${icon("plus")}</button><button type="button" data-ai-mode="edit" data-prompt="문자의 앞뒤 공백을 정리해 줘" hidden>공백 정리 ${icon("plus")}</button></div>
        <div id="ai-edit-scope" class="ai-edit-scope" hidden><strong id="ai-range-label"></strong><span>이 범위의 일반 값만 바꿔요.</span></div>
        <label for="ai-prompt">무엇을 도와드릴까요?</label><textarea id="ai-prompt" maxlength="2000" rows="4" required></textarea>
        <label class="checkbox"><input type="checkbox" id="include-cells">선택한 셀 내용도 전송 (최대 200셀)</label>
        <p class="privacy" id="ai-privacy">질문과 동의한 셀만 AI 제공자에게 전송됩니다.</p>
        <button id="ai-send" class="primary wide" type="submit">AI에게 질문하기</button>
      </form>
      <p id="ai-error" class="ai-error" role="alert" hidden></p>
      <section id="ai-result" class="ai-result" aria-labelledby="ai-result-label" hidden>
        <h3 id="ai-result-label" aria-live="polite" tabindex="-1">AI의 제안</h3>
        <p id="ai-explanation"></p>
        <pre id="ai-formula" hidden></pre>
        <div id="ai-changes" class="ai-changes" role="region" aria-label="셀 변경 전후" tabindex="0" hidden>
          <table><thead><tr><th scope="col">셀</th><th scope="col">기존 값</th><th scope="col">새 값</th></tr></thead><tbody id="ai-change-rows"></tbody></table>
        </div>
        <button id="ai-apply" class="primary wide" hidden>선택 셀에 수식 적용</button>
        <p class="privacy" id="ai-result-note">범위와 계산 결과를 확인해 주세요.</p>
      </section>
    </div>
  </dialog>`;
}

type Options = {
  editor: () => Editor;
  changed: () => void;
  toast: (text: string) => void;
  modal: (title: string, body: string, ok?: string) => Promise<boolean>;
};
type Status = {
  aiAvailable: boolean;
  managed: boolean;
  signedIn: boolean;
  provider?: "openai" | "claude" | null;
  openAIUserRef?: string | null;
};
type Preview =
  | {
      mode: "edit";
      selection: AISelection;
      changes: ReturnType<typeof reviewAIChanges>;
    }
  | {
      mode: "question";
      bookId: string;
      sheetId: string;
      address: string;
      previous: string;
      formula: string;
    };

function displayValue(value: string | number | boolean | null) {
  if (value == null) return "빈칸";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return `“${value}”`;
  return String(value);
}

export function mountAiPanel(options: Options) {
  const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
    document.getElementById(id) as T;
  const dialog = $<HTMLDialogElement>("ai-dialog");
  const form = $<HTMLFormElement>("ai-form");
  const prompt = $<HTMLTextAreaElement>("ai-prompt");
  const include = $<HTMLInputElement>("include-cells");
  const send = $<HTMLButtonElement>("ai-send");
  const apply = $<HTMLButtonElement>("ai-apply");
  const feedback = $("ai-error");
  const changesTable = $("ai-changes");
  const changeRows = $("ai-change-rows");
  let status: Status | undefined;
  let preview: Preview | undefined;
  let pending: AbortController | undefined;
  let opening = 0;
  const mode = () =>
    new FormData(form).get("ai-mode") === "edit" ? "edit" : "question";

  function resetResult() {
    preview = undefined;
    feedback.hidden = true;
    $("ai-result").hidden = true;
    apply.hidden = true;
    $("ai-formula").hidden = true;
    changesTable.hidden = true;
    changeRows.replaceChildren();
  }
  function showError(error: unknown) {
    feedback.textContent =
      error instanceof DOMException && error.name === "TimeoutError"
        ? "AI 응답 시간이 초과됐습니다. 선택 범위를 줄여 다시 요청해 주세요."
        : error instanceof TypeError
          ? "AI 서버와 연결하지 못했습니다. 잠시 후 다시 요청해 주세요."
          : error instanceof Error
            ? error.message
            : "AI 요청을 처리하지 못했습니다. 다시 시도해 주세요.";
    feedback.hidden = false;
    feedback.scrollIntoView({ block: "nearest" });
  }
  function updateMode() {
    const editing = mode() === "edit";
    $("ai-edit-scope").hidden = !editing;
    if (editing)
      $("ai-range-label").textContent =
        `선택 범위 ${options.editor().range().getA1Notation()}`;
    for (const button of form.querySelectorAll<HTMLButtonElement>(
      "[data-ai-mode]",
    ))
      button.hidden = button.dataset.aiMode !== mode();
    prompt.placeholder = editing
      ? "예: 빈칸을 0으로 채워 줘"
      : "예: 완료된 행의 금액을 합산하는 수식을 알려 줘";
    send.textContent = pending
      ? "AI 응답을 기다리고 있어요…"
      : editing
        ? "변경안 만들기"
        : "AI에게 질문하기";
    send.disabled = Boolean(pending);
    const provider =
      status?.provider === "openai"
        ? "OpenAI"
        : status?.provider === "claude"
          ? "관리형 Claude"
          : "AI 제공자";
    $("ai-privacy").textContent =
      `${editing ? "변경안을 만들려면 선택한 셀을 전송해야 합니다. " : ""}질문과 동의한 셀만 ${provider}에 전송됩니다.`;
  }
  function showStatus() {
    const notice = $("ai-status");
    if (status?.managed && !status.signedIn) {
      const link = document.createElement("a");
      link.href = `https://mcp.coders.kr/sso/login?return_to=${encodeURIComponent(location.origin)}`;
      link.textContent = "로그인하기 →";
      notice.replaceChildren("AI 사용에는 Coders 로그인이 필요합니다. ", link);
    } else if (!status?.aiAvailable) {
      notice.textContent =
        "AI 연결 전입니다. 엑셀 편집과 저장은 지금도 사용할 수 있어요.";
    } else {
      notice.textContent =
        status.provider === "openai"
          ? "OpenAI 연결 설정됨"
          : "Coders 관리형 AI 연결 설정됨";
    }
    $("ai-owner").hidden = !status?.openAIUserRef;
    $("ai-owner-code").textContent = status?.openAIUserRef ?? "";
    updateMode();
  }
  $("ai-open").addEventListener("click", async () => {
    const id = ++opening;
    status = undefined;
    include.checked = false;
    $("ai-owner").hidden = true;
    $("ai-status").textContent = "AI 연결 상태를 확인하고 있어요.";
    resetResult();
    updateMode();
    dialog.showModal();
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
      if (dialog.open && id === opening)
        $("ai-status").textContent =
          "AI 연결 상태를 확인하지 못했습니다. 잠시 후 다시 열어 주세요.";
    }
  });
  dialog.addEventListener("close", () => {
    pending?.abort();
    pending = undefined;
    resetResult();
  });
  form.querySelector(".ai-modes")!.addEventListener("change", () => {
    pending?.abort();
    pending = undefined;
    include.checked = false;
    resetResult();
    updateMode();
  });
  form
    .querySelectorAll<HTMLButtonElement>("[data-prompt]")
    .forEach((button) => {
      button.onclick = () => {
        prompt.value = button.dataset.prompt ?? "";
        prompt.focus();
      };
    });
  $("ai-copy-code").addEventListener("click", async () => {
    try {
      if (!status?.openAIUserRef) return;
      await navigator.clipboard.writeText(status.openAIUserRef);
      options.toast("연결 코드를 복사했습니다.");
    } catch {
      options.toast("표시된 연결 코드를 선택해 복사해 주세요.");
    }
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (pending) return;
    resetResult();
    let request: AbortController | undefined;
    try {
      if (status?.managed && !status.signedIn)
        throw new Error("위의 로그인하기를 누른 뒤 다시 요청해 주세요.");
      if (!status?.aiAvailable)
        throw new Error(
          "AI가 아직 연결되지 않았습니다. 편집 기능은 계속 사용할 수 있습니다.",
        );
      const action = mode();
      if (action === "edit" && !include.checked) {
        include.focus();
        throw new Error("표를 바꾸려면 선택한 셀 내용 전송에 동의해 주세요.");
      }
      const editor = options.editor();
      const workbook = editor.current();
      const sheet = workbook.getActiveSheet();
      const bounds = editor.range().getRange();
      const cell = sheet.getRange(bounds.startRow, bounds.startColumn);
      const target = {
        bookId: workbook.getId(),
        sheetId: sheet.getSheetId(),
        address: cell.getA1Notation(),
        previous: JSON.stringify(cell.getCellData()),
      };
      const selection = include.checked
        ? captureAISelection(editor)
        : undefined;
      pending = request = new AbortController();
      updateMode();
      const response = await fetch("/api/ai", {
        method: "POST",
        redirect: "manual",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          mode: action,
          prompt: prompt.value,
          context: selection?.context ?? "",
        }),
        signal: AbortSignal.any([request.signal, AbortSignal.timeout(45000)]),
      });
      if (request !== pending || !dialog.open) return;
      if (response.type === "opaqueredirect")
        throw new Error("Coders 로그인 상태를 확인한 뒤 다시 요청해 주세요.");
      if (!response.headers.get("content-type")?.includes("application/json"))
        throw new Error("로그인 또는 AI 사용량을 확인해 주세요.");
      const data = await response.json();
      if (request !== pending || !dialog.open) return;
      if (!response.ok)
        throw new Error(
          typeof data.error === "string"
            ? data.error
            : "AI 요청에 실패했습니다.",
        );
      $("ai-explanation").textContent =
        typeof data.explanation === "string" ? data.explanation : "";
      if (action === "edit" && selection) {
        const changes = reviewAIChanges(selection, data.edits);
        $("ai-result-label").textContent = changes.length
          ? `${selection.label} · ${changes.length}개 셀 변경안`
          : "바뀌는 셀이 없어요";
        for (const change of changes) {
          const row = document.createElement("tr");
          const address = document.createElement("th");
          address.scope = "row";
          const before = document.createElement("td");
          const after = document.createElement("td");
          address.textContent = change.address;
          before.textContent = displayValue(change.cell.before);
          after.textContent = displayValue(change.after);
          row.append(address, before, after);
          changeRows.append(row);
        }
        if (changes.length) preview = { mode: "edit", selection, changes };
        changesTable.hidden = !changes.length;
        apply.textContent = `${changes.length}개 셀에 적용`;
        $("ai-result-note").textContent = changes.length
          ? "적용하면 실제 셀이 바뀝니다. 실행 취소로 한 번에 되돌릴 수 있어요."
          : "요청이나 선택 범위를 바꿔 다시 시도해 보세요.";
      } else {
        const formula =
          typeof data.formula === "string" && data.formula.startsWith("=")
            ? data.formula
            : null;
        if (formula) preview = { mode: "question", ...target, formula };
        $("ai-result-label").textContent = "AI의 제안";
        $("ai-formula").textContent = formula;
        $("ai-formula").hidden = !formula;
        apply.textContent = `${target.address}에 수식 적용`;
        $("ai-result-note").textContent = "범위와 계산 결과를 확인해 주세요.";
      }
      apply.hidden = !preview;
      $("ai-result").hidden = false;
      $("ai-result").scrollIntoView({ block: "nearest" });
    } catch (error) {
      if (!request?.signal.aborted) showError(error);
    } finally {
      if (request && request === pending) {
        pending = undefined;
        updateMode();
      }
    }
  });

  apply.addEventListener("click", async () => {
    const current = preview;
    if (!current || apply.disabled) return;
    apply.disabled = true;
    try {
      const editor = options.editor();
      if (current.mode === "edit") {
        applyAIChanges(editor, current.selection, current.changes);
        $("ai-result-label").textContent =
          `${current.changes.length}개 셀에 적용했습니다`;
      } else {
        const target = () => {
          const workbook = editor.current();
          const sheet = workbook.getActiveSheet();
          if (
            workbook.getId() !== current.bookId ||
            sheet.getSheetId() !== current.sheetId
          )
            throw new Error(
              "질문할 때의 문서와 시트로 돌아간 뒤 적용해 주세요.",
            );
          const cell = sheet.getRange(current.address);
          if (JSON.stringify(cell.getCellData()) !== current.previous)
            throw new Error("셀 내용이 바뀌었습니다. 다시 질문해 주세요.");
          return cell;
        };
        const cell = target();
        if (
          cell.getValue() != null &&
          cell.getValue() !== "" &&
          !(await options.modal(
            "기존 셀 덮어쓰기",
            `<p>${current.address}의 내용을 AI 수식으로 바꿀까요?</p>`,
            "적용",
          ))
        )
          return;
        if (preview !== current || !dialog.open) return;
        target().setValue(current.formula);
        $("ai-result-label").textContent =
          `${current.address}에 수식을 적용했습니다`;
      }
      options.changed();
      preview = undefined;
      apply.hidden = true;
      $("ai-result-note").textContent =
        "편집기의 실행 취소로 적용 전 상태를 복구할 수 있어요.";
      $("ai-result-label").focus();
    } catch (error) {
      preview = undefined;
      apply.hidden = true;
      showError(error);
    } finally {
      apply.disabled = false;
    }
  });
}
