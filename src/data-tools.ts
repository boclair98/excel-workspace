import type { Editor } from "./editor";
import { validateBook } from "./model";
import {
  createSparseTextPatch,
  findCells,
  planReplacement,
  planTrim,
  type CellMatch,
  type ChangePlan,
  type DataBounds,
  type FindOptions,
} from "./data-operations";

export function openDataTools(
  editor: Editor,
  onChanged: () => void,
  toast: (message: string) => void,
  mode: "find" | "trim" = "find",
): HTMLDialogElement {
  const existing = document.getElementById(
    "data-tools-dialog",
  ) as HTMLDialogElement | null;
  if (existing?.open) {
    existing.querySelector<HTMLInputElement>("input")?.focus();
    return existing;
  }
  const trigger =
    document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null;
  const workbookId = editor.current().getId(),
    sheetId = editor.current().getActiveSheet().getSheetId();
  const capturedSelection: DataBounds = { ...editor.range().getRange() },
    selectionLabel = editor.range().getA1Notation();
  const dialog = document.createElement("dialog");
  dialog.id = "data-tools-dialog";
  dialog.className = "data-tools-dialog";
  dialog.setAttribute("aria-labelledby", "data-tools-title");
  dialog.innerHTML = `<div class="dialog-heading"><div><h2 id="data-tools-title">찾기·바꾸기</h2><p class="data-tools-caption" id="data-tools-sheet"></p></div><button type="button" class="icon-btn" data-action="close" aria-label="데이터 도구 닫기">×</button></div>
    <div class="data-tools-tabs" role="tablist" aria-label="데이터 도구"><button type="button" id="data-tab-find" role="tab" aria-controls="data-find-panel">찾기·바꾸기</button><button type="button" id="data-tab-trim" role="tab" aria-controls="data-trim-panel">앞뒤 공백 정리</button></div>
    <label class="data-tools-field" for="data-scope">적용 범위<select id="data-scope"><option value="sheet">현재 시트 전체</option><option value="selection"></option></select></label>
    <section id="data-find-panel" role="tabpanel" aria-labelledby="data-tab-find"><form id="data-find-form"><label class="data-tools-field" for="data-find-query">찾을 내용<input id="data-find-query" maxlength="500" autocomplete="off" placeholder="찾을 단어나 값을 입력하세요" required></label><div class="data-find-options"><label><input type="checkbox" id="data-case">대소문자 구분</label><label><input type="checkbox" id="data-whole">셀 전체 일치</label><button class="primary" type="submit">찾기</button></div></form><label class="data-tools-field" for="data-replacement">바꿀 내용<input id="data-replacement" maxlength="2000" autocomplete="off" placeholder="비워 두면 찾은 글자를 삭제합니다"></label><div class="data-result-heading"><p id="data-find-count" role="status">찾을 내용을 입력해 주세요.</p><div><button type="button" data-action="previous" aria-label="이전 검색 결과" disabled>이전</button><button type="button" data-action="next" aria-label="다음 검색 결과" disabled>다음</button></div></div><div id="data-results" class="data-results"></div><div class="data-replace-actions"><button type="button" data-action="replace-one" disabled>현재 결과 바꾸기</button><button type="button" data-action="replace-all" disabled>모두 바꾸기</button></div></section>
    <section id="data-trim-panel" role="tabpanel" aria-labelledby="data-tab-trim" hidden><p class="data-tools-description">텍스트의 앞뒤 공백만 지웁니다. 글자 사이의 공백과 줄바꿈은 그대로 유지합니다.</p><button type="button" class="primary" data-action="preview-trim">정리할 셀 미리보기</button></section>
    <p class="data-tools-protection">수식·숫자·논리값·서식 있는 텍스트는 바꾸거나 정리하지 않습니다. 숨겨진 행·열도 선택한 범위에 포함됩니다.</p>
    <section id="data-confirm" class="data-confirm" aria-labelledby="data-confirm-title" hidden><h3 id="data-confirm-title"></h3><p id="data-confirm-caption"></p><div id="data-change-preview" class="data-change-preview"></div><div class="dialog-actions"><button type="button" data-action="cancel-plan">미리보기 취소</button><button type="button" class="primary" data-action="apply-plan">변경 적용</button></div></section>
    <p id="data-error" class="data-tools-error" role="alert" hidden></p><div class="dialog-actions"><button type="button" data-action="close">닫기</button></div>`;
  const get = <T extends HTMLElement = HTMLElement>(id: string) =>
    dialog.querySelector<T>(`#${id}`)!;
  const action = (name: string) =>
    dialog.querySelector<HTMLButtonElement>(`[data-action="${name}"]`)!;
  const query = get<HTMLInputElement>("data-find-query"),
    replacement = get<HTMLInputElement>("data-replacement"),
    scope = get<HTMLSelectElement>("data-scope");
  scope.options[1].textContent = `열 때 선택한 범위 (${selectionLabel})`;
  scope.value = mode === "trim" ? "selection" : "sheet";
  get("data-tools-sheet").textContent = editor
    .current()
    .getActiveSheet()
    .getSheetName();
  let matches: CellMatch[] = [],
    active = -1,
    pending: ChangePlan | null = null,
    currentMode = mode,
    debounce: ReturnType<typeof setTimeout> | undefined;
  function targetSheet() {
    try {
      const workbook = editor.current();
      const sheet = workbook.getActiveSheet();
      if (workbook.getId() === workbookId && sheet.getSheetId() === sheetId)
        return sheet;
    } catch {
      /* Editor may have been disposed while this dialog was open. */
    }
    throw new Error(
      "문서나 시트가 바뀌었습니다. 이 창을 닫고 원하는 시트에서 다시 열어 주세요.",
    );
  }
  function currentData() {
    targetSheet();
    const snapshot = editor.snapshot();
    const sheet = snapshot.sheets[sheetId];
    if (!sheet)
      throw new Error("대상 시트를 찾을 수 없습니다. 창을 다시 열어 주세요.");
    return sheet.cellData;
  }
  const options = (): FindOptions => ({
    caseSensitive: get<HTMLInputElement>("data-case").checked,
    wholeCell: get<HTMLInputElement>("data-whole").checked,
    ...(scope.value === "selection" ? { bounds: capturedSelection } : {}),
  });
  function invalidatePlan() {
    pending = null;
    get("data-confirm").hidden = true;
  }
  function clearError() {
    get("data-error").hidden = true;
  }
  function guard(work: () => void) {
    try {
      clearError();
      work();
    } catch (error) {
      get("data-error").textContent =
        error instanceof Error
          ? error.message
          : "작업을 완료하지 못했습니다. 다시 시도해 주세요.";
      get("data-error").hidden = false;
    }
  }
  function renderResults() {
    get("data-results").replaceChildren();
    const editable = matches.filter((match) => match.editable).length;
    get("data-find-count").textContent = matches.length
      ? `${matches.length.toLocaleString()}개 셀 · ${active + 1}번째 결과 · 바꾸기 가능 ${editable.toLocaleString()}개`
      : query.value
        ? "일치하는 셀이 없습니다."
        : "찾을 내용을 입력해 주세요.";
    action("previous").disabled = action("next").disabled = !matches.length;
    action("replace-one").disabled = !matches[active]?.editable;
    action("replace-all").disabled = !editable;
    const start = Math.max(0, Math.min(active - 5, matches.length - 20));
    matches.slice(start, start + 20).forEach((match, offset) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "data-result-item";
      item.setAttribute("aria-current", String(start + offset === active));
      const address = document.createElement("strong");
      address.textContent = match.address;
      const value = document.createElement("span");
      value.textContent =
        match.text.length > 140 ? match.text.slice(0, 140) + "…" : match.text;
      const state = document.createElement("small");
      state.textContent = match.editable
        ? "텍스트"
        : `${match.protectedReason} · 보호`;
      item.append(address, value, state);
      item.addEventListener("click", () => guard(() => move(start + offset)));
      get("data-results").append(item);
    });
    if (matches.length > 20) {
      const note = document.createElement("p");
      note.className = "data-tools-caption";
      note.textContent = `${start + 1}~${Math.min(start + 20, matches.length)}번째 결과 표시 · 이전/다음으로 모든 결과를 확인할 수 있습니다.`;
      get("data-results").append(note);
    }
  }
  function move(index: number) {
    if (!matches.length) return;
    targetSheet();
    active = (index + matches.length) % matches.length;
    const focus =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : null;
    editor.select(matches[active].address);
    renderResults();
    if (focus?.isConnected && dialog.contains(focus))
      focus.focus({ preventScroll: true });
    else
      dialog
        .querySelector<HTMLButtonElement>(
          '.data-result-item[aria-current="true"]',
        )
        ?.focus({ preventScroll: true });
  }
  function search(navigate = true) {
    clearTimeout(debounce);
    invalidatePlan();
    if (!query.value) {
      matches = [];
      active = -1;
      renderResults();
      return;
    }
    matches = findCells(currentData(), query.value, options());
    active = matches.length
      ? Math.max(0, Math.min(active, matches.length - 1))
      : -1;
    if (navigate && matches.length) move(active);
    else renderResults();
  }
  function preview(plan: ChangePlan, title: string) {
    invalidatePlan();
    if (!plan.changes.length) {
      get("data-error").textContent =
        "바뀌는 텍스트 셀이 없습니다. 찾을 내용과 범위를 확인해 주세요.";
      get("data-error").hidden = false;
      return;
    }
    pending = plan;
    get("data-confirm-title").textContent =
      `${plan.changes.length.toLocaleString()}개 셀 ${title}`;
    get("data-confirm-caption").textContent =
      `내용을 확인한 뒤 적용하세요. ${plan.protectedCount ? `보호된 ${plan.protectedCount.toLocaleString()}개 셀은 제외했습니다. ` : ""}적용 후 한 번의 실행 취소로 되돌릴 수 있습니다.`;
    const list = get("data-change-preview");
    list.replaceChildren();
    for (const change of plan.changes.slice(0, 8)) {
      const row = document.createElement("div");
      row.className = "data-preview-row";
      const address = document.createElement("strong");
      address.textContent = change.address;
      const before = document.createElement("span");
      before.textContent = JSON.stringify(
        change.before.length > 100
          ? change.before.slice(0, 100) + "…"
          : change.before,
      );
      const arrow = document.createElement("span");
      arrow.textContent = "→";
      arrow.setAttribute("aria-label", "변경 후");
      const after = document.createElement("span");
      after.textContent = JSON.stringify(
        change.after.length > 100
          ? change.after.slice(0, 100) + "…"
          : change.after,
      );
      row.append(address, before, arrow, after);
      list.append(row);
    }
    if (plan.changes.length > 8) {
      const p = document.createElement("p");
      p.className = "data-tools-caption";
      p.textContent = `앞의 8개 셀을 표시했습니다. 총 ${plan.changes.length.toLocaleString()}개 셀에 적용합니다.`;
      list.append(p);
    }
    action("apply-plan").textContent =
      `${plan.changes.length.toLocaleString()}개 셀에 적용`;
    get("data-confirm").hidden = false;
    get("data-confirm").scrollIntoView({
      block: "nearest",
      behavior: matchMedia("(prefers-reduced-motion: reduce)").matches
        ? "instant"
        : "smooth",
    });
  }
  function replace(one: boolean) {
    targetSheet();
    const selected = one ? matches.slice(active, active + 1) : matches;
    preview(
      planReplacement(selected, query.value, replacement.value, options()),
      "바꾸기",
    );
  }
  function activate(next: "find" | "trim") {
    currentMode = next;
    invalidatePlan();
    clearError();
    for (const kind of ["find", "trim"] as const) {
      const selected = kind === next;
      get(`data-tab-${kind}`).setAttribute("aria-selected", String(selected));
      get(`data-tab-${kind}`).tabIndex = selected ? 0 : -1;
      get(`data-${kind}-panel`).hidden = !selected;
    }
    get("data-tools-title").textContent =
      next === "find" ? "찾기·바꾸기" : "앞뒤 공백 정리";
  }
  dialog
    .querySelectorAll<HTMLButtonElement>('[data-action="close"]')
    .forEach((button) =>
      button.addEventListener("click", () => dialog.close()),
    );
  get("data-find-form").addEventListener("submit", (event) => {
    event.preventDefault();
    guard(() => {
      active = -1;
      search();
    });
  });
  query.addEventListener("input", () => {
    invalidatePlan();
    clearTimeout(debounce);
    matches = [];
    active = -1;
    renderResults();
    debounce = setTimeout(() => guard(() => search(false)), 180);
  });
  for (const id of ["data-case", "data-whole"])
    get(id).addEventListener("change", () =>
      guard(() => {
        active = -1;
        search(false);
      }),
    );
  scope.addEventListener("change", () =>
    guard(() => {
      invalidatePlan();
      if (currentMode === "find") {
        active = -1;
        search(false);
      }
    }),
  );
  replacement.addEventListener("input", invalidatePlan);
  action("previous").addEventListener("click", () =>
    guard(() => move(active - 1)),
  );
  action("next").addEventListener("click", () => guard(() => move(active + 1)));
  action("replace-one").addEventListener("click", () =>
    guard(() => replace(true)),
  );
  action("replace-all").addEventListener("click", () =>
    guard(() => replace(false)),
  );
  action("preview-trim").addEventListener("click", () =>
    guard(() =>
      preview(planTrim(currentData(), options().bounds), "공백 정리"),
    ),
  );
  action("cancel-plan").addEventListener("click", invalidatePlan);
  action("apply-plan").addEventListener("click", () =>
    guard(() => {
      if (!pending) return;
      const plan = pending,
        sheet = targetSheet(),
        patch = createSparseTextPatch(currentData(), plan.changes),
        b = patch.bounds;
      // Verify the final workbook remains saveable before touching the live editor.
      const candidate = structuredClone(editor.snapshot()),
        data = candidate.sheets[sheetId].cellData!;
      for (const change of plan.changes)
        data[change.row][change.col] = {
          ...data[change.row][change.col],
          ...patch.values[change.row][change.col],
        };
      validateBook(candidate);
      // Univer sparse matrices use absolute worksheet coordinates. One command = one undo item.
      sheet
        .getRange(
          b.startRow,
          b.startColumn,
          b.endRow - b.startRow + 1,
          b.endColumn - b.startColumn + 1,
        )
        .setValues(patch.values);
      const updated = currentData();
      const applied = plan.changes.filter((change) => {
        const cell = updated?.[change.row]?.[change.col];
        return cell?.v === change.after && !cell.f && !cell.si;
      }).length;
      invalidatePlan();
      if (applied) onChanged();
      if (applied !== plan.changes.length)
        throw new Error(
          `${applied.toLocaleString()}개 셀만 반영되었습니다. 시트 보호나 병합 상태를 확인하고 필요하면 실행 취소해 주세요.`,
        );
      toast(
        `${applied.toLocaleString()}개 텍스트 셀을 변경했습니다. 실행 취소로 되돌릴 수 있습니다.`,
      );
      if (currentMode === "find") search(false);
    }),
  );
  for (const kind of ["find", "trim"] as const) {
    get(`data-tab-${kind}`).addEventListener("click", () => activate(kind));
    get(`data-tab-${kind}`).addEventListener("keydown", (event) => {
      if (["ArrowLeft", "ArrowRight"].includes(event.key)) {
        event.preventDefault();
        const next = kind === "find" ? "trim" : "find";
        activate(next);
        get(`data-tab-${next}`).focus();
      }
    });
  }
  dialog.addEventListener(
    "close",
    () => {
      clearTimeout(debounce);
      const toastNode = dialog.querySelector("#toast");
      if (toastNode) document.body.append(toastNode);
      dialog.remove();
      if (trigger?.isConnected) trigger.focus({ preventScroll: true });
    },
    { once: true },
  );
  document.body.append(dialog);
  activate(mode);
  dialog.showModal();
  if (mode === "find") query.focus();
  else {
    action("preview-trim").focus();
    guard(() =>
      preview(planTrim(currentData(), options().bounds), "공백 정리"),
    );
  }
  return dialog;
}
