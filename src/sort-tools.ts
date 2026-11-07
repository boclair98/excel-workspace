import type { Editor } from "./editor";
import { SheetsSortService } from "@univerjs/sheets-sort";
import {
  assertSortUnchanged,
  createSortConfiguration,
  rangeLabel,
  sortColumns,
  sortRangeFingerprint,
  validateSortRange,
  type SortCriterion,
  type SortConfiguration,
} from "./sort-configuration";
import type { DataBounds } from "./data-operations";

export function openSortTools(
  editor: Editor,
  changed: () => void,
  toast: (message: string) => void,
): HTMLDialogElement {
  const existing = document.getElementById(
    "sort-tools-dialog",
  ) as HTMLDialogElement | null;
  if (existing?.open) {
    existing.focus();
    return existing;
  }
  const trigger =
    document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null;
  const workbookId = editor.current().getId(),
    sheetId = editor.current().getActiveSheet().getSheetId();
  const selected: DataBounds = { ...editor.range().getRange() },
    originalSheet = structuredClone(editor.snapshot().sheets[sheetId]);
  const dialog = document.createElement("dialog");
  dialog.id = "sort-tools-dialog";
  dialog.className = "sort-tools-dialog";
  dialog.setAttribute("aria-labelledby", "sort-tools-title");
  dialog.innerHTML = `<div class="dialog-heading"><div><h2 id="sort-tools-title">표 정렬</h2><p id="sort-sheet-name" class="sort-caption"></p></div><button type="button" class="icon-btn" data-sort-action="close" aria-label="정렬 창 닫기">×</button></div><div class="sort-range-card"><span>선택한 범위</span><strong id="sort-selected-range"></strong><label><input type="checkbox" id="sort-has-header" checked>첫 행은 제목</label></div><p class="sort-description">위쪽 기준부터 순서대로 적용합니다. 제목 행은 제자리에 두고 같은 행의 셀을 함께 이동합니다.</p><div id="sort-criteria" class="sort-criteria"></div><button type="button" id="sort-add-criterion" class="sort-add">+ 정렬 기준 추가</button><section class="sort-summary" aria-labelledby="sort-summary-title"><h3 id="sort-summary-title">정렬 대상 및 기준</h3><p id="sort-target-summary"></p><ol id="sort-order-summary"></ol></section><p class="sort-warning"><strong>표의 모든 열을 선택했는지 확인하세요.</strong><br>선택한 열 안의 셀만 이동합니다. 범위 밖의 열과 행은 정렬하지 않습니다. 숨김·필터 행이 있으면 해제한 뒤 정렬하세요.</p><p id="sort-error" class="sort-error" role="alert" hidden></p><div class="dialog-actions"><button type="button" data-sort-action="close">취소</button><button type="button" class="primary" id="sort-apply">정렬 적용</button></div>`;
  const get = <T extends HTMLElement = HTMLElement>(id: string) =>
    dialog.querySelector<T>(`#${id}`)!;
  const header = get<HTMLInputElement>("sort-has-header"),
    apply = get<HTMLButtonElement>("sort-apply"),
    add = get<HTMLButtonElement>("sort-add-criterion");
  get("sort-sheet-name").textContent = originalSheet.name ?? "현재 시트";
  const criteria: SortCriterion[] = [{ column: 0, ascending: true }];
  let fingerprint = "",
    config: SortConfiguration | null = null,
    setupError = "";
  try {
    validateSortRange(originalSheet, selected);
    fingerprint = sortRangeFingerprint(originalSheet, selected);
    get("sort-selected-range").textContent = rangeLabel(selected);
  } catch (error) {
    setupError =
      error instanceof Error ? error.message : "정렬 범위를 확인해 주세요.";
    get("sort-selected-range").textContent = editor.range().getA1Notation();
  }
  function target() {
    try {
      const workbook = editor.current(),
        sheet = workbook.getActiveSheet();
      if (workbook.getId() === workbookId && sheet.getSheetId() === sheetId)
        return sheet;
    } catch {
      /* Old editor may have been disposed. */
    }
    throw new Error(
      "문서나 시트가 바뀌었습니다. 이 창을 닫고 정렬할 시트에서 다시 열어 주세요.",
    );
  }
  function visibleRowsOnly(range: DataBounds) {
    const sheet = target(),
      model = sheet.getSheet();
    for (let row = range.startRow; row <= range.endRow; row++)
      if (model.getRowFiltered(row) || model.getRowRawVisible(row) === false)
        throw new Error(
          "선택 범위에 숨겨지거나 필터로 가려진 행이 있습니다. 행을 모두 표시한 뒤 정렬해 주세요.",
        );
    if (
      !sheet
        .getInject()
        .get(SheetsSortService)
        .formulaCheck({ unitId: workbookId, subUnitId: sheetId, range })
    )
      throw new Error(
        "선택 범위가 배열 수식의 결과와 겹칩니다. 배열 수식 밖의 범위를 선택해 주세요.",
      );
  }
  function error(message: string) {
    get("sort-error").textContent = message;
    get("sort-error").hidden = false;
  }
  function refreshSummary() {
    config = null;
    apply.disabled = true;
    get("sort-error").hidden = true;
    get("sort-order-summary").replaceChildren();
    try {
      if (setupError) throw new Error(setupError);
      config = createSortConfiguration(
        originalSheet,
        selected,
        header.checked,
        criteria,
      );
      visibleRowsOnly(config.selected);
      get("sort-target-summary").textContent =
        `${rangeLabel(config.dataRange)} · ${config.rows.toLocaleString()}개 행 × ${config.width}개 열${config.hasHeader ? ` · ${selected.startRow + 1}행 제목 유지` : " · 제목 제외 없음"}`;
      for (const criterion of config.criteria) {
        const item = document.createElement("li");
        item.textContent = `${config.columns[criterion.column].label} — ${criterion.ascending ? "오름차순" : "내림차순"}`;
        get("sort-order-summary").append(item);
      }
      apply.textContent = `${config.rows.toLocaleString()}개 행 정렬`;
      apply.disabled = false;
    } catch (e) {
      get("sort-target-summary").textContent =
        "범위와 정렬 기준을 확인해 주세요.";
      error(e instanceof Error ? e.message : "정렬 설정을 확인해 주세요.");
    }
  }
  function renderCriteria(focusIndex?: number) {
    const list = get("sort-criteria");
    list.replaceChildren();
    if (setupError) {
      add.disabled = header.disabled = true;
      refreshSummary();
      return;
    }
    const columns = sortColumns(originalSheet, selected, header.checked);
    criteria.forEach((criterion, index) => {
      const row = document.createElement("div");
      row.className = "sort-criterion";
      const priority = document.createElement("span");
      priority.className = "sort-priority";
      priority.textContent = String(index + 1);
      priority.setAttribute("aria-label", `${index + 1}번째 정렬 기준`);
      const columnLabel = document.createElement("label");
      columnLabel.textContent =
        index === 0 ? "먼저 정렬할 열" : "값이 같으면 다음 열";
      columnLabel.htmlFor = `sort-column-${index}`;
      const column = document.createElement("select");
      column.id = `sort-column-${index}`;
      column.setAttribute("aria-label", `${index + 1}번째 기준 열`);
      for (const info of columns) {
        const option = document.createElement("option");
        option.value = String(info.column);
        option.textContent = info.label;
        option.disabled = criteria.some(
          (c, i) => i !== index && c.column === info.column,
        );
        column.append(option);
      }
      column.value = String(criterion.column);
      columnLabel.append(column);
      const directionLabel = document.createElement("label");
      directionLabel.textContent = "정렬 방향";
      directionLabel.htmlFor = `sort-direction-${index}`;
      const direction = document.createElement("select");
      direction.id = `sort-direction-${index}`;
      direction.setAttribute("aria-label", `${index + 1}번째 정렬 방향`);
      for (const [value, label] of [
        ["asc", "오름차순"],
        ["desc", "내림차순"],
      ]) {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = label;
        direction.append(option);
      }
      direction.value = criterion.ascending ? "asc" : "desc";
      directionLabel.append(direction);
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "icon-btn sort-remove";
      remove.textContent = "×";
      remove.setAttribute("aria-label", `${index + 1}번째 정렬 기준 삭제`);
      remove.disabled = criteria.length === 1;
      column.addEventListener("change", () => {
        criterion.column = Number(column.value);
        renderCriteria(index);
      });
      direction.addEventListener("change", () => {
        criterion.ascending = direction.value === "asc";
        refreshSummary();
      });
      remove.addEventListener("click", () => {
        criteria.splice(index, 1);
        renderCriteria(Math.max(0, index - 1));
      });
      row.append(priority, columnLabel, directionLabel, remove);
      list.append(row);
    });
    add.disabled = criteria.length >= Math.min(3, columns.length);
    refreshSummary();
    if (focusIndex != null)
      get<HTMLSelectElement>(`sort-column-${focusIndex}`).focus();
  }
  header.addEventListener("change", () => renderCriteria());
  add.addEventListener("click", () => {
    const width = selected.endColumn - selected.startColumn + 1;
    const column = Array.from({ length: width }, (_, i) => i).find(
      (i) => !criteria.some((c) => c.column === i),
    );
    if (column != null && criteria.length < 3) {
      criteria.push({ column, ascending: true });
      renderCriteria(criteria.length - 1);
    }
  });
  apply.addEventListener("click", () => {
    if (!config) return;
    apply.disabled = true;
    try {
      const sheet = target(),
        current = editor.snapshot().sheets[sheetId];
      assertSortUnchanged(current, selected, fingerprint);
      const verified = createSortConfiguration(
        current,
        selected,
        header.checked,
        criteria,
      );
      visibleRowsOnly(verified.selected);
      const range = verified.dataRange;
      let mutationObserved = false,
        movedRows = 0;
      // FRange.sort returns the range rather than the command's boolean result. Observe the
      // synchronous reorder mutation so a permission rejection cannot be reported as success.
      const subscription = editor.api.addEvent(
        editor.api.Event.CommandExecuted,
        (event) => {
          if (
            event.id === "sheet.mutation.reorder-range" &&
            event.params?.unitId === workbookId &&
            event.params?.subUnitId === sheetId
          ) {
            mutationObserved = true;
            const order = event.params.order as Record<string, number>;
            movedRows = Object.entries(order ?? {}).filter(
              ([row, source]) => Number(row) !== source,
            ).length;
          }
        },
      );
      try {
        // The installed Univer 1.3.1 facade adds startColumn itself. Criteria stay relative.
        sheet
          .getRange(
            range.startRow,
            range.startColumn,
            verified.rows,
            verified.width,
          )
          .sort(verified.criteria);
      } finally {
        subscription.dispose();
      }
      if (!mutationObserved)
        throw new Error(
          "정렬을 적용하지 못했습니다. 시트 보호와 편집 권한을 확인해 주세요.",
        );
      if (movedRows) changed();
      dialog.close();
      toast(
        movedRows
          ? `${rangeLabel(range)}의 ${verified.rows.toLocaleString()}개 행을 정렬했습니다. 실행 취소로 되돌릴 수 있습니다.`
          : "이미 지정한 순서입니다. 데이터 순서는 바뀌지 않았습니다.",
      );
    } catch (e) {
      error(
        e instanceof Error
          ? e.message
          : "정렬을 적용하지 못했습니다. 범위를 다시 확인해 주세요.",
      );
      apply.disabled = false;
    }
  });
  dialog
    .querySelectorAll<HTMLButtonElement>('[data-sort-action="close"]')
    .forEach((button) =>
      button.addEventListener("click", () => dialog.close()),
    );
  dialog.addEventListener(
    "close",
    () => {
      const toastNode = dialog.querySelector("#toast");
      if (toastNode) document.body.append(toastNode);
      dialog.remove();
      if (trigger?.isConnected) trigger.focus({ preventScroll: true });
    },
    { once: true },
  );
  document.body.append(dialog);
  renderCriteria();
  dialog.showModal();
  if (!setupError) header.focus();
  return dialog;
}
