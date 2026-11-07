import type { Editor } from "./editor";
import { validateBook } from "./model";
import {
  prepareSheetDuplicate,
  requireRemovableSheet,
  validateSheetName,
} from "./sheet-guards";

export const SHEET_COLORS = [
  { name: "초록", value: "#18794e" },
  { name: "파랑", value: "#2563eb" },
  { name: "보라", value: "#7c3aed" },
  { name: "분홍", value: "#db2777" },
  { name: "주황", value: "#c2410c" },
  { name: "회색", value: "#64748b" },
] as const;
let activeDialog: HTMLDialogElement | null = null;

export function openSheetManager(
  editor: Editor,
  changed: () => void,
  selectionChanged: () => void,
): void {
  if (activeDialog?.open) {
    activeDialog.focus();
    return;
  }
  const sourceId = editor.current().getId();
  let selectedId = editor.current().getActiveSheet().getSheetId();
  const previousFocus = document.activeElement;
  const controller = new AbortController();
  const dialog = document.createElement("dialog");
  dialog.id = "sheet-manager-dialog";
  dialog.className = "sheet-manager-dialog";
  dialog.setAttribute("aria-labelledby", "sheet-manager-title");
  dialog.setAttribute("aria-describedby", "sheet-manager-description");
  dialog.innerHTML = `
    <div class="dialog-heading sheet-manager-heading"><div><span class="sheet-manager-eyebrow">통합문서 정리</span><h2 id="sheet-manager-title">시트 관리</h2></div><button type="button" id="sheet-manager-close" class="icon-btn" aria-label="시트 관리 닫기">×</button></div>
    <div class="sheet-manager-body">
      <p id="sheet-manager-description">시트를 선택하고 이름, 순서, 탭 색상을 정리하세요. 변경은 편집기의 실행 취소로 되돌릴 수 있어요.</p>
      <p id="sheet-manager-status" class="sheet-manager-status" role="status" hidden></p>
      <p id="sheet-manager-error" class="sheet-manager-error" role="alert" hidden></p>
      <div class="sheet-manager-layout">
        <section class="sheet-manager-index" aria-labelledby="sheet-manager-list-title"><div class="sheet-manager-list-heading"><h3 id="sheet-manager-list-title">문서의 시트</h3><span id="sheet-manager-count"></span></div><div id="sheet-manager-list" aria-label="전체 시트 목록"></div></section>
        <section class="sheet-manager-detail" aria-labelledby="sheet-manager-selected">
          <div class="sheet-manager-selected-heading"><h3 id="sheet-manager-selected"></h3><span id="sheet-manager-state"></span></div>
          <form id="sheet-manager-rename-form"><label for="sheet-manager-name">시트 이름</label><div class="sheet-manager-name-row"><input id="sheet-manager-name" maxlength="31" required autocomplete="off"><button type="submit" id="sheet-manager-rename">이름 변경</button></div></form>
          <fieldset class="sheet-manager-colors"><legend>탭 색상</legend><div id="sheet-manager-colors"></div></fieldset>
          <div class="sheet-manager-order"><span>시트 순서</span><div><button type="button" id="sheet-manager-left">← 왼쪽으로</button><button type="button" id="sheet-manager-right">오른쪽으로 →</button></div></div>
          <div class="sheet-manager-actions"><button type="button" id="sheet-manager-open" class="primary">시트 열기</button><button type="button" id="sheet-manager-duplicate">시트 복제</button><button type="button" id="sheet-manager-visibility">시트 숨기기</button><button type="button" id="sheet-manager-delete" class="sheet-manager-danger">시트 삭제</button></div>
          <p id="sheet-manager-protection" class="sheet-manager-note"></p>
        </section>
      </div>
      <section id="sheet-manager-confirm" class="sheet-manager-confirm" aria-labelledby="sheet-manager-confirm-title" hidden><h3 id="sheet-manager-confirm-title" tabindex="-1"></h3><p id="sheet-manager-confirm-description"></p><div><button type="button" id="sheet-manager-confirm-cancel">취소</button><button type="button" id="sheet-manager-confirm-apply" class="primary">확인</button></div></section>
      <p class="sheet-manager-note">숨긴 시트의 내용과 수식은 유지되며 이 창에서 다시 표시할 수 있어요. 숨기기는 접근 권한이나 비밀번호 보호가 아닙니다.</p>
    </div>`;
  const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
    dialog.querySelector<T>(`#${id}`)!;
  let pending: {
    kind: "delete" | "hide";
    sheetId: string;
    trigger: HTMLElement;
  } | null = null;
  let busy = false,
    closed = false;
  function current() {
    const workbook = editor.current();
    if (workbook.getId() !== sourceId)
      throw new Error(
        "열린 문서가 바뀌었어요. 시트 관리 창을 닫고 다시 열어 주세요.",
      );
    return workbook;
  }
  function selected(id = selectedId) {
    const sheet = current().getSheetBySheetId(id);
    if (!sheet)
      throw new Error(
        "선택한 시트가 없어졌어요. 시트 관리 창을 다시 열어 주세요.",
      );
    return sheet;
  }
  function clearError() {
    $("sheet-manager-error").hidden = true;
    $("sheet-manager-error").textContent = "";
    $("sheet-manager-status").hidden = true;
  }
  function cancelConfirmation(focus = true) {
    const trigger = pending?.trigger;
    pending = null;
    $("sheet-manager-confirm").hidden = true;
    if (focus && trigger?.isConnected) trigger.focus();
  }
  function render() {
    const workbook = current(),
      sheets = workbook.getSheets();
    if (!sheets.some((sheet) => sheet.getSheetId() === selectedId))
      selectedId = workbook.getActiveSheet().getSheetId();
    const sheet = selected(),
      hidden = sheet.isSheetHidden(),
      index = sheets.findIndex((item) => item.getSheetId() === selectedId);
    const visible = sheets.filter((item) => !item.isSheetHidden()).length;
    const protectedSheet = !hidden && visible <= 1;
    $("sheet-manager-count").textContent = `${sheets.length} / 30`;
    $("sheet-manager-list").replaceChildren();
    for (const item of sheets) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "sheet-manager-item";
      button.dataset.sheet = item.getSheetId();
      button.setAttribute(
        "aria-pressed",
        String(item.getSheetId() === selectedId),
      );
      const swatch = document.createElement("span");
      swatch.className = "sheet-manager-swatch";
      swatch.style.backgroundColor = item.getTabColor() || "#cbd5ce";
      swatch.setAttribute("aria-hidden", "true");
      const name = document.createElement("strong");
      name.textContent = item.getSheetName();
      const state = document.createElement("span");
      state.className = "sheet-manager-badge";
      state.textContent = item.isSheetHidden()
        ? "숨김"
        : item.getSheetId() === workbook.getActiveSheet().getSheetId()
          ? "열린 시트"
          : "";
      button.append(swatch, name, state);
      button.addEventListener(
        "click",
        () => {
          if (busy) return;
          try {
            selected(item.getSheetId());
            selectedId = item.getSheetId();
            clearError();
            cancelConfirmation(false);
            render();
            dialog
              .querySelector<HTMLButtonElement>(
                `.sheet-manager-item[aria-pressed="true"]`,
              )
              ?.focus();
          } catch (error) {
            $("sheet-manager-error").textContent =
              error instanceof Error
                ? error.message
                : "시트 관리 창을 다시 열어 주세요.";
            $("sheet-manager-error").hidden = false;
          }
        },
        { signal: controller.signal },
      );
      $("sheet-manager-list").append(button);
    }
    $("sheet-manager-selected").textContent = sheet.getSheetName();
    $("sheet-manager-state").textContent = hidden
      ? "숨긴 시트"
      : `${index + 1}번째 시트`;
    $<HTMLInputElement>("sheet-manager-name").value = sheet.getSheetName();
    $<HTMLButtonElement>("sheet-manager-left").disabled = index === 0;
    $<HTMLButtonElement>("sheet-manager-right").disabled =
      index === sheets.length - 1;
    $<HTMLButtonElement>("sheet-manager-duplicate").disabled =
      sheets.length >= 30;
    $<HTMLButtonElement>("sheet-manager-open").disabled = hidden;
    $<HTMLButtonElement>("sheet-manager-delete").disabled = protectedSheet;
    $<HTMLButtonElement>("sheet-manager-visibility").disabled = protectedSheet;
    $("sheet-manager-visibility").textContent = hidden
      ? "다시 표시"
      : "시트 숨기기";
    $("sheet-manager-protection").textContent = protectedSheet
      ? "마지막으로 표시된 시트는 삭제하거나 숨길 수 없어요."
      : hidden
        ? "다시 표시하면 편집기 아래의 시트 목록에서 열 수 있어요."
        : "삭제하면 다른 시트에서 참조하는 수식에 #REF! 오류가 생길 수 있어요.";
    for (const button of $(
      "sheet-manager-colors",
    ).querySelectorAll<HTMLButtonElement>("button"))
      button.setAttribute(
        "aria-pressed",
        String(
          (sheet.getTabColor() || "").toLowerCase() === button.dataset.color,
        ),
      );
  }
  async function run(action: () => void | Promise<void>, message?: string) {
    if (busy || closed) return;
    busy = true;
    dialog.setAttribute("aria-busy", "true");
    clearError();
    const focusId = (document.activeElement as HTMLElement | null)?.id;
    try {
      current();
      await action();
      if (!closed) {
        render();
        if (message) {
          const status = $("sheet-manager-status");
          status.textContent = message;
          status.hidden = false;
        }
        if (focusId && !pending) {
          const target = $<HTMLElement>(focusId);
          if (target?.getClientRects().length) target.focus();
          else $("sheet-manager-name").focus();
        }
      }
    } catch (error) {
      if (!closed) {
        $("sheet-manager-error").textContent =
          error instanceof Error
            ? error.message
            : "시트를 변경하지 못했어요. 다시 시도해 주세요.";
        $("sheet-manager-error").hidden = false;
      }
    } finally {
      busy = false;
      dialog.setAttribute("aria-busy", "false");
    }
  }
  function mutation(action: () => void, message: string) {
    void run(() => {
      selected();
      cancelConfirmation(false);
      action();
      changed();
    }, message);
  }
  function confirm(kind: "delete" | "hide", trigger: HTMLElement) {
    void run(() => {
      const sheet = selected();
      requireRemovableSheet(editor.snapshot(), selectedId);
      pending = { kind, sheetId: selectedId, trigger };
      $("sheet-manager-confirm-title").textContent =
        kind === "delete"
          ? `‘${sheet.getSheetName()}’ 시트를 삭제할까요?`
          : `‘${sheet.getSheetName()}’ 시트를 숨길까요?`;
      $("sheet-manager-confirm-description").textContent =
        kind === "delete"
          ? "이 시트의 내용이 제거되고 다른 시트의 수식 참조가 #REF!로 바뀔 수 있어요. 삭제 직후 실행 취소로 복구할 수 있습니다. 중요한 변경 전에는 복구 지점을 남겨 주세요."
          : "이 시트의 내용은 유지되지만 편집기의 시트 목록에서 보이지 않게 됩니다. 이 창에서 언제든 다시 표시할 수 있어요.";
      $("sheet-manager-confirm-apply").textContent =
        kind === "delete" ? "시트 삭제" : "시트 숨기기";
      $("sheet-manager-confirm").hidden = false;
      $("sheet-manager-confirm-title").focus();
      $("sheet-manager-confirm").scrollIntoView({ block: "nearest" });
    });
  }
  for (const color of [...SHEET_COLORS, { name: "색상 초기화", value: "" }]) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.color = color.value;
    button.className = "sheet-manager-color";
    button.setAttribute(
      "aria-label",
      color.name === "색상 초기화" ? color.name : `${color.name} 탭 색상`,
    );
    button.title = color.name;
    button.id = `sheet-manager-color-${color.value ? color.value.slice(1) : "reset"}`;
    if (color.value) {
      const swatch = document.createElement("span");
      swatch.style.backgroundColor = color.value;
      swatch.setAttribute("aria-hidden", "true");
      button.append(swatch);
    } else button.textContent = "초기화";
    button.addEventListener(
      "click",
      () =>
        mutation(() => {
          const sheet = selected();
          sheet.setTabColor(color.value);
          if ((sheet.getTabColor() ?? "").toLowerCase() !== color.value)
            throw new Error("탭 색상을 바꾸지 못했어요.");
        }, "탭 색상을 변경했습니다."),
      { signal: controller.signal },
    );
    $("sheet-manager-colors").append(button);
  }
  const on = (id: string, handler: () => void) =>
    $(id).addEventListener(
      "click",
      () => {
        try {
          handler();
        } catch (error) {
          $("sheet-manager-error").textContent =
            error instanceof Error
              ? error.message
              : "시트 관리 창을 다시 열어 주세요.";
          $("sheet-manager-error").hidden = false;
        }
      },
      { signal: controller.signal },
    );
  $("sheet-manager-rename-form").addEventListener(
    "submit",
    (event) => {
      event.preventDefault();
      mutation(() => {
        const name = validateSheetName(
          editor.snapshot(),
          selectedId,
          $<HTMLInputElement>("sheet-manager-name").value,
        );
        selected().setName(name);
        if (selected().getSheetName() !== name)
          throw new Error("시트 이름을 바꾸지 못했어요.");
      }, "시트 이름을 변경했습니다.");
    },
    { signal: controller.signal },
  );
  on("sheet-manager-duplicate", () =>
    mutation(() => {
      const workbook = current(),
        snapshot = editor.snapshot();
      prepareSheetDuplicate(snapshot, selectedId);
      workbook.duplicateSheet(selected());
      const added = workbook
        .getSheets()
        .find((sheet) => !snapshot.sheetOrder.includes(sheet.getSheetId()));
      if (!added) throw new Error("시트를 복제하지 못했어요.");
      try {
        validateBook(editor.snapshot());
      } catch (error) {
        workbook.undo();
        throw error;
      }
      selectedId = added.getSheetId();
    }, "내용과 서식을 새 시트로 복제했습니다."),
  );
  for (const [id, direction] of [
    ["sheet-manager-left", -1],
    ["sheet-manager-right", 1],
  ] as const)
    on(id, () =>
      mutation(() => {
        const workbook = current(),
          index = workbook
            .getSheets()
            .findIndex((sheet) => sheet.getSheetId() === selectedId);
        const target = index + direction;
        if (target < 0 || target >= workbook.getNumSheets()) return;
        workbook.moveSheet(selected(), target);
        if (workbook.getSheets()[target]?.getSheetId() !== selectedId)
          throw new Error("시트 순서를 바꾸지 못했어요.");
      }, "시트 순서를 변경했습니다."),
    );
  on("sheet-manager-open", () => {
    void run(() => {
      const sheet = selected();
      if (sheet.isSheetHidden())
        throw new Error("숨긴 시트를 먼저 다시 표시해 주세요.");
      current().setActiveSheet(sheet);
      selectionChanged();
      dialog.close();
    });
  });
  on("sheet-manager-visibility", () => {
    if (selected().isSheetHidden())
      mutation(() => {
        const sheet = selected();
        sheet.showSheet();
        if (sheet.isSheetHidden())
          throw new Error("시트를 다시 표시하지 못했어요.");
      }, "시트를 다시 표시했습니다.");
    else confirm("hide", $("sheet-manager-visibility"));
  });
  on("sheet-manager-delete", () =>
    confirm("delete", $("sheet-manager-delete")),
  );
  on("sheet-manager-confirm-cancel", () => {
    if (!busy) cancelConfirmation();
  });
  on("sheet-manager-confirm-apply", () => {
    if (!pending) return;
    const action = pending;
    mutation(
      () => {
        const sheet = selected(action.sheetId);
        requireRemovableSheet(editor.snapshot(), action.sheetId);
        if (action.kind === "hide") {
          sheet.hideSheet();
          if (!sheet.isSheetHidden())
            throw new Error("시트를 숨기지 못했어요.");
        } else if (!current().deleteSheet(sheet))
          throw new Error("시트를 삭제하지 못했어요.");
      },
      action.kind === "hide"
        ? "시트를 숨겼습니다. 시트 관리에서 다시 표시할 수 있어요."
        : "시트를 삭제했습니다. 실행 취소로 복구할 수 있어요.",
    );
  });
  on("sheet-manager-close", () => {
    if (!busy) dialog.close();
  });
  dialog.addEventListener(
    "cancel",
    (event) => {
      if (busy || pending) {
        event.preventDefault();
        if (!busy) cancelConfirmation();
      }
    },
    { signal: controller.signal },
  );
  dialog.addEventListener(
    "close",
    () => {
      closed = true;
      controller.abort();
      const toastNode = dialog.querySelector("#toast");
      if (toastNode) document.body.append(toastNode);
      dialog.remove();
      if (activeDialog === dialog) activeDialog = null;
      if (previousFocus instanceof HTMLElement && previousFocus.isConnected)
        previousFocus.focus({ preventScroll: true });
    },
    { once: true },
  );
  render();
  document.body.append(dialog);
  activeDialog = dialog;
  dialog.showModal();
  dialog
    .querySelector<HTMLButtonElement>(
      '.sheet-manager-item[aria-pressed="true"]',
    )
    ?.focus();
}
