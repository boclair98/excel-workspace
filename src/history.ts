import type { IWorkbookData } from "@univerjs/core";
import {
  createCheckpoint,
  deleteCheckpoint,
  listCheckpoints,
  type Checkpoint,
} from "./storage";

export interface HistoryOptions {
  snapshot: () => IWorkbookData;
  save: () => Promise<void>;
  restore: (book: IWorkbookData) => Promise<void>;
  toast: (message: string) => void;
}

let activeDialog: HTMLDialogElement | null = null;
const dateLabel = (time: number) =>
  new Intl.DateTimeFormat("ko-KR", {
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
  }).format(time);
const messageOf = (error: unknown) => {
  if (error instanceof DOMException && error.name === "QuotaExceededError")
    return "브라우저 저장 공간이 부족합니다. 불필요한 보관본을 정리하거나 파일을 내려받아 주세요.";
  return error instanceof Error
    ? error.message
    : "복구 지점 작업을 완료하지 못했습니다. 다시 시도해 주세요.";
};

/** Opens a local checkpoint manager. Restores always create a separate workbook. */
export function openHistory(options: HistoryOptions): void {
  if (activeDialog?.open) {
    activeDialog.querySelector<HTMLInputElement>("input")?.focus();
    return;
  }
  const source = options.snapshot();
  const sourceId = source.id;
  const previousFocus = document.activeElement;
  const controller = new AbortController();
  const dialog = document.createElement("dialog");
  dialog.id = "history-dialog";
  dialog.className = "history-dialog";
  dialog.setAttribute("aria-labelledby", "history-title");
  dialog.setAttribute("aria-describedby", "history-description");
  // All dynamic workbook/checkpoint content is assigned with textContent below.
  dialog.innerHTML = `
    <div class="dialog-heading history-heading">
      <div><span class="history-eyebrow">이 브라우저에만 보관</span><h2 id="history-title">복구 지점</h2></div>
      <button type="button" class="icon-btn history-close" aria-label="복구 지점 닫기"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" aria-hidden="true"><path d="m6 6 12 12M6 18 18 6"/></svg></button>
    </div>
    <div class="history-body">
      <p id="history-document" class="history-document"></p>
      <p id="history-description">중요한 수정 전에 지금 상태를 남겨 두세요. 복원할 때는 새 문서로 열어 현재 문서를 덮어쓰지 않습니다.</p>
      <form class="history-create" id="history-create">
        <label for="history-name">복구 지점 이름</label>
        <div class="history-create-row"><input id="history-name" type="text" maxlength="80" required autocomplete="off" placeholder="예: 수식 수정 전" autofocus><button type="submit" class="primary" id="history-save">지점 남기기</button></div>
        <p class="history-policy" id="history-policy">문서마다 최근 10개를 보관합니다. 새 지점이 추가되면 가장 오래된 지점부터 정리합니다.</p>
      </form>
      <p id="history-error" class="history-error" role="alert" hidden></p>
      <div class="history-list-heading"><h3>저장한 지점</h3><span id="history-count"></span></div>
      <div id="history-loading" class="history-loading" role="status">복구 지점을 불러오는 중입니다…</div>
      <div id="history-empty" class="history-empty" hidden><strong>아직 남긴 지점이 없어요.</strong><p>현재 문서를 보관한 뒤 첫 복구 지점을 만들어 보세요.</p></div>
      <div id="history-list" class="history-list" role="list" aria-label="저장한 복구 지점"></div>
      <button type="button" id="history-retry" class="history-retry" hidden>다시 불러오기</button>
      <section id="history-confirm" class="history-confirm" aria-labelledby="history-confirm-title" hidden>
        <h3 id="history-confirm-title" tabindex="-1"></h3><p id="history-confirm-description"></p>
        <div class="history-confirm-actions"><button type="button" id="history-confirm-cancel">취소</button><button type="button" id="history-confirm-apply" class="primary">확인</button></div>
      </section>
      <p class="history-footnote">복구 지점도 브라우저 데이터와 함께 지워질 수 있습니다. 중요한 파일은 XLSX 또는 JSON 백업으로 별도 보관하세요. 문서 보관본을 삭제하면 해당 복구 지점도 삭제됩니다.</p>
    </div>`;

  const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
    dialog.querySelector<T>(`#${id}`)!;
  const nameInput = $<HTMLInputElement>("history-name");
  const saveButton = $<HTMLButtonElement>("history-save");
  const form = $<HTMLFormElement>("history-create");
  const list = $("history-list");
  const confirmation = $("history-confirm");
  const errorNode = $("history-error");
  let checkpoints: Checkpoint[] = [];
  let busy = false;
  let loaded = false;
  let action: { kind: "restore" | "delete"; checkpoint: Checkpoint } | null =
    null;
  let actionTrigger: HTMLElement | null = null;
  let loadVersion = 0;
  let closed = false;

  const defaultName = () =>
    `수정 전 · ${new Intl.DateTimeFormat("ko-KR", { month: "numeric", day: "numeric", hour: "numeric", minute: "2-digit" }).format(Date.now())}`;
  nameInput.value = defaultName();
  $("history-document").textContent = source.name;

  function showError(error: unknown) {
    if (closed) return;
    errorNode.textContent = messageOf(error);
    errorNode.hidden = false;
  }
  function clearError() {
    errorNode.textContent = "";
    errorNode.hidden = true;
  }
  function setBusy(value: boolean) {
    busy = value;
    if (closed) return;
    dialog.setAttribute("aria-busy", String(value));
    nameInput.disabled = value;
    saveButton.disabled = value || !loaded;
    list
      .querySelectorAll<HTMLButtonElement>("button")
      .forEach((button) => (button.disabled = value));
    $<HTMLButtonElement>("history-confirm-apply").disabled = value;
    $<HTMLButtonElement>("history-confirm-cancel").disabled = value;
    $<HTMLButtonElement>("history-retry").disabled = value;
    dialog.querySelector<HTMLButtonElement>(".history-close")!.disabled = value;
  }
  function hideConfirmation(focus = true) {
    confirmation.hidden = true;
    action = null;
    if (focus && actionTrigger?.isConnected) actionTrigger.focus();
    actionTrigger = null;
  }
  function confirm(
    kind: "restore" | "delete",
    checkpoint: Checkpoint,
    trigger: HTMLElement,
  ) {
    if (busy) return;
    clearError();
    action = { kind, checkpoint };
    actionTrigger = trigger;
    $("history-confirm-title").textContent =
      kind === "restore" ? "복구 사본을 열까요?" : "이 복구 지점을 삭제할까요?";
    $("history-confirm-description").textContent =
      kind === "restore"
        ? `‘${checkpoint.name}’ 시점의 문서를 ‘복구 사본’으로 엽니다. 현재 문서의 변경 사항을 먼저 보관하며 보관에 실패하면 복원하지 않습니다.`
        : `‘${checkpoint.name}’ 복구 지점을 삭제합니다. 현재 문서와 다른 지점은 유지됩니다. 삭제한 지점은 되돌릴 수 없습니다.`;
    $("history-confirm-apply").textContent =
      kind === "restore" ? "사본으로 복원" : "지점 삭제";
    $("history-confirm-apply").classList.toggle(
      "history-danger",
      kind === "delete",
    );
    confirmation.hidden = false;
    $("history-confirm-title").focus({ preventScroll: true });
    confirmation.scrollIntoView({ block: "nearest" });
  }
  function render() {
    list.replaceChildren();
    $("history-count").textContent = `${checkpoints.length} / 10`;
    $("history-empty").hidden = checkpoints.length > 0;
    const oldest = checkpoints.at(-1);
    $("history-policy").textContent =
      checkpoints.length >= 10 && oldest
        ? `최근 10개가 보관되어 있습니다. 지금 지점을 남기면 가장 오래된 ‘${oldest.name}’ (${dateLabel(oldest.createdAt)}) 지점이 정리됩니다.`
        : "문서마다 최근 10개를 보관합니다. 새 지점이 추가되면 가장 오래된 지점부터 정리합니다.";
    for (const checkpoint of checkpoints) {
      const row = document.createElement("div");
      row.className = "history-row";
      row.setAttribute("role", "listitem");
      const detail = document.createElement("div");
      detail.className = "history-row-detail";
      const title = document.createElement("strong");
      title.textContent = checkpoint.name;
      const time = document.createElement("time");
      time.dateTime = new Date(checkpoint.createdAt).toISOString();
      time.textContent = dateLabel(checkpoint.createdAt);
      detail.append(title, time);
      const buttons = document.createElement("div");
      buttons.className = "history-row-actions";
      const restore = document.createElement("button");
      restore.type = "button";
      restore.textContent = "사본으로 복원";
      restore.className = "history-restore";
      restore.setAttribute("aria-label", `${checkpoint.name} 사본으로 복원`);
      const remove = document.createElement("button");
      remove.type = "button";
      remove.textContent = "삭제";
      remove.className = "history-delete";
      remove.setAttribute("aria-label", `${checkpoint.name} 복구 지점 삭제`);
      restore.addEventListener(
        "click",
        () => confirm("restore", checkpoint, restore),
        { signal: controller.signal },
      );
      remove.addEventListener(
        "click",
        () => confirm("delete", checkpoint, remove),
        { signal: controller.signal },
      );
      buttons.append(restore, remove);
      row.append(detail, buttons);
      list.append(row);
    }
  }
  async function refresh() {
    const version = ++loadVersion;
    $("history-loading").hidden = false;
    $("history-retry").hidden = true;
    try {
      const entries = await listCheckpoints(sourceId);
      if (closed || version !== loadVersion) return;
      checkpoints = entries;
      loaded = true;
      render();
    } catch (error) {
      if (closed || version !== loadVersion) return;
      loaded = false;
      $("history-retry").hidden = false;
      showError(error);
    } finally {
      if (!closed && version === loadVersion) {
        $("history-loading").hidden = true;
        saveButton.disabled = busy || !loaded;
      }
    }
  }
  form.addEventListener(
    "submit",
    (event) => {
      event.preventDefault();
      if (busy || !loaded) return;
      const name = nameInput.value.trim();
      if (!name) {
        nameInput.setCustomValidity("복구 지점 이름을 입력해 주세요.");
        nameInput.reportValidity();
        return;
      }
      void (async () => {
        clearError();
        hideConfirmation(false);
        setBusy(true);
        saveButton.textContent = "보관 중…";
        try {
          if (options.snapshot().id !== sourceId)
            throw new Error(
              "열린 문서가 바뀌었습니다. 복구 지점 창을 닫고 다시 열어 주세요.",
            );
          await options.save();
          const book = options.snapshot();
          if (book.id !== sourceId)
            throw new Error(
              "열린 문서가 바뀌었습니다. 복구 지점 창을 다시 열어 주세요.",
            );
          await createCheckpoint(book, name);
          options.toast("현재 상태를 복구 지점으로 남겼습니다.");
          if (!closed) {
            nameInput.value = defaultName();
            await refresh();
          }
        } catch (error) {
          showError(error);
        } finally {
          setBusy(false);
          if (!closed) {
            saveButton.textContent = "지점 남기기";
            nameInput.focus();
          }
        }
      })();
    },
    { signal: controller.signal },
  );
  nameInput.addEventListener("input", () => nameInput.setCustomValidity(""), {
    signal: controller.signal,
  });
  $("history-confirm-cancel").addEventListener(
    "click",
    () => hideConfirmation(),
    { signal: controller.signal },
  );
  $("history-confirm-apply").addEventListener(
    "click",
    () => {
      if (busy || !action) return;
      const pending = action;
      void (async () => {
        clearError();
        setBusy(true);
        $("history-confirm-apply").textContent =
          pending.kind === "restore" ? "사본을 여는 중…" : "삭제 중…";
        try {
          if (pending.kind === "restore") {
            const copy = structuredClone(pending.checkpoint.book);
            copy.id = crypto.randomUUID();
            copy.name = `${copy.name.slice(0, 92)} · 복구 사본`;
            await options.restore(copy);
            options.toast(
              "복구 지점을 새 사본으로 열었습니다. 원래 문서는 유지됩니다.",
            );
            if (!closed) dialog.close();
          } else {
            await deleteCheckpoint(pending.checkpoint.id);
            options.toast("복구 지점을 삭제했습니다.");
            if (!closed) {
              hideConfirmation(false);
              await refresh();
            }
          }
        } catch (error) {
          showError(error);
        } finally {
          setBusy(false);
          if (!closed && action)
            $("history-confirm-apply").textContent =
              action.kind === "restore" ? "사본으로 복원" : "지점 삭제";
          else if (!closed && pending.kind === "delete") nameInput.focus();
        }
      })();
    },
    { signal: controller.signal },
  );
  $("history-retry").addEventListener(
    "click",
    () => {
      if (!busy) {
        clearError();
        void refresh();
      }
    },
    { signal: controller.signal },
  );
  dialog
    .querySelector<HTMLButtonElement>(".history-close")!
    .addEventListener("click", () => dialog.close(), {
      signal: controller.signal,
    });
  dialog.addEventListener(
    "cancel",
    (event) => {
      // Escape first dismisses an action confirmation, then closes the manager.
      if (busy) {
        event.preventDefault();
        return;
      }
      if (action && !busy) {
        event.preventDefault();
        hideConfirmation();
      }
    },
    { signal: controller.signal },
  );
  dialog.addEventListener(
    "close",
    () => {
      closed = true;
      loadVersion++;
      controller.abort();
      const toast = dialog.querySelector("#toast");
      if (toast) document.body.append(toast);
      dialog.remove();
      if (activeDialog === dialog) activeDialog = null;
      if (previousFocus instanceof HTMLElement && previousFocus.isConnected)
        previousFocus.focus({ preventScroll: true });
    },
    { once: true },
  );
  document.body.append(dialog);
  activeDialog = dialog;
  saveButton.disabled = true;
  dialog.showModal();
  void refresh();
}
