import "@fontsource-variable/noto-sans-kr";
import "./style.css";
import {
  createIcons,
  Table2,
  FolderOpen,
  Plus,
  Download,
  Save,
  PanelLeft,
  Sparkles,
  Undo2,
  Redo2,
  Bold,
  Italic,
  HelpCircle,
  X,
  Search,
  ArrowUpAZ,
  Filter,
  FileSpreadsheet,
  ChevronDown,
  ArrowUpRight,
  Check,
  ShieldCheck,
  Trash2,
  Copy,
  Settings2,
  History,
  Pencil,
  Layers,
} from "lucide";
import { newBook, validateBook, LIMITS, cellText } from "./model";
import {
  listBooks,
  saveBook,
  deleteBook,
  StorageConflictError,
  type SavedBook,
} from "./storage";
import type { IWorkbookData } from "@univerjs/core";
import type { Editor } from "./editor";
import { mountRibbon } from "./ribbon";
import { openHistory } from "./history";
import { openDataTools } from "./data-tools";

const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
  document.getElementById(id) as T;
const esc = (s: string) =>
  s.replace(
    /[&<>"']/g,
    (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        c
      ]!,
  );
const icon = (name: string) =>
  `<i data-lucide="${name}" aria-hidden="true"></i>`;
function icons() {
  createIcons({
    icons: {
      Table2,
      FolderOpen,
      Plus,
      Download,
      Save,
      PanelLeft,
      Sparkles,
      Undo2,
      Redo2,
      Bold,
      Italic,
      HelpCircle,
      X,
      Search,
      ArrowUpAZ,
      Filter,
      FileSpreadsheet,
      ChevronDown,
      ArrowUpRight,
      Check,
      ShieldCheck,
      Trash2,
      Copy,
      Settings2,
      History,
      Pencil,
      Layers,
    },
    attrs: { "stroke-width": 1.7 },
  });
}
$("app").innerHTML = `
<header class="topbar"><div class="brand">${icon("table-2")}<h1>오픈소스 엑셀</h1><span class="beta">BETA</span></div><div class="document-heading"><button id="rename" title="통합문서 이름 변경"><span id="document-name">새 통합문서</span>${icon("chevron-down")}</button><span id="save-state" role="status">이 브라우저에 자동 보관</span></div><button class="icon-btn" id="help" aria-label="이용 안내">${icon("help-circle")}</button></header>
<nav class="filebar" aria-label="파일 도구"><div class="file-actions"><button id="library-toggle" class="icon-btn" aria-label="내 통합문서 열기">${icon("panel-left")}</button><span class="separator"></span><button id="new">${icon("plus")}새 문서</button><button id="open">${icon("folder-open")}열기</button><button id="save">${icon("save")}보관</button><button id="export" class="primary">${icon("download")}XLSX 저장</button><button id="csv" class="small-action">CSV</button></div><button id="ai-open" class="ai-trigger">${icon("sparkles")}<span>AI 도우미</span></button></nav>
<div class="workspace"><aside id="sidebar" class="sidebar" aria-label="내 통합문서"><div class="side-top"><h2>내 통합문서</h2><span id="book-count">0</span></div><label class="search">${icon("search")}<input id="search" type="search" placeholder="이름으로 찾기" aria-label="보관한 통합문서 검색"></label><div class="library" id="library"></div><div class="side-bottom"><button id="example">${icon("file-spreadsheet")}예시 가계부 열기</button><button id="history">${icon("history")}복구 지점</button><button id="backup">${icon("download")}현재 문서 백업</button><a href="https://github.com/boclair98/excel-workspace" target="_blank" rel="noopener">소스 코드 보기 ${icon("arrow-up-right")}</a><p>${icon("shield-check")}문서는 이 브라우저에 보관돼요.<br>중요한 파일은 내려받아 주세요.</p></div></aside>
<main class="editor-area" id="editor-area" aria-label="스프레드시트 편집기"><div class="quickbar" role="toolbar" aria-label="간편 편집"><button id="undo" class="icon-btn" aria-label="실행 취소">${icon("undo-2")}</button><button id="redo" class="icon-btn" aria-label="다시 실행">${icon("redo-2")}</button><span class="separator"></span><button id="bold" class="icon-btn" aria-label="굵게">${icon("bold")}</button><button id="italic" class="icon-btn" aria-label="기울임">${icon("italic")}</button><label class="sr-only" for="font-size">글자 크기</label><select id="font-size"><option>10</option><option selected>11</option><option>12</option><option>14</option><option>16</option><option>20</option><option>24</option></select><label class="sr-only" for="number-format">표시 형식</label><select id="number-format"><option value="General">일반</option><option value="#,##0">숫자</option><option value="₩#,##0">원화</option><option value="0.0%">백분율</option><option value="yyyy-mm-dd">날짜</option></select><span class="separator"></span><button id="filter" title="선택한 범위에 필터 설정">${icon("filter")}필터</button><button id="sort" title="제목 행과 기준 열을 지정해 선택 범위 정렬">${icon("arrow-up-a-z")}정렬</button><button id="full-tools" aria-pressed="false">${icon("settings-2")}전체 도구</button></div>
<form class="formula" id="cell-form"><label class="sr-only" for="address">셀 주소 또는 범위</label><input id="address" value="A1" spellcheck="false" aria-label="셀 주소 또는 범위"><span class="fx" aria-hidden="true">ƒx</span><label class="sr-only" for="cell-input">셀 내용 또는 수식</label><textarea id="cell-input" rows="1" placeholder="셀 내용 또는 =SUM(A1:A10)" autocomplete="off" aria-description="Enter로 입력, Alt+Enter 또는 Shift+Enter로 줄 바꿈"></textarea><button type="submit" class="icon-btn" aria-label="셀에 입력">${icon("check")}</button></form>
<output id="formula-help" class="formula-help" hidden></output><div id="save-conflict" class="save-conflict" role="alert" hidden><div><strong>다른 탭에서 이 문서가 변경됐어요.</strong><span>내 변경을 사본으로 보관하거나 최신 보관본을 열어 주세요.</span></div><div class="conflict-actions"><button id="conflict-copy" class="primary">내 변경 사본 보관</button><button id="conflict-reload">최신 보관본 열기</button></div></div><div class="sheet-stage"><div id="editor"></div><div id="loading" class="loading" role="status"><span class="spinner"></span><strong>통합문서를 준비하고 있어요</strong><span>표와 수식을 불러오는 중입니다.</span></div><div class="start-hint" id="start-hint"><strong>빈 셀에서 시작하세요.</strong><span>더블 클릭해 입력하거나, 엑셀 파일을 열어 이어서 작업하세요.</span><button id="hint-example">예시 가계부로 둘러보기 →</button><button id="hint-close" class="icon-btn" aria-label="시작 안내 닫기">${icon("x")}</button></div></div>
<nav class="sheetbar" aria-label="시트 탐색"><div id="sheet-tabs" class="sheet-tabs" role="tablist" aria-label="통합문서 시트"></div><label class="sr-only" for="sheet-picker">현재 시트</label><select id="sheet-picker" aria-label="현재 시트"></select><button id="sheet-add" class="icon-btn" aria-label="시트 추가" title="시트 추가">${icon("plus")}</button><button id="sheet-rename" class="icon-btn" aria-label="시트 이름 변경" title="시트 이름 변경">${icon("pencil")}</button><button id="sheet-manage" class="icon-btn" aria-label="시트 관리" title="시트 관리">${icon("layers")}</button><label class="sr-only" for="sheet-zoom">시트 확대 배율</label><select id="sheet-zoom" aria-label="시트 확대 배율"><option value="0.75">75%</option><option value="1" selected>100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select></nav><footer class="statusbar"><span id="selection-status">A1 선택</span><span>Univer 기반 · 파일은 기기에 저장</span></footer></main></div>
<div id="toast" role="status" class="toast" hidden></div>
<input type="file" id="file" accept=".xlsx,.csv,.json" hidden>
<dialog id="library-dialog" class="library-dialog" aria-label="내 통합문서"><button class="dialog-close icon-btn" data-close="library-dialog" aria-label="보관함 닫기">${icon("x")}</button><div id="library-host"></div></dialog>
<dialog id="modal" aria-labelledby="modal-title"><form method="dialog"><div class="dialog-heading"><h2 id="modal-title"></h2><button class="icon-btn" value="cancel" aria-label="대화상자 닫기">${icon("x")}</button></div><div id="modal-body"></div><div class="dialog-actions"><button value="cancel">취소</button><button id="modal-ok" class="primary" value="ok">확인</button></div></form></dialog>
<dialog id="ai-dialog" class="ai-dialog" aria-labelledby="ai-title"><div class="dialog-heading"><div><span class="eyebrow">SPREADSHEET ASSISTANT</span><h2 id="ai-title">${icon("sparkles")}AI 도우미</h2></div><button class="icon-btn" data-close="ai-dialog" aria-label="AI 도우미 닫기">${icon("x")}</button></div><div class="ai-content"><div class="ai-intro"><span class="ai-mark">${icon("sparkles")}</span><h3>수식이 막힐 땐,<br>말로 물어보세요.</h3><p>수식 만들기부터 표 해석까지.<br>제안을 확인하고 직접 적용할 수 있어요.</p></div><div id="ai-status" class="notice">AI 연결 상태를 확인하고 있어요.</div><div class="suggestions"><button data-prompt="B2부터 B10까지 합계를 구하는 수식을 만들어 줘">합계 수식 만들기 ${icon("plus")}</button><button data-prompt="선택 범위의 값을 바탕으로 주요 내용을 한국어로 요약해 줘. 데이터가 없으면 요청해 줘.">선택한 데이터 요약하기 ${icon("plus")}</button><button data-prompt="IF와 SUMIF 함수의 차이를 짧은 예시와 함께 설명해 줘">함수 사용법 알아보기 ${icon("plus")}</button></div><form id="ai-form"><label for="ai-prompt">무엇을 도와드릴까요?</label><textarea id="ai-prompt" maxlength="2000" rows="4" placeholder="예: B열이 ‘완료’인 행의 C열 금액을 합산해 줘" required></textarea><label class="checkbox"><input type="checkbox" id="include-cells">선택한 셀 내용도 전송 (최대 200셀)</label><p class="privacy">질문과 체크한 셀만 AI 제공자에게 전송됩니다. 개인정보는 제외해 주세요. Coders 로그인과 사용량 한도가 적용됩니다.</p><button id="ai-send" class="primary wide" type="submit">${icon("sparkles")}AI에게 질문하기</button></form><section id="ai-result" class="ai-result" aria-live="polite" hidden><span class="eyebrow">AI의 제안</span><p id="ai-explanation"></p><pre id="ai-formula" hidden></pre><button id="ai-apply" class="primary wide" hidden>선택 셀에 수식 적용</button><p class="privacy">AI 답변은 틀릴 수 있어요. 범위와 계산 결과를 확인해 주세요.</p></section></div></dialog>`;
icons();
$("ai-open").before($("export"));

let editor: Editor,
  books: SavedBook[] = [],
  dirty = false,
  revision = 0,
  autoTimer: ReturnType<typeof setTimeout>,
  toastTimer: ReturnType<typeof setTimeout>,
  saving: Promise<void> = Promise.resolve(),
  loading = false;
const knownVersions = new Map<string, number | null>();
let conflictBookId: string | null = null;
let sheetSignature = "";
let aiTarget:
    | { book: string; sheet: string; address: string; previous: string }
    | undefined,
  formula: string | null = null;
const toast = (text: string, error = false) => {
  const host =
    document.activeElement?.closest("dialog[open]") ??
    Array.from(document.querySelectorAll("dialog[open]")).at(-1) ??
    document.body;
  host.append($("toast"));
  $("toast").textContent = text;
  $("toast").className = "toast" + (error ? " error" : "");
  $("toast").hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => ($("toast").hidden = true), 6000);
};
const fail = (e: unknown) =>
  toast(
    e instanceof Error
      ? e.message
      : "작업을 완료하지 못했습니다. 다시 시도해 주세요.",
    true,
  );
const run = (fn: () => unknown) => async () => {
  try {
    await fn();
  } catch (e) {
    fail(e);
  }
};
const bind = (id: string, fn: () => unknown) =>
  $(id).addEventListener("click", run(fn));
mountRibbon({ editor: () => editor, changed: onChange, modal, toast });
document.addEventListener("data-tools", (event) => {
  run(() =>
    openDataTools(
      editor,
      onChange,
      toast,
      (event as CustomEvent).detail === "trim" ? "trim" : "find",
    ),
  )();
});
document.addEventListener("add-sheet", run(addSheet));
document.addEventListener(
  "duplicate-workbook",
  run(async () => {
    const book = editor.snapshot();
    book.id = crypto.randomUUID();
    book.name = book.name.slice(0, 95) + " 사본";
    await openBook(book);
    await saveCurrent(false);
    toast("현재 문서를 별도 사본으로 보관했습니다.");
  }),
);
function onChange() {
  if (!editor || loading) return;
  dirty = true;
  revision++;
  updateSheets();
  if (conflictBookId === editor.current().getId()) {
    $("save-state").textContent = "저장 충돌 · 사본 보관이 필요해요";
    return;
  }
  $("save-state").textContent = "변경 사항 보관 중…";
  $("start-hint").hidden = true;
  clearTimeout(autoTimer);
  autoTimer = setTimeout(() => run(() => saveCurrent(false))(), 1200);
}
function onSelection() {
  if (!editor) return;
  const r = editor.range();
  const bounds = r.getRange();
  const count =
    (bounds.endRow - bounds.startRow + 1) *
    (bounds.endColumn - bounds.startColumn + 1);
  if (document.activeElement !== $("address"))
    $("address").setAttribute("value", r.getA1Notation());
  if (document.activeElement !== $("address"))
    ($("address") as HTMLInputElement).value = r.getA1Notation();
  if (document.activeElement !== $("cell-input"))
    ($("cell-input") as HTMLTextAreaElement).value =
      r.getCellData()?.f || String(r.getCellData()?.v ?? "");
  resizeFormulaInput();
  let message = `${r.getA1Notation()} · ${count.toLocaleString()}개 셀`;
  if (count <= 10000) {
    const data = editor.snapshot().sheets[r.getSheetId()].cellData ?? {};
    const nums: number[] = [];
    for (let row = bounds.startRow; row <= bounds.endRow; row++) {
      for (
        let column = bounds.startColumn;
        column <= bounds.endColumn;
        column++
      ) {
        const value = data[row]?.[column]?.v;
        if (typeof value === "number" && Number.isFinite(value))
          nums.push(value);
      }
    }
    if (nums.length) {
      const sum = nums.reduce((a, b) => a + b, 0);
      const format = (n: number) =>
        n.toLocaleString("ko-KR", { maximumFractionDigits: 6 });
      message += ` · 숫자 ${nums.length.toLocaleString()}개 · 합계 ${format(sum)} · 평균 ${format(sum / nums.length)}`;
    }
  }
  $("selection-status").textContent = message;
  updateSheets();
}
async function refreshLibrary() {
  books = await listBooks();
  $("book-count").textContent = String(books.length);
  renderLibrary();
}
function updateSheets() {
  if (!editor) return;
  const current = editor.current();
  const active = current.getActiveSheet();
  const sheets = current.getSheets().filter((sheet) => !sheet.isSheetHidden());
  const signature =
    JSON.stringify(
      sheets.map((sheet) => [
        sheet.getSheetId(),
        sheet.getSheetName(),
        sheet.getTabColor(),
      ]),
    ) + active.getSheetId();
  if (signature !== sheetSignature) {
    sheetSignature = signature;
    $("sheet-tabs").innerHTML = sheets
      .map((sheet) => {
        const selected = sheet.getSheetId() === active.getSheetId();
        const rawColor = sheet.getTabColor() ?? "";
        const color = /^#[\da-f]{6}$/i.test(rawColor)
          ? rawColor
          : "transparent";
        return `<button role="tab" aria-controls="editor" data-sheet="${esc(sheet.getSheetId())}" aria-selected="${selected}" tabindex="${selected ? 0 : -1}" style="--sheet-color:${color}">${esc(sheet.getSheetName())}</button>`;
      })
      .join("");
    $("sheet-picker").innerHTML = sheets
      .map(
        (sheet) =>
          `<option value="${esc(sheet.getSheetId())}">${esc(sheet.getSheetName())}</option>`,
      )
      .join("");
    ($("sheet-picker") as HTMLSelectElement).value = active.getSheetId();
    $("sheet-tabs")
      .querySelector('[aria-selected="true"]')
      ?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  for (const id of ["sheet-zoom", "zoom"]) {
    const select = $(id) as HTMLSelectElement;
    const zoom = String(active.getZoom());
    for (const option of select.querySelectorAll("[data-custom]"))
      option.remove();
    if (!Array.from(select.options).some((option) => option.value === zoom)) {
      const option = new Option(`${Math.round(active.getZoom() * 100)}%`, zoom);
      option.dataset.custom = "true";
      select.add(option);
    }
    select.value = zoom;
  }
}
function activateSheet(id: string) {
  editor.current().setActiveSheet(id);
  onSelection();
}
function addSheet() {
  if (editor.current().getSheets().length >= LIMITS.sheets)
    throw new Error(
      `통합문서에는 최대 ${LIMITS.sheets}개 시트를 만들 수 있어요.`,
    );
  const sheet = editor.current().insertSheet();
  editor.current().setActiveSheet(sheet);
  onChange();
  onSelection();
}
$("sheet-tabs").addEventListener("click", (event) => {
  const id = (event.target as HTMLElement).closest<HTMLElement>("[data-sheet]")
    ?.dataset.sheet;
  if (id) run(() => activateSheet(id))();
});
$("sheet-tabs").addEventListener("keydown", (event) => {
  if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
  const buttons = Array.from(
    $("sheet-tabs").querySelectorAll<HTMLButtonElement>("button"),
  );
  const index = buttons.indexOf(event.target as HTMLButtonElement);
  if (index < 0) return;
  event.preventDefault();
  const next =
    event.key === "Home"
      ? 0
      : event.key === "End"
        ? buttons.length - 1
        : (index + (event.key === "ArrowRight" ? 1 : -1) + buttons.length) %
          buttons.length;
  const id = buttons[next].dataset.sheet!;
  run(() => {
    activateSheet(id);
    $("sheet-tabs")
      .querySelector<HTMLButtonElement>('[aria-selected="true"]')
      ?.focus();
  })();
});
$("sheet-picker").addEventListener(
  "change",
  run(() => activateSheet(($("sheet-picker") as HTMLSelectElement).value)),
);
bind("sheet-add", addSheet);
bind("sheet-manage", async () => {
  const target = editor;
  const { openSheetManager } = await import("./sheet-manager");
  if (target !== editor || loading || document.querySelector("dialog[open]"))
    return;
  openSheetManager(target, updateAfterEdit, onSelection);
});
document.addEventListener(
  "manage-sheets",
  run(() => $("sheet-manage").click()),
);
bind("sheet-rename", async () => {
  const sheet = editor.current().getActiveSheet();
  if (
    !(await modal(
      "시트 이름 변경",
      `<label for="sheet-name">시트 이름</label><input id="sheet-name" maxlength="31" value="${esc(sheet.getSheetName())}" required autofocus><p class="muted">최대 31자. : \\ / ? * [ ] 문자는 사용할 수 없어요.</p>`,
      "이름 변경",
    ))
  )
    return;
  const name = ($("sheet-name") as HTMLInputElement).value.trim();
  if (
    !name ||
    /[:\\/?*\[\]]/.test(name) ||
    name.startsWith("'") ||
    name.endsWith("'")
  )
    throw new Error("올바른 시트 이름을 입력해 주세요.");
  if (
    editor
      .current()
      .getSheets()
      .some(
        (other) =>
          other.getSheetId() !== sheet.getSheetId() &&
          other.getSheetName().toLowerCase() === name.toLowerCase(),
      )
  )
    throw new Error("같은 이름의 시트가 있어요. 다른 이름을 입력해 주세요.");
  sheet.setName(name);
  onChange();
  toast("시트 이름을 변경했습니다.");
});
$("sheet-zoom").addEventListener(
  "change",
  run(() => {
    editor
      .current()
      .getActiveSheet()
      .zoom(Number(($("sheet-zoom") as HTMLSelectElement).value));
    updateSheets();
  }),
);
bind("history", () =>
  openHistory({
    snapshot: () => editor.snapshot(),
    save: async () => {
      await saveCurrent(false);
    },
    restore: async (book) => {
      await openBook(book);
      await saveCurrent(false);
      closeLibrary();
    },
    toast,
  }),
);
bind("conflict-copy", async () => {
  await saving.catch(() => {});
  const copy = editor.snapshot();
  copy.id = crypto.randomUUID();
  copy.name = copy.name.slice(0, 89) + " 내 변경 사본";
  await saveBook(copy, { expectedUpdatedAt: null });
  await refreshLibrary();
  await openBook(copy, false);
  toast(
    "내 변경을 별도 사본으로 보관했습니다. 기존 문서의 최신 보관본은 유지됩니다.",
  );
});
bind("conflict-reload", async () => {
  if (
    !(await modal(
      "최신 보관본 열기",
      "<p>이 탭의 미보관 변경을 버리고 다른 탭의 최신 보관본을 엽니다. 변경을 유지하려면 취소하고 ‘내 변경 사본 보관’을 눌러 주세요.</p>",
      "최신 보관본 열기",
    ))
  )
    return;
  await saving.catch(() => {});
  await refreshLibrary();
  const latest = books.find((book) => book.id === editor.current().getId());
  if (!latest)
    throw new Error(
      "다른 탭에서 보관본이 삭제됐어요. 내 변경을 사본으로 보관해 주세요.",
    );
  await openBook(structuredClone(latest.book), false);
});
function renderLibrary() {
  const query = ($("search") as HTMLInputElement).value.trim().toLowerCase();
  const matched = books.filter((b) => b.name.toLowerCase().includes(query));
  $("library").innerHTML = matched.length
    ? matched
        .map(
          (b) =>
            `<div class="book-row ${editor?.current().getId() === b.id ? "active" : ""}"><button class="book-open" data-book="${esc(b.id)}">${icon("file-spreadsheet")}<span><strong>${esc(b.name)}</strong><small>${new Date(b.updatedAt).toLocaleDateString("ko-KR")} · 이 기기</small></span></button><button class="book-delete icon-btn" data-delete="${esc(b.id)}" aria-label="${esc(b.name)} 보관본 삭제">${icon("trash-2")}</button></div>`,
        )
        .join("")
    : `<div class="empty-library">${icon("folder-open")}<strong>${query ? "검색 결과가 없어요" : "아직 보관한 문서가 없어요"}</strong><p>${query ? "다른 이름으로 찾아보세요." : "셀을 편집하면 이곳에<br>자동으로 보관됩니다."}</p></div>`;
  icons();
}
async function saveCurrent(notify = true) {
  if (!editor) return;
  clearTimeout(autoTimer);
  const snapshot = editor.snapshot(),
    version = revision;
  saving = saving
    .catch(() => {})
    .then(async () => {
      try {
        const saved = await saveBook(snapshot, {
          expectedUpdatedAt: knownVersions.get(snapshot.id) ?? null,
        });
        knownVersions.set(snapshot.id, saved.updatedAt);
        if (editor.current().getId() === snapshot.id && revision === version) {
          dirty = false;
          $("save-state").textContent = "이 브라우저에 보관됨";
        }
        try {
          localStorage.setItem("excel-workspace-last", snapshot.id);
        } catch {
          /* IndexedDB remains the source of truth. */
        }
        await refreshLibrary();
        if (notify)
          toast(
            "이 브라우저에 보관했습니다. 파일 저장은 별도로 내려받아 주세요.",
          );
      } catch (e) {
        if (
          e instanceof StorageConflictError &&
          editor.current().getId() === snapshot.id
        ) {
          conflictBookId = snapshot.id;
          $("save-conflict").hidden = false;
          $("save-state").textContent = "저장 충돌 · 사본 보관이 필요해요";
          clearTimeout(autoTimer);
        } else
          $("save-state").textContent = "보관 실패 · 파일을 내려받아 주세요";
        throw e;
      }
    });
  return saving;
}
async function openBook(book: IWorkbookData, savePrevious = true) {
  if (loading)
    throw new Error(
      "통합문서를 준비하고 있습니다. 잠시 후 다시 시도해 주세요.",
    );
  validateBook(book);
  if (savePrevious && dirty) await saveCurrent(false);
  clearTimeout(autoTimer);
  knownVersions.set(
    book.id,
    books.find((b) => b.id === book.id)?.updatedAt ?? null,
  );
  loading = true;
  $("loading").hidden = false;
  $("editor-area").setAttribute("aria-busy", "true");
  for (const el of document.querySelectorAll<HTMLElement>(
    ".ribbon,.filebar,#cell-form,.sheetbar",
  ))
    el.inert = true;
  try {
    const { createEditor } = await import("./editor");
    editor?.dispose();
    $("editor").replaceChildren();
    editor = createEditor($("editor"), book, onChange, onSelection);
    await editor.ready;
    $("document-name").textContent = book.name;
    document.title = `${book.name} · 오픈소스 엑셀`;
    dirty = false;
    conflictBookId = null;
    $("save-conflict").hidden = true;
    sheetSignature = "";
    revision++;
    $("save-state").textContent = books.some((b) => b.id === book.id)
      ? "이 브라우저에 보관됨"
      : "편집하면 자동 보관";
    $("start-hint").hidden =
      Object.keys(book.sheets[book.sheetOrder[0]].cellData ?? {}).length > 0;
    onSelection();
    renderLibrary();
  } finally {
    loading = false;
    $("loading").hidden = true;
    $("editor-area").removeAttribute("aria-busy");
    for (const el of document.querySelectorAll<HTMLElement>(
      ".ribbon,.filebar,#cell-form,.sheetbar",
    ))
      el.inert = false;
  }
}
async function modal(
  title: string,
  body: string,
  ok = "확인",
): Promise<boolean> {
  const dialog = $("modal") as HTMLDialogElement;
  $("modal-title").textContent = title;
  $("modal-body").innerHTML = body;
  $("modal-ok").textContent = ok;
  dialog.showModal();
  return new Promise((resolve) =>
    dialog.addEventListener(
      "close",
      () => resolve(dialog.returnValue === "ok"),
      { once: true },
    ),
  );
}
function closeLibrary() {
  const d = $("library-dialog") as HTMLDialogElement;
  if (d.open) d.close();
}
$("library-dialog").addEventListener("close", () =>
  document.querySelector(".workspace")!.prepend($("sidebar")),
);
document
  .querySelectorAll<HTMLElement>("[data-close]")
  .forEach(
    (b) =>
      (b.onclick = () => ($(b.dataset.close!) as HTMLDialogElement).close()),
  );
bind("library-toggle", () => {
  if (matchMedia("(max-width: 800px)").matches) {
    $("library-host").append($("sidebar"));
    ($("library-dialog") as HTMLDialogElement).showModal();
  } else {
    document.body.classList.toggle("sidebar-hidden");
    window.dispatchEvent(new Event("resize"));
  }
});
bind("new", async () => {
  await openBook(newBook());
  closeLibrary();
});
bind("example", async () => {
  await openBook(newBook(true));
  await saveCurrent(false);
  closeLibrary();
});
bind("hint-example", () => $("example").click());
bind("hint-close", () => ($("start-hint").hidden = true));
bind("save", () => saveCurrent());
bind("rename", async () => {
  if (
    await modal(
      "통합문서 이름",
      `<label for="new-name">문서 이름</label><input id="new-name" maxlength="100" value="${esc(editor.current().getName())}" required autofocus>`,
      "이름 변경",
    )
  ) {
    const name = ($("new-name") as HTMLInputElement).value.trim();
    if (!name) throw new Error("이름을 입력해 주세요.");
    editor.current().setName(name);
    $("document-name").textContent = name;
    document.title = `${name} · 오픈소스 엑셀`;
    onChange();
  }
});
bind("open", async () => {
  if (
    await modal(
      "파일 가져오기",
      '<p>XLSX·CSV 또는 이 앱의 JSON 백업을 열 수 있습니다. 파일은 이 기기에서 처리합니다.</p><div class="notice">값·수식·기본 글꼴과 색·표시 형식·병합·행 높이·열 너비를 가져옵니다. 차트·피벗·매크로·그림·고급 서식은 보존되지 않습니다. 원본을 따로 보관해 주세요.</div><p class="muted">최대 10MB · 30시트 · 100,000개 셀<br>구형 .xls와 암호화된 파일은 지원하지 않습니다.</p>',
      "파일 선택",
    )
  )
    ($("file") as HTMLInputElement).click();
});
$("file").addEventListener(
  "change",
  run(async () => {
    const input = $("file") as HTMLInputElement,
      file = input.files?.[0];
    input.value = "";
    if (!file) return;
    if (file.size > LIMITS.bytes)
      throw new Error("파일은 10MB 이하로 선택해 주세요.");
    $("loading").hidden = false;
    try {
      const io = await import("./files");
      let book: IWorkbookData;
      if (/\.xlsx$/i.test(file.name))
        book = await io.importXlsx(await file.arrayBuffer(), file.name);
      else if (/\.csv$/i.test(file.name))
        book = io.importCsv(await file.text(), file.name);
      else if (/\.json$/i.test(file.name)) {
        const data = JSON.parse(await file.text());
        if (data.format !== "excel-workspace" || data.version !== 1)
          throw new Error("이 앱에서 만든 v1 백업을 선택해 주세요.");
        validateBook(data.book);
        book = data.book;
        book.id = crypto.randomUUID();
      } else throw new Error("XLSX·CSV·JSON 백업만 열 수 있습니다.");
      await openBook(book);
      await saveCurrent(false);
      toast("파일을 열었습니다. 원본 파일은 변경하지 않았습니다.");
    } catch (e) {
      if (
        e instanceof Error &&
        /zip|central|corrupt|encrypted/i.test(e.message)
      )
        throw new Error(
          "손상되었거나 암호화된 XLSX입니다. 엑셀에서 확인한 뒤 다시 저장해 주세요.",
        );
      throw e;
    } finally {
      $("loading").hidden = true;
    }
  }),
);
bind("export", async () => {
  const io = await import("./files");
  io.download(
    await io.exportXlsx(editor.snapshot()),
    editor.current().getName() + ".xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  toast("XLSX 다운로드를 시작했습니다. 다운로드 폴더를 확인해 주세요.");
});
bind("csv", async () => {
  const io = await import("./files");
  io.download(
    io.exportCsv(
      editor.snapshot(),
      editor.current().getActiveSheet().getSheetId(),
    ),
    editor.current().getName() + ".csv",
    "text/csv;charset=utf-8",
  );
  toast(
    "현재 시트의 값을 CSV로 내려받습니다. 서식과 수식은 포함하지 않습니다.",
  );
});
bind("backup", async () => {
  const io = await import("./files");
  io.download(
    JSON.stringify({
      format: "excel-workspace",
      version: 1,
      book: editor.snapshot(),
    }),
    editor.current().getName() + ".excel-workspace.json",
    "application/json",
  );
  toast(
    "현재 문서의 백업 다운로드를 시작했습니다. 열기에서 복원할 수 있습니다.",
  );
});
$("search").addEventListener("input", renderLibrary);
$("library").addEventListener("click", async (e) => {
  const target = (e.target as HTMLElement).closest<HTMLElement>(
    "[data-book],[data-delete]",
  );
  if (!target) return;
  await run(async () => {
    const id = target.dataset.book || target.dataset.delete;
    const saved = books.find((b) => b.id === id);
    if (!saved) return;
    if (target.dataset.book) {
      await openBook(structuredClone(saved.book));
      closeLibrary();
    } else if (
      await modal(
        "보관본 삭제",
        `<p>‘${esc(saved.name)}’의 브라우저 보관본을 삭제할까요? 내려받은 파일은 유지됩니다.</p>`,
        "삭제",
      )
    ) {
      if (editor.current().getId() === id) {
        clearTimeout(autoTimer);
        dirty = false;
        await openBook(newBook(), false);
      }
      await deleteBook(id!);
      await refreshLibrary();
    }
  })();
});
bind("undo", () => editor.api.undo());
bind("redo", () => editor.api.redo());
bind("bold", () => {
  editor
    .range()
    .setFontWeight(editor.range().getCellStyle()?.bold ? "normal" : "bold");
  onChange();
});
bind("italic", () => {
  editor
    .range()
    .setFontStyle(editor.range().getCellStyle()?.italic ? "normal" : "italic");
  onChange();
});
$("font-size").addEventListener(
  "change",
  run(() => {
    editor
      .range()
      .setFontSize(Number(($("font-size") as HTMLSelectElement).value));
    onChange();
  }),
);
$("number-format").addEventListener(
  "change",
  run(() => {
    editor
      .range()
      .setNumberFormat(($("number-format") as HTMLSelectElement).value);
    onChange();
  }),
);
bind("filter", () => {
  const filter = editor.current().getActiveSheet().getFilter();
  if (filter) filter.remove();
  else {
    const r = editor.range().getRange();
    if (r.endRow <= r.startRow)
      throw new Error("제목 행과 데이터 행을 함께 선택해 주세요.");
    editor.range().createFilter();
  }
  onChange();
});
function updateAfterEdit() {
  onChange();
  onSelection();
}
bind("sort", async () => {
  const target = editor;
  const { openSortTools } = await import("./sort-tools");
  if (target !== editor || loading || document.querySelector("dialog[open]"))
    return;
  openSortTools(target, updateAfterEdit, toast);
});
$("address").addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    run(() => editor.select(($("address") as HTMLInputElement).value.trim()))();
  }
});
$("cell-form").addEventListener("submit", (e) => {
  e.preventDefault();
  run(() => {
    editor.setCell(($("cell-input") as HTMLTextAreaElement).value);
    ($("cell-input") as HTMLTextAreaElement).blur();
    onSelection();
  })();
});
function resizeFormulaInput() {
  const input = $("cell-input") as HTMLTextAreaElement;
  input.rows = Math.min(3, Math.max(1, input.value.split("\n").length));
  const functionName = input.value
    .match(/^=\s*([A-Z]+)\s*\(/i)?.[1]
    .toUpperCase();
  const hints: Record<string, string> = {
    SUM: "SUM(범위) · 숫자를 더해요. 예: =SUM(B2:B10)",
    AVERAGE: "AVERAGE(범위) · 숫자의 평균을 구해요. 빈 셀은 제외해요.",
    IF: 'IF(조건, 참일 때, 거짓일 때) · 예: =IF(B2>=60,"통과","재검토")',
    COUNT: "COUNT(범위) · 숫자가 있는 셀의 개수를 세요.",
    COUNTA: "COUNTA(범위) · 비어 있지 않은 셀의 개수를 세요.",
    SUMIF:
      'SUMIF(조건 범위, 조건, 합계 범위) · 예: =SUMIF(A2:A10,"완료",B2:B10)',
    COUNTIF: 'COUNTIF(범위, 조건) · 예: =COUNTIF(A2:A10,"완료")',
    IFERROR: "IFERROR(계산, 오류일 때) · 예: =IFERROR(A2/B2,0)",
    ROUND: "ROUND(숫자, 자릿수) · 예: =ROUND(B2,2)",
    VLOOKUP:
      "VLOOKUP(찾을 값, 표 범위, 열 번호, FALSE) · 표의 첫 열에서 정확히 일치하는 값을 찾아요.",
  };
  const hint = functionName && hints[functionName];
  $("formula-help").textContent = hint ? `함수 도움말 · ${hint}` : "";
  $("formula-help").hidden = !hint || document.activeElement !== input;
}
$("cell-input").addEventListener("input", resizeFormulaInput);
$("cell-input").addEventListener("focus", resizeFormulaInput);
$("cell-input").addEventListener("blur", () => {
  $("formula-help").hidden = true;
});
$("cell-input").addEventListener("keydown", (e) => {
  if (e.key !== "Enter" || e.isComposing) return;
  if (e.altKey) {
    e.preventDefault();
    const field = $("cell-input") as HTMLTextAreaElement;
    field.setRangeText("\n", field.selectionStart, field.selectionEnd, "end");
    resizeFormulaInput();
  } else if (!e.shiftKey) {
    e.preventDefault();
    ($("cell-form") as HTMLFormElement).requestSubmit();
  }
});
bind("help", async () => {
  await modal(
    "오픈소스 엑셀 사용 안내",
    "<p>셀을 더블 클릭하거나 위의 입력줄로 값을 수정하세요. <b>=</b>로 시작하면 수식을 계산합니다. 셀 주소에 <b>A1:D10</b>을 입력하면 범위를 선택합니다.</p><p>Tab·Enter·방향키로 셀 이동, Ctrl+C/V로 복사·붙여넣기, Ctrl+Z로 실행 취소합니다. 시트 탭의 +로 시트를 추가하고, 우클릭 메뉴에서 행·열과 시트를 관리합니다.</p><p><b>보관</b>은 이 브라우저의 IndexedDB 저장입니다. 브라우저 데이터 삭제 시 사라질 수 있습니다. <b>XLSX 저장</b>이나 JSON 백업으로 별도 보관하세요. 다른 기기와 자동 동기화하거나 공동 편집하지 않습니다.</p><p>Univer(Apache-2.0)·ExcelJS(MIT) 기반의 독립 편집기입니다. Microsoft Excel 공식 제품이 아니며 차트·피벗·매크로·그림·고급 서식과 완전 호환을 보장하지 않습니다. 모바일은 간편 입력을 제공하며 복잡한 편집은 PC를 권장합니다.</p><p>AI는 개인 API 키 없이 Coders 관리형 Claude를 사용합니다. 질문과 선택 전송한 셀만 제공자에게 전달됩니다. 결과를 검토한 뒤 적용하세요.</p>",
    "확인",
  );
});
let aiStatus:
  | { aiAvailable: boolean; signedIn: boolean; managed: boolean }
  | undefined;
bind("ai-open", async () => {
  ($("ai-dialog") as HTMLDialogElement).showModal();
  try {
    const res = await fetch("/api/status");
    if (!res.ok) throw new Error();
    aiStatus = await res.json();
    $("ai-status").innerHTML = aiStatus?.aiAvailable
      ? aiStatus.managed && !aiStatus.signedIn
        ? `AI 사용에는 Coders 로그인이 필요합니다. <a href="https://mcp.coders.kr/sso/login?return_to=${encodeURIComponent(location.origin)}">로그인하기 →</a>`
        : "AI 연결 설정됨 · 질문과 동의한 셀만 전송"
      : "AI 연결 전입니다. 배포 환경에서 연결되면 사용할 수 있어요. 엑셀 기능은 지금도 사용 가능합니다.";
  } catch {
    $("ai-status").textContent =
      "AI 서버 상태를 확인하지 못했습니다. 잠시 후 다시 열어 주세요.";
  }
});
document.querySelectorAll<HTMLElement>("[data-prompt]").forEach(
  (b) =>
    (b.onclick = () => {
      ($("ai-prompt") as HTMLTextAreaElement).value = b.dataset.prompt!;
      $("ai-prompt").focus();
    }),
);
$("ai-form").addEventListener("submit", (e) => {
  e.preventDefault();
  run(async () => {
    if (!aiStatus?.aiAvailable)
      throw new Error(
        "AI가 아직 연결되지 않았습니다. 편집 기능은 계속 사용할 수 있습니다.",
      );
    if (aiStatus.managed && !aiStatus.signedIn)
      throw new Error("위의 로그인하기를 누른 뒤 다시 질문해 주세요.");
    const r = editor.range(),
      bounds = r.getRange(),
      sheet = editor.current().getActiveSheet(),
      cell = sheet.getRange(bounds.startRow, bounds.startColumn);
    const target = {
      book: editor.current().getId(),
      sheet: sheet.getSheetId(),
      address: cell.getA1Notation(),
      previous: JSON.stringify(cell.getCellData()),
    };
    let context = "";
    if (($("include-cells") as HTMLInputElement).checked) {
      if (
        (bounds.endRow - bounds.startRow + 1) *
          (bounds.endColumn - bounds.startColumn + 1) >
        200
      )
        throw new Error("선택 범위를 200셀 이내로 줄여 주세요.");
      const selectedData =
        editor.snapshot().sheets[sheet.getSheetId()].cellData ?? {};
      context = JSON.stringify({
        sheet: sheet.getSheetName(),
        range: r.getA1Notation(),
        values: Array.from(
          { length: bounds.endRow - bounds.startRow + 1 },
          (_, row) =>
            Array.from(
              { length: bounds.endColumn - bounds.startColumn + 1 },
              (_, column) =>
                cellText(
                  selectedData[bounds.startRow + row]?.[
                    bounds.startColumn + column
                  ],
                ),
            ),
        ),
        formulas: r.getFormulas(),
      });
      if (context.length > 16000)
        throw new Error("선택한 내용이 너무 깁니다. 범위를 줄여 주세요.");
    }
    ($("ai-send") as HTMLButtonElement).disabled = true;
    $("ai-send").textContent = "AI가 생각하고 있어요…";
    $("ai-result").hidden = true;
    formula = null;
    try {
      const res = await fetch("/api/ai", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          prompt: ($("ai-prompt") as HTMLTextAreaElement).value,
          context,
        }),
        signal: AbortSignal.timeout(45000),
      });
      if (!res.headers.get("content-type")?.includes("application/json"))
        throw new Error(
          "로그인 또는 사용량 상태를 확인한 뒤 다시 시도해 주세요.",
        );
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "AI 요청에 실패했습니다.");
      formula = data.formula;
      aiTarget = target;
      $("ai-explanation").textContent = data.explanation;
      $("ai-formula").textContent = formula;
      $("ai-formula").hidden = !formula;
      $("ai-apply").hidden = !formula;
      $("ai-apply").textContent = `${target.address}에 수식 적용`;
      $("ai-result").hidden = false;
      $("ai-result").scrollIntoView({ behavior: "smooth", block: "nearest" });
    } finally {
      ($("ai-send") as HTMLButtonElement).disabled = false;
      $("ai-send").textContent = "AI에게 질문하기";
    }
  })();
});
bind("ai-apply", async () => {
  if (!formula || !aiTarget) return;
  if (
    editor.current().getId() !== aiTarget.book ||
    editor.current().getActiveSheet().getSheetId() !== aiTarget.sheet
  )
    throw new Error("질문할 때의 문서와 시트로 돌아간 뒤 적용해 주세요.");
  const cell = editor.current().getActiveSheet().getRange(aiTarget.address);
  if (JSON.stringify(cell.getCellData()) !== aiTarget.previous)
    throw new Error("질문 이후 셀 내용이 변경되었습니다. 다시 질문해 주세요.");
  if (
    cell.getValue() != null &&
    cell.getValue() !== "" &&
    !(await modal(
      "기존 셀 덮어쓰기",
      `<p>${esc(aiTarget.address)}의 내용을 AI 수식으로 바꿀까요? 실행 취소로 되돌릴 수 있습니다.</p>`,
      "적용",
    ))
  )
    return;
  cell.setValue(formula);
  onChange();
  $("ai-apply").hidden = true;
  toast(
    `${aiTarget.address}에 수식을 적용했습니다. 계산 결과를 확인해 주세요.`,
  );
});
window.addEventListener("beforeunload", (e) => {
  if (dirty) {
    e.preventDefault();
    e.returnValue = "";
  }
});
document.addEventListener("keydown", (e) => {
  if (
    (e.ctrlKey || e.metaKey) &&
    ["f", "h"].includes(e.key.toLowerCase()) &&
    editor &&
    !loading &&
    !document.querySelector("dialog[open]")
  ) {
    e.preventDefault();
    run(() => openDataTools(editor, onChange, toast, "find"))();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
    e.preventDefault();
    run(() => saveCurrent())();
  }
});
async function boot() {
  try {
    await refreshLibrary();
    let last: string | null = null;
    try {
      last = localStorage.getItem("excel-workspace-last");
    } catch {
      /* Storage disabled. */
    }
    await openBook(books.find((b) => b.id === last)?.book ?? newBook(), false);
  } catch (e) {
    fail(e);
    if (!editor) await openBook(newBook(), false);
  }
}
void boot();
