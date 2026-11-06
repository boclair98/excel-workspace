import {
  createIcons,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Underline,
  PaintBucket,
  Type,
  Merge,
  WrapText,
  Rows3,
  Columns3,
  FilePlus2,
  Snowflake,
  Eraser,
  Search,
  Settings2,
  PanelLeft,
  Grid2X2,
  ChevronUp,
  ChevronDown,
} from "lucide";
import type { Editor } from "./editor";
type Options = {
  editor: () => Editor;
  changed: () => void;
  modal: (title: string, body: string, ok?: string) => Promise<boolean>;
  toast: (message: string, error?: boolean) => void;
};
const i = (name: string) => `<i data-lucide="${name}" aria-hidden="true"></i>`;
const button = (id: string, name: string, label: string, small = false) =>
  `<button id="${id}" ${small ? `class="icon-btn" aria-label="${label}" title="${label}"` : ""}>${i(name)}${small ? "" : label}</button>`;
const group = (label: string, body: string) =>
  `<div class="ribbon-group"><div class="ribbon-controls">${body}</div><span class="ribbon-caption">${label}</span></div>`;
export function mountRibbon(options: Options) {
  const bar = document.querySelector(".quickbar")!;
  const old = (id: string) => document.getElementById(id)!.outerHTML;
  const undo = old("undo"),
    redo = old("redo"),
    bold = old("bold"),
    italic = old("italic"),
    fontSize = old("font-size"),
    format = old("number-format"),
    filter = old("filter"),
    sort = old("sort");
  const ribbon = document.createElement("section");
  ribbon.className = "ribbon";
  ribbon.setAttribute("aria-label", "편집 도구");
  ribbon.innerHTML = `<div class="ribbon-tabs" role="tablist" aria-label="리본 메뉴">${["홈", "삽입", "데이터", "보기"].map((s, k) => `<button role="tab" id="tab-${k}" aria-controls="panel-${k}" aria-selected="${k === 0}" tabindex="${k === 0 ? 0 : -1}">${s}</button>`).join("")}<span class="ribbon-note">익숙한 표, 더 가벼운 작업</span><button id="full-tools" class="icon-btn" aria-label="리본 접기" title="리본 접기" aria-pressed="false">${i("chevron-up")}</button></div>
  <div class="ribbon-panel" id="panel-0" role="tabpanel" aria-labelledby="tab-0">${group("실행 취소", undo + redo)}${group("글꼴", `<label class="sr-only" for="font-family">글꼴</label><select id="font-family"><option value="Noto Sans KR Variable">Noto Sans KR</option><option value="Arial">Arial</option><option value="Malgun Gothic">맑은 고딕</option><option value="Georgia">Georgia</option><option value="Courier New">Courier New</option></select><label class="sr-only" for="font-size">글자 크기</label>${fontSize}${bold}${italic}${button("underline", "underline", "밑줄", true)}<label class="color-picker" title="글자 색상">${i("type")}<input type="color" id="font-color" value="#191f28" aria-label="글자 색상"></label><label class="color-picker" title="셀 채우기">${i("paint-bucket")}<input type="color" id="fill-color" value="#eaf5ef" aria-label="셀 채우기 색상"></label>`)}${group("맞춤", button("align-left", "align-left", "왼쪽 정렬", true) + button("align-center", "align-center", "가운데 정렬", true) + button("align-right", "align-right", "오른쪽 정렬", true) + button("wrap", "wrap-text", "줄 바꿈", true) + button("merge", "merge", "병합", true) + button("border", "grid-2-x2", "모든 테두리", true))}${group("표시 형식", `<label class="sr-only" for="number-format">표시 형식</label>${format}`)}</div>
  <div class="ribbon-panel" id="panel-1" role="tabpanel" aria-labelledby="tab-1" hidden>${group("행과 열", button("insert-row", "rows-3", "행 삽입") + button("insert-column", "columns-3", "열 삽입"))}${group("통합문서", button("insert-sheet", "file-plus-2", "새 시트") + button("duplicate-book", "file-plus-2", "문서 복제"))}${group("셀 크기", `<button id="row-height">행 높이</button><button id="column-width">열 너비</button>`)}${group("표 만들기", `<button id="template-budget">예시 가계부</button>`)}</div>
  <div class="ribbon-panel" id="panel-2" role="tabpanel" aria-labelledby="tab-2" hidden>${group("정렬 및 필터", sort + filter)}${group("찾기", button("find", "search", "내용 찾기"))}${group("데이터 정리", button("clear-cells", "eraser", "선택 내용 지우기"))}<p class="ribbon-tip">먼저 셀 범위를 선택하세요.<br>수식 입력줄 왼쪽에 A1:D10을 입력해도 돼요.</p></div>
  <div class="ribbon-panel" id="panel-3" role="tabpanel" aria-labelledby="tab-3" hidden>${group("틀 고정", button("freeze-row", "snowflake", "첫 행 고정") + button("freeze-column", "snowflake", "첫 열 고정") + `<button id="unfreeze">고정 해제</button>`)}${group("화면", `<label class="sr-only" for="zoom">화면 배율</label><select id="zoom" aria-label="화면 배율"><option value="0.75">75%</option><option value="1" selected>100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select>${button("view-library", "panel-left", "보관함")}`)}${group("내 작업 공간", button("personalize", "settings-2", "화면 설정"))}</div>`;
  bar.replaceWith(ribbon);
  createIcons({
    icons: {
      AlignLeft,
      AlignCenter,
      AlignRight,
      Underline,
      PaintBucket,
      Type,
      Merge,
      WrapText,
      Rows3,
      Columns3,
      FilePlus2,
      Snowflake,
      Eraser,
      Search,
      Settings2,
      PanelLeft,
      Grid2X2,
      ChevronUp,
      ChevronDown,
    },
    attrs: { "stroke-width": 1.7 },
  });
  const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
    document.getElementById(id) as T;
  const run = (fn: () => unknown) => async () => {
    try {
      await fn();
    } catch (e) {
      options.toast(
        e instanceof Error ? e.message : "작업을 완료하지 못했습니다.",
        true,
      );
    }
  };
  const bind = (id: string, fn: () => unknown) =>
    $(id).addEventListener("click", run(fn));
  const edit = (fn: (editor: Editor) => void) => () => {
    fn(options.editor());
    options.changed();
  };
  function activate(index: number) {
    for (let k = 0; k < 4; k++) {
      $(`tab-${k}`).setAttribute("aria-selected", String(k === index));
      $(`tab-${k}`).tabIndex = k === index ? 0 : -1;
      $(`panel-${k}`).hidden = k !== index;
    }
    document.body.classList.remove("ribbon-collapsed");
    $("full-tools").setAttribute("aria-pressed", "false");
  }
  for (let k = 0; k < 4; k++) {
    bind(`tab-${k}`, () => activate(k));
    $(`tab-${k}`).addEventListener("keydown", (e) => {
      if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
        e.preventDefault();
        const next = (k + (e.key === "ArrowRight" ? 1 : 3)) % 4;
        activate(next);
        $(`tab-${next}`).focus();
      }
    });
  }
  bind("full-tools", () => {
    const collapsed = document.body.classList.toggle("ribbon-collapsed");
    $("full-tools").setAttribute("aria-pressed", String(collapsed));
    $("full-tools").setAttribute(
      "aria-label",
      collapsed ? "리본 펼치기" : "리본 접기",
    );
    window.dispatchEvent(new Event("resize"));
  });
  $("font-family").addEventListener(
    "change",
    run(
      edit((e) =>
        e.range().setFontFamily(($("font-family") as HTMLSelectElement).value),
      ),
    ),
  );
  $("font-color").addEventListener(
    "change",
    run(
      edit((e) =>
        e.range().setFontColor(($("font-color") as HTMLInputElement).value),
      ),
    ),
  );
  $("fill-color").addEventListener(
    "change",
    run(
      edit((e) =>
        e.range().setBackground(($("fill-color") as HTMLInputElement).value),
      ),
    ),
  );
  bind(
    "underline",
    edit((e) =>
      e
        .range()
        .setFontLine(
          e.range().getCellStyle()?.underline?.build().s === 1
            ? "none"
            : "underline",
        ),
    ),
  );
  bind(
    "align-left",
    edit((e) => e.range().setHorizontalAlignment("left")),
  );
  bind(
    "align-center",
    edit((e) => e.range().setHorizontalAlignment("center")),
  );
  bind(
    "align-right",
    edit((e) => e.range().setHorizontalAlignment("normal")),
  );
  bind(
    "wrap",
    edit((e) => e.range().setWrap(!e.range().getWrap())),
  );
  bind(
    "border",
    edit((e) =>
      e
        .range()
        .setBorder(
          e.api.Enum.BorderType.ALL,
          e.api.Enum.BorderStyleTypes.THIN,
          "#b7c5bc",
        ),
    ),
  );
  bind("merge", async () => {
    const e = options.editor(),
      r = e.range();
    if (r.isMerged()) {
      r.breakApart();
      options.changed();
      return;
    }
    if (
      await options.modal(
        "셀 병합",
        "<p>선택 범위를 병합합니다. 왼쪽 위 셀의 값만 남으므로 다른 값이 있으면 먼저 옮겨 주세요.</p>",
        "병합",
      )
    ) {
      r.merge();
      options.changed();
    }
  });
  bind(
    "insert-row",
    edit((e) =>
      e
        .current()
        .getActiveSheet()
        .insertRowBefore(e.range().getRange().startRow),
    ),
  );
  bind(
    "insert-column",
    edit((e) =>
      e
        .current()
        .getActiveSheet()
        .insertColumnBefore(e.range().getRange().startColumn),
    ),
  );
  bind(
    "insert-sheet",
    edit((e) => e.current().insertSheet()),
  );
  bind("duplicate-book", () =>
    document.dispatchEvent(new CustomEvent("duplicate-workbook")),
  );
  bind("template-budget", () => $("example").click());
  for (const [id, label, min, max, initial] of [
    ["row-height", "행 높이", 20, 300, 30],
    ["column-width", "열 너비", 40, 600, 120],
  ] as const) {
    bind(id, async () => {
      if (
        await options.modal(
          `${label} 조절`,
          `<label for="dimension">${label} (픽셀)</label><input id="dimension" type="number" min="${min}" max="${max}" value="${initial}" required>`,
          "적용",
        )
      ) {
        const value = Number(($("dimension") as HTMLInputElement).value);
        if (!Number.isFinite(value) || value < min || value > max)
          throw new Error(`${min}~${max} 사이 값을 입력해 주세요.`);
        const e = options.editor(),
          r = e.range().getRange();
        if (id === "row-height")
          e.current()
            .getActiveSheet()
            .setRowHeightsForced(r.startRow, r.endRow - r.startRow + 1, value);
        else
          e.current()
            .getActiveSheet()
            .setColumnWidths(
              r.startColumn,
              r.endColumn - r.startColumn + 1,
              value,
            );
        options.changed();
      }
    });
  }
  bind("clear-cells", async () => {
    if (
      await options.modal(
        "선택 내용 지우기",
        "<p>선택 범위의 값과 수식을 지웁니다. 서식은 유지하며 실행 취소로 되돌릴 수 있습니다.</p>",
        "지우기",
      )
    ) {
      options.editor().range().clearContent();
      options.changed();
    }
  });
  bind("find", async () => {
    if (
      !(await options.modal(
        "현재 시트에서 찾기",
        '<label for="find-text">찾을 내용 (첫 번째 일치 셀로 이동)</label><input id="find-text" type="search" required autofocus>',
        "찾기",
      ))
    )
      return;
    const query = ($("find-text") as HTMLInputElement).value
      .trim()
      .toLowerCase();
    if (!query) return;
    const e = options.editor(),
      s = e.snapshot().sheets[e.current().getActiveSheet().getSheetId()];
    for (const [r, row] of Object.entries(s.cellData ?? {}))
      for (const [c, value] of Object.entries(row)) {
        const cell = value as { v?: unknown; f?: string };
        if (
          String(cell?.v ?? cell?.f ?? "")
            .toLowerCase()
            .includes(query)
        ) {
          const range = e.current().getActiveSheet().getRange(+r, +c);
          e.current().getActiveSheet().setActiveRange(range);
          e.current().getActiveSheet().scrollToCell(+r, +c, 0);
          options.toast(`${range.getA1Notation()}에서 찾았습니다.`);
          return;
        }
      }
    options.toast("일치하는 셀이 없습니다.");
  });
  bind(
    "freeze-row",
    edit((e) => e.current().getActiveSheet().setFrozenRows(1)),
  );
  bind(
    "freeze-column",
    edit((e) => e.current().getActiveSheet().setFrozenColumns(1)),
  );
  bind(
    "unfreeze",
    edit((e) => {
      e.current().getActiveSheet().setFrozenRows(0);
      e.current().getActiveSheet().setFrozenColumns(0);
    }),
  );
  $("zoom").addEventListener(
    "change",
    run(() =>
      options
        .editor()
        .current()
        .getActiveSheet()
        .zoom(Number(($("zoom") as HTMLSelectElement).value)),
    ),
  );
  bind("view-library", () => $("library-toggle").click());
  const applyPreferences = (accent: string, density: string) => {
    document.documentElement.dataset.accent = accent;
    document.documentElement.dataset.density = density;
  };
  try {
    const p = JSON.parse(
      localStorage.getItem("excel-workspace-display") ?? "{}",
    );
    applyPreferences(p.accent ?? "green", p.density ?? "comfortable");
  } catch {
    /* Keep readable defaults. */
  }
  bind("personalize", async () => {
    if (
      await options.modal(
        "내 작업 공간 꾸미기",
        `<p>문서 내용과 서식은 바뀌지 않습니다.</p><label for="accent">강조색</label><select id="accent"><option value="green">포레스트 그린</option><option value="blue">클래식 블루</option><option value="slate">차분한 슬레이트</option></select><label for="density">도구 간격</label><select id="density"><option value="comfortable">여유롭게</option><option value="compact">촘촘하게 (PC)</option></select>`,
        "저장",
      )
    ) {
      const accent = ($("accent") as HTMLSelectElement).value,
        density = ($("density") as HTMLSelectElement).value;
      applyPreferences(accent, density);
      try {
        localStorage.setItem(
          "excel-workspace-display",
          JSON.stringify({ accent, density }),
        );
      } catch {
        options.toast(
          "화면에 적용했지만 브라우저 설정은 보관하지 못했습니다.",
          true,
        );
      }
      window.dispatchEvent(new Event("resize"));
    }
  });
}
