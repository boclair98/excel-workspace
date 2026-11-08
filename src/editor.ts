import {
  createUniver,
  LocaleType,
  mergeLocales,
  defaultTheme,
  type IWorkbookData,
} from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import CoreKo from "@univerjs/preset-sheets-core/locales/ko-KR";
import { UniverSheetsFilterPreset } from "@univerjs/preset-sheets-filter";
import FilterKo from "@univerjs/preset-sheets-filter/locales/ko-KR";
import { UniverSheetsSortPreset } from "@univerjs/preset-sheets-sort";
import SortKo from "@univerjs/preset-sheets-sort/locales/ko-KR";
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-filter/lib/index.css";
import "@univerjs/preset-sheets-sort/lib/index.css";

export function createEditor(
  container: HTMLElement,
  book: IWorkbookData,
  onChange: () => void,
  onSelection: () => void,
) {
  const { univer, univerAPI: api } = createUniver({
    theme: {
      ...defaultTheme,
      primary: {
        50: "#f0f8f3",
        100: "#e3f2e8",
        200: "#c0e3ce",
        300: "#91cbaa",
        400: "#57ab7e",
        500: "#278659",
        600: "#18794e",
        700: "#12623e",
        800: "#104d34",
        900: "#0c3c29",
      },
    },
    locale: LocaleType.KO_KR,
    locales: { [LocaleType.KO_KR]: mergeLocales(CoreKo, FilterKo, SortKo) },
    presets: [
      UniverSheetsCorePreset({
        container,
        header: false,
        toolbar: false,
        formulaBar: false,
        footer: false,
        disableAutoFocus: true,
      }),
      UniverSheetsFilterPreset(),
      UniverSheetsSortPreset(),
    ],
  });
  const ready = new Promise<void>((resolve, reject) => {
    const timeout = setTimeout(
      () =>
        reject(
          new Error(
            "편집기 준비에 시간이 오래 걸립니다. 새로고침 후 다시 시도해 주세요.",
          ),
        ),
      25000,
    );
    const listener = api.addEvent(api.Event.LifeCycleChanged, (event) => {
      if (event.stage >= 2) {
        clearTimeout(timeout);
        listener.dispose();
        resolve();
      }
    });
  });
  api.createWorkbook(book);
  const shortcuts = api.getShortcut();
  const syncInputFocus = () => {
    const target = document.activeElement;
    const external = target instanceof HTMLElement && !container.contains(target) &&
      Boolean(target.closest('input,textarea,select,[contenteditable="true"],#ai-dialog,#ai-resize'));
    if (external) shortcuts.disableShortcut();
    else shortcuts.enableShortcut();
  };
  document.addEventListener("focusin", syncInputFocus);
  syncInputFocus();
  api.addEvent(api.Event.CommandExecuted, (event) => {
    if (event.type === 2) queueMicrotask(onSelection);
    if (
      event.type === 2 &&
      !event.id.includes("formula-calculation") &&
      !event.id.includes("set-formula")
    )
      onChange();
  });
  api.addEvent(api.Event.SelectionChanged, () => onSelection());
  const current = () => {
    const b = api.getActiveWorkbook();
    if (!b) throw new Error("통합문서를 준비하는 중입니다.");
    return b;
  };
  const range = () =>
    current().getActiveRange() ?? current().getActiveSheet().getRange("A1");
  return {
    api,
    ready,
    current,
    range,
    snapshot: () => current().save(),
    dispose: () => {
      document.removeEventListener("focusin", syncInputFocus);
      shortcuts.enableShortcut();
      univer.dispose();
    },
    select(address: string) {
      if (
        !/^[A-Za-z]{1,2}[1-9]\d{0,3}(?::[A-Za-z]{1,2}[1-9]\d{0,3})?$/.test(
          address,
        )
      )
        throw new Error("A1 또는 A1:D10처럼 입력해 주세요.");
      const r = current().getActiveSheet().getRange(address);
      const bounds = r.getRange();
      if (bounds.endColumn >= 200 || bounds.endRow >= 10000)
        throw new Error("10,000행·200열 안에서 선택해 주세요.");
      current().getActiveSheet().setActiveRange(r);
      current()
        .getActiveSheet()
        .scrollToCell(bounds.startRow, bounds.startColumn, 0);
      onSelection();
      return r;
    },
    setCell(value: string) {
      const r = range().getRange();
      current()
        .getActiveSheet()
        .getRange(r.startRow, r.startColumn)
        .setValue(value);
      onChange();
    },
  };
}
export type Editor = ReturnType<typeof createEditor>;
