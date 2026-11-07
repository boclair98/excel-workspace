import {
  CellValueType,
  type ICellData,
  type IObjectMatrixPrimitiveType,
  type IWorkbookData,
} from "@univerjs/core";
import { safeFormula } from "../ai-formula.mjs";
import { cellAddress, cellFingerprint } from "./data-operations";
import { cellText, LIMITS, validateBook } from "./model";
import type { Editor } from "./editor";

type Value = string | number | boolean | null;
type Cell = {
  row: number;
  col: number;
  before: Value;
  fingerprint: string;
  protected: boolean;
};
type Change = { address: string; cell: Cell; after: Value; formula: string | null };

function protectedCell(
  sheet: IWorkbookData["sheets"][string],
  row: number,
  col: number,
) {
  const cell = sheet.cellData?.[row]?.[col];
  return Boolean(
    cell?.f ||
      cell?.si ||
      cell?.p ||
      cell?.ref ||
      cell?.fd ||
      cell?.ft != null ||
      sheet.mergeData?.some(
        (merge) =>
          row >= merge.startRow &&
          row <= merge.endRow &&
          col >= merge.startColumn &&
          col <= merge.endColumn,
      ),
  );
}

export function resolveAIRange(editor: Editor, prompt: string) {
  const selected = editor.range();
  const sheet = editor.current().getActiveSheet();
  const addresses = Array.from(
    prompt.matchAll(/(?<![\w.$!])\$?([a-z]{1,2})\$?([1-9]\d{0,4})(?![\w.(])/gi),
    ([, col, row]) => `${col.toUpperCase()}${row}`,
  );
  let range = selected;
  if (addresses.length) {
    const cells = addresses.map((address) => sheet.getRange(address).getRange());
    const startRow = Math.min(...cells.map((cell) => cell.startRow));
    const startColumn = Math.min(...cells.map((cell) => cell.startColumn));
    const endRow = Math.max(...cells.map((cell) => cell.endRow));
    const endColumn = Math.max(...cells.map((cell) => cell.endColumn));
    range = sheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
  }
  const bounds = range.getRange();
  if (
    (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1) > 200 ||
    bounds.endRow >= LIMITS.rows ||
    bounds.endColumn >= LIMITS.cols
  )
    throw new Error("표에 적용하려면 200셀 이내로 범위를 적어 주세요.");
  return range;
}

export function captureAISelection(editor: Editor, range: ReturnType<typeof resolveAIRange>) {
  const bounds = range.getRange();
  const book = editor.snapshot();
  const sheetId = range.getSheetId();
  const sheet = book.sheets[sheetId];
  if (!sheet) throw new Error("현재 시트에서 범위를 다시 선택해 주세요.");
  if (
    bounds.endRow >= (sheet.rowCount ?? LIMITS.rows) ||
    bounds.endColumn >= (sheet.columnCount ?? LIMITS.cols)
  )
    throw new Error("현재 시트에 있는 셀 주소를 적어 주세요.");
  const cells = new Map<string, Cell>();
  const values: Value[][] = [];
  const formulas: string[][] = [];
  for (let row = bounds.startRow; row <= bounds.endRow; row++) {
    const rowValues: Value[] = [];
    const rowFormulas: string[] = [];
    for (let col = bounds.startColumn; col <= bounds.endColumn; col++) {
      const cell = sheet.cellData?.[row]?.[col];
      const value = cell?.v == null && !cell?.p ? null : cellText(cell);
      const before =
        cell?.t === CellValueType.BOOLEAN && value != null
          ? Boolean(value)
          : value;
      cells.set(cellAddress(row, col), {
        row,
        col,
        before,
        fingerprint: cellFingerprint(cell),
        protected: protectedCell(sheet, row, col),
      });
      rowValues.push(before);
      rowFormulas.push(cell?.f ?? "");
    }
    values.push(rowValues);
    formulas.push(rowFormulas);
  }
  const label = range.getA1Notation();
  const context = JSON.stringify({
    sheet: sheet.name,
    range: label,
    values,
    formulas,
  });
  if (context.length > 16000)
    throw new Error("선택한 내용이 너무 깁니다. 범위를 줄여 주세요.");
  return {
    bookId: book.id,
    sheetId,
    sheetName: sheet.name,
    rows: sheet.rowCount,
    cols: sheet.columnCount,
    bounds,
    label,
    cells,
    context,
  };
}

export type AISelection = ReturnType<typeof captureAISelection>;

export function reviewAIChanges(
  selection: AISelection,
  edits: unknown,
): Change[] {
  if (!Array.isArray(edits) || edits.length > selection.cells.size)
    throw new Error("AI 변경안의 범위를 확인할 수 없습니다.");
  const seen = new Set<string>();
  const changes: Change[] = [];
  for (const edit of edits) {
    const address =
      typeof edit?.address === "string" ? edit.address.toUpperCase() : "";
    const cell = selection.cells.get(address);
    if (!cell || seen.has(address))
      throw new Error("AI 변경안에 범위 밖이나 중복된 셀이 있습니다.");
    if (cell.protected)
      throw new Error(
        `${address}는 수식, 병합 또는 서식 있는 텍스트 셀입니다. 일반 값만 있는 범위를 선택해 주세요.`,
      );
    seen.add(address);
    const after: unknown = edit.value;
    const formula = edit.kind === "formula" ? safeFormula(after) : null;
    const scalar =
      typeof after === "boolean" ||
      (typeof after === "number" && Number.isFinite(after)) ||
      (typeof after === "string" && after.length <= 2000);
    if (
      !(edit.kind === "set" && scalar) &&
      !(edit.kind === "clear" && after === null) &&
      !formula
    )
      throw new Error("지원하지 않는 셀 변경이 포함되어 있습니다.");
    if (
      formula || (!Object.is(cell.before, after) && !(cell.before == null && after === ""))
    )
      changes.push({ address, cell, after: formula ? null : after as Value, formula });
  }
  return changes;
}

export function applyAIChanges(
  editor: Editor,
  selection: AISelection,
  changes: Change[],
) {
  const workbook = editor.current();
  const active = workbook.getActiveSheet();
  const book = editor.snapshot();
  const sheet = book.sheets[selection.sheetId];
  if (
    book.id !== selection.bookId ||
    active.getSheetId() !== selection.sheetId ||
    !sheet ||
    sheet.name !== selection.sheetName ||
    sheet.rowCount !== selection.rows ||
    sheet.columnCount !== selection.cols
  )
    throw new Error(
      "문서나 시트 구조가 바뀌었습니다. 새 변경안을 요청해 주세요.",
    );
  for (const { row, col, fingerprint } of selection.cells.values()) {
    if (cellFingerprint(sheet.cellData?.[row]?.[col]) !== fingerprint)
      throw new Error(
        "선택 범위의 내용이 바뀌었습니다. 새 변경안을 요청해 주세요.",
      );
  }
  const patch: IObjectMatrixPrimitiveType<ICellData> = {};
  for (const {
    address,
    cell: { row, col },
    after,
    formula,
  } of changes) {
    if (protectedCell(sheet, row, col))
      throw new Error(
        `${address} 셀을 바꿀 수 없습니다. 새 변경안을 요청해 주세요.`,
      );
    (patch[row] ??= {})[col] = {
      v: after,
      t:
        after == null
          ? null
          : typeof after === "number"
            ? CellValueType.NUMBER
            : typeof after === "boolean"
              ? CellValueType.BOOLEAN
              : CellValueType.STRING,
      f: formula,
      p: null,
      si: null,
    };
  }
  const candidate = structuredClone(book);
  const targetData = (candidate.sheets[selection.sheetId].cellData ??= {});
  for (const {
    cell: { row, col },
  } of changes)
    (targetData[row] ??= {})[col] = {
      ...targetData[row]?.[col],
      ...patch[row][col],
    };
  validateBook(candidate);
  const { startRow, startColumn, endRow, endColumn } = selection.bounds;
  active
    .getRange(
      startRow,
      startColumn,
      endRow - startRow + 1,
      endColumn - startColumn + 1,
    )
    .setValues(patch);
  const updated = editor.snapshot().sheets[selection.sheetId].cellData;
  const resultMatches = ({ cell: { row, col }, after, formula }: Change) => {
    const actual = updated?.[row]?.[col];
    if (formula) return actual?.f === formula;
    if (actual?.f) return false;
    if (typeof after === "boolean")
      return (
        actual?.t === CellValueType.BOOLEAN &&
        (actual.v === after || actual.v === Number(after))
      );
    return (actual?.v ?? null) === after;
  };
  if (!changes.every(resultMatches)) {
    if (
      changes.some(
        ({ cell: { row, col, fingerprint } }) =>
          cellFingerprint(updated?.[row]?.[col]) !== fingerprint,
      )
    )
      workbook.undo();
    throw new Error(
      "선택한 셀에 변경안을 적용하지 못했습니다. 시트 보호와 범위를 확인해 주세요.",
    );
  }
}
