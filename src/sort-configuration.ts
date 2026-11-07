import type { ICellData, IWorksheetData } from "@univerjs/core";
import { cellAddress, type DataBounds } from "./data-operations";

export type SortCriterion = { column: number; ascending: boolean };
export type SortConfiguration = {
  selected: DataBounds;
  dataRange: DataBounds;
  hasHeader: boolean;
  columns: { column: number; absoluteColumn: number; label: string }[];
  criteria: SortCriterion[];
  rows: number;
  width: number;
  populatedRows: number;
};
const MAX_SORT_CELLS = 100000;
const hasValue = (cell: ICellData | null | undefined) =>
  (cell?.v != null && cell.v !== "") ||
  Boolean(cell?.p?.body?.dataStream?.replace(/[\r\n]+$/, ""));
const intersects = (a: DataBounds, b: DataBounds) =>
  a.startRow <= b.endRow &&
  a.endRow >= b.startRow &&
  a.startColumn <= b.endColumn &&
  a.endColumn >= b.startColumn;
export function rangeLabel(range: DataBounds): string {
  return `${cellAddress(range.startRow, range.startColumn)}:${cellAddress(range.endRow, range.endColumn)}`;
}
export function validateSortRange(
  sheet: Partial<IWorksheetData>,
  selected: DataBounds,
): void {
  const keys = ["startRow", "endRow", "startColumn", "endColumn"] as const;
  if (
    keys.some((key) => !Number.isInteger(selected[key]) || selected[key] < 0) ||
    selected.startRow > selected.endRow ||
    selected.startColumn > selected.endColumn ||
    selected.endRow >= Math.min(10000, sheet.rowCount ?? 0) ||
    selected.endColumn >= Math.min(200, sheet.columnCount ?? 0)
  )
    throw new Error("정렬할 범위를 다시 선택해 주세요.");
  if (selected.startRow === selected.endRow)
    throw new Error("정렬할 데이터 행을 두 줄 이상 함께 선택해 주세요.");
  if (
    (selected.endRow - selected.startRow + 1) *
      (selected.endColumn - selected.startColumn + 1) >
    MAX_SORT_CELLS
  )
    throw new Error(
      "정렬은 한 번에 100,000개 셀까지 지원합니다. 범위를 줄여 주세요.",
    );
  if (sheet.mergeData?.some((merge) => intersects(merge, selected)))
    throw new Error(
      "선택 범위에 병합된 셀이 있습니다. 병합을 해제하거나 다른 범위를 선택해 주세요.",
    );
  for (let row = selected.startRow; row <= selected.endRow; row++) {
    if (sheet.rowData?.[row]?.hd === 1)
      throw new Error(
        "선택 범위에 숨겨진 행이 있습니다. 행 숨김을 해제한 뒤 정렬해 주세요.",
      );
    for (let col = selected.startColumn; col <= selected.endColumn; col++) {
      const cell = sheet.cellData?.[row]?.[col];
      if (cell?.ref || cell?.fd === 1 || cell?.ft === 2 || cell?.ft === 3)
        throw new Error(
          "선택 범위에 배열 수식이나 데이터 표 수식이 있습니다. 해당 수식 밖의 범위를 선택해 주세요.",
        );
    }
  }
}
export function sortColumns(
  sheet: Partial<IWorksheetData>,
  selected: DataBounds,
  hasHeader: boolean,
) {
  validateSortRange(sheet, selected);
  return Array.from(
    { length: selected.endColumn - selected.startColumn + 1 },
    (_, column) => {
      const absoluteColumn = selected.startColumn + column,
        cell = sheet.cellData?.[selected.startRow]?.[absoluteColumn];
      const name =
        cell?.p?.body?.dataStream?.replace(/[\r\n]+$/, "") ||
        (cell?.v == null ? "" : String(cell.v));
      const letter = cellAddress(0, absoluteColumn).replace(/\d+$/, "");
      return {
        column,
        absoluteColumn,
        label:
          hasHeader && name.trim()
            ? `${letter}열 · ${name.slice(0, 60)}${name.length > 60 ? "…" : ""}`
            : `${letter}열`,
      };
    },
  );
}
export function createSortConfiguration(
  sheet: Partial<IWorksheetData>,
  selected: DataBounds,
  hasHeader: boolean,
  criteria: SortCriterion[],
): SortConfiguration {
  validateSortRange(sheet, selected);
  const dataRange = {
      ...selected,
      startRow: selected.startRow + (hasHeader ? 1 : 0),
    },
    width = selected.endColumn - selected.startColumn + 1;
  if (dataRange.endRow - dataRange.startRow + 1 < 2)
    throw new Error(
      "제목 행을 제외한 데이터가 두 줄 이상 필요합니다. 제목이 없다면 ‘첫 행은 제목’을 꺼 주세요.",
    );
  if (!criteria.length || criteria.length > Math.min(3, width))
    throw new Error("정렬 기준은 1~3개까지 선택할 수 있습니다.");
  const seen = new Set<number>();
  for (const criterion of criteria) {
    if (
      !Number.isInteger(criterion.column) ||
      criterion.column < 0 ||
      criterion.column >= width ||
      typeof criterion.ascending !== "boolean"
    )
      throw new Error("정렬 기준 열이나 방향을 다시 확인해 주세요.");
    if (seen.has(criterion.column))
      throw new Error("같은 열을 정렬 기준으로 중복 선택할 수 없습니다.");
    seen.add(criterion.column);
  }
  let populatedRows = 0,
    keyValues = 0;
  for (let row = dataRange.startRow; row <= dataRange.endRow; row++) {
    let filled = false;
    for (let col = dataRange.startColumn; col <= dataRange.endColumn; col++)
      if (hasValue(sheet.cellData?.[row]?.[col])) filled = true;
    if (filled) populatedRows++;
    if (
      criteria.some((criterion) =>
        hasValue(
          sheet.cellData?.[row]?.[selected.startColumn + criterion.column],
        ),
      )
    )
      keyValues++;
  }
  if (!populatedRows)
    throw new Error("선택한 범위에 정렬할 데이터가 없습니다.");
  if (!keyValues)
    throw new Error(
      "선택한 기준 열에 정렬할 값이 없습니다. 다른 기준 열을 선택해 주세요.",
    );
  return {
    selected: { ...selected },
    dataRange,
    hasHeader,
    columns: sortColumns(sheet, selected, hasHeader),
    criteria: criteria.map((criterion) => ({ ...criterion })),
    rows: dataRange.endRow - dataRange.startRow + 1,
    width,
    populatedRows,
  };
}
export function sortRangeFingerprint(
  sheet: Partial<IWorksheetData>,
  selected: DataBounds,
): string {
  validateSortRange(sheet, selected);
  const cells: unknown[] = [];
  for (let row = selected.startRow; row <= selected.endRow; row++) {
    for (let col = selected.startColumn; col <= selected.endColumn; col++) {
      const cell = sheet.cellData?.[row]?.[col];
      if (cell && Object.keys(cell).length) cells.push([row, col, cell]);
    }
  }
  return JSON.stringify([
    sheet.id,
    sheet.name,
    sheet.rowCount,
    sheet.columnCount,
    selected,
    cells,
  ]);
}
export function assertSortUnchanged(
  sheet: Partial<IWorksheetData>,
  selected: DataBounds,
  fingerprint: string,
): void {
  if (sortRangeFingerprint(sheet, selected) !== fingerprint)
    throw new Error(
      "정렬 창을 연 뒤 데이터나 시트 구조가 바뀌었습니다. 창을 닫고 범위를 다시 선택해 주세요.",
    );
}
