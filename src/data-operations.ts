import type { ICellData, IObjectMatrixPrimitiveType } from "@univerjs/core";

export type DataMatrix = IObjectMatrixPrimitiveType<ICellData> | undefined;
export type DataBounds = {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
};
export type FindOptions = {
  caseSensitive?: boolean;
  wholeCell?: boolean;
  bounds?: DataBounds;
};
export type CellMatch = {
  row: number;
  col: number;
  address: string;
  text: string;
  editable: boolean;
  protectedReason: string;
  fingerprint: string;
};
export type TextChange = Pick<
  CellMatch,
  "row" | "col" | "address" | "fingerprint"
> & { before: string; after: string };
export type ChangePlan = {
  changes: TextChange[];
  protectedCount: number;
  matchedCount: number;
};
const MAX_ROWS = 10000,
  MAX_COLS = 200,
  MAX_CELLS = 100000,
  MAX_CELL_TEXT = 32767,
  MAX_CHANGE_BYTES = 10 * 1024 * 1024;
const own = (object: object, key: string | number) =>
  Object.prototype.hasOwnProperty.call(object, key);

function validCoordinate(row: number, col: number) {
  if (
    !Number.isInteger(row) ||
    row < 0 ||
    row >= MAX_ROWS ||
    !Number.isInteger(col) ||
    col < 0 ||
    col >= MAX_COLS
  )
    throw new Error("셀 주소가 지원 범위를 벗어났습니다.");
}
function validateBounds(bounds?: DataBounds) {
  if (!bounds) return;
  validCoordinate(bounds.startRow, bounds.startColumn);
  validCoordinate(bounds.endRow, bounds.endColumn);
  if (bounds.startRow > bounds.endRow || bounds.startColumn > bounds.endColumn)
    throw new Error("선택 범위를 다시 확인해 주세요.");
}
export function cellAddress(row: number, col: number): string {
  validCoordinate(row, col);
  let label = "",
    n = col + 1;
  while (n > 0) {
    label = String.fromCharCode(65 + ((n - 1) % 26)) + label;
    n = Math.floor((n - 1) / 26);
  }
  return `${label}${row + 1}`;
}
function reason(cell: ICellData): string {
  if (cell.f || cell.si || cell.ref || cell.ft != null || cell.fd === 1)
    return "수식 셀";
  if (cell.p != null) return "서식 있는 텍스트";
  if (typeof cell.v === "number" || cell.t === 2) return "숫자 셀";
  if (typeof cell.v === "boolean" || cell.t === 3) return "논리값 셀";
  if (typeof cell.v !== "string") return "텍스트가 아닌 셀";
  return "";
}
export function cellFingerprint(cell: ICellData | null | undefined): string {
  return JSON.stringify([
    cell?.v,
    cell?.t,
    cell?.f,
    cell?.si,
    cell?.p,
    cell?.ft,
    cell?.fd,
    cell?.ref,
    cell?.xf,
  ]);
}
function displayText(cell: ICellData): string {
  if (cell.p?.body?.dataStream)
    return cell.p.body.dataStream.replace(/[\r\n]+$/, "");
  return cell.v == null ? "" : String(cell.v);
}
function collect(matrix: DataMatrix, bounds?: DataBounds): CellMatch[] {
  validateBounds(bounds);
  if (!matrix) return [];
  const result: CellMatch[] = [];
  let count = 0;
  // Visit populated cells only; a sheet-wide search never expands a dense blank matrix.
  for (const rowKey of Object.keys(matrix)) {
    if (!/^(0|[1-9]\d*)$/.test(rowKey))
      throw new Error("셀 주소가 올바르지 않습니다.");
    const row = Number(rowKey),
      rowData = matrix[row];
    validCoordinate(row, 0);
    if (!rowData || typeof rowData !== "object" || Array.isArray(rowData))
      throw new Error("셀 데이터가 올바르지 않습니다.");
    for (const colKey of Object.keys(rowData)) {
      if (++count > MAX_CELLS)
        throw new Error("찾기·정리는 100,000개 셀까지 지원합니다.");
      if (!/^(0|[1-9]\d*)$/.test(colKey))
        throw new Error("셀 주소가 올바르지 않습니다.");
      const col = Number(colKey);
      validCoordinate(row, col);
      if (
        bounds &&
        (row < bounds.startRow ||
          row > bounds.endRow ||
          col < bounds.startColumn ||
          col > bounds.endColumn)
      )
        continue;
      const cell = rowData[col];
      if (!cell) continue;
      const text = displayText(cell);
      if (!text) continue;
      const protectedReason = reason(cell);
      result.push({
        row,
        col,
        address: cellAddress(row, col),
        text,
        editable: !protectedReason,
        protectedReason,
        fingerprint: cellFingerprint(cell),
      });
    }
  }
  return result.sort((a, b) => a.row - b.row || a.col - b.col);
}
function matcher(query: string, options: FindOptions) {
  if (!query || query.length > 500)
    throw new Error("찾을 내용을 1~500자로 입력해 주세요.");
  const literal = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(
    options.wholeCell ? `^(?:${literal})(?![\\s\\S])` : literal,
    options.caseSensitive ? "g" : "gi",
  );
}
export function findCells(
  matrix: DataMatrix,
  query: string,
  options: FindOptions = {},
): CellMatch[] {
  const regex = matcher(query, options);
  return collect(matrix, options.bounds).filter((cell) => {
    regex.lastIndex = 0;
    return regex.test(cell.text);
  });
}
function replaceLiteral(
  text: string,
  query: string,
  replacement: string,
  options: FindOptions,
): string {
  const regex = matcher(query, options);
  let outputLength = 0,
    previous = 0;
  const parts: string[] = [];
  for (let match = regex.exec(text); match; match = regex.exec(text)) {
    outputLength += match.index - previous + replacement.length;
    if (outputLength > MAX_CELL_TEXT)
      throw new Error(
        "바꾼 뒤의 셀 내용이 32,767자를 넘습니다. 바꿀 내용을 줄여 주세요.",
      );
    parts.push(text.slice(previous, match.index), replacement);
    previous = match.index + match[0].length;
  }
  outputLength += text.length - previous;
  if (outputLength > MAX_CELL_TEXT)
    throw new Error(
      "바꾼 뒤의 셀 내용이 32,767자를 넘습니다. 바꿀 내용을 줄여 주세요.",
    );
  parts.push(text.slice(previous));
  return parts.join("");
}
export function planReplacement(
  matches: CellMatch[],
  query: string,
  replacement: string,
  options: FindOptions = {},
): ChangePlan {
  matcher(query, options);
  if (replacement.length > 2000)
    throw new Error("바꿀 내용은 2,000자 이내로 입력해 주세요.");
  if (matches.length > MAX_CELLS)
    throw new Error("한 번에 바꿀 수 있는 셀 수를 넘었습니다.");
  const changes: TextChange[] = [];
  let protectedCount = 0,
    changedBytes = 0;
  for (const match of matches) {
    if (!match.editable) {
      protectedCount++;
      continue;
    }
    const after = replaceLiteral(match.text, query, replacement, options);
    if (after !== match.text) {
      changedBytes += new TextEncoder().encode(after).byteLength;
      if (changedBytes > MAX_CHANGE_BYTES)
        throw new Error(
          "바뀌는 텍스트가 10MB를 넘습니다. 선택 범위나 바꿀 내용을 줄여 주세요.",
        );
      changes.push({
        row: match.row,
        col: match.col,
        address: match.address,
        fingerprint: match.fingerprint,
        before: match.text,
        after,
      });
    }
  }
  return { changes, protectedCount, matchedCount: matches.length };
}
export function planTrim(matrix: DataMatrix, bounds?: DataBounds): ChangePlan {
  const cells = collect(matrix, bounds),
    changes: TextChange[] = [];
  let protectedCount = 0;
  for (const cell of cells) {
    if (!cell.editable) {
      protectedCount++;
      continue;
    }
    const after = cell.text.trim();
    if (after !== cell.text)
      changes.push({
        row: cell.row,
        col: cell.col,
        address: cell.address,
        fingerprint: cell.fingerprint,
        before: cell.text,
        after,
      });
  }
  return { changes, protectedCount, matchedCount: cells.length };
}
export function createSparseTextPatch(
  matrix: DataMatrix,
  changes: TextChange[],
): { values: IObjectMatrixPrimitiveType<ICellData>; bounds: DataBounds } {
  if (!changes.length || changes.length > MAX_CELLS)
    throw new Error("바꿀 셀을 다시 확인해 주세요.");
  const seen = new Set<string>();
  const bounds = {
    startRow: MAX_ROWS,
    endRow: 0,
    startColumn: MAX_COLS,
    endColumn: 0,
  };
  // Validate the entire plan before producing a single atomic editor command.
  for (const change of changes) {
    validCoordinate(change.row, change.col);
    const key = `${change.row}:${change.col}`;
    if (seen.has(key)) throw new Error("변경 대상에 중복된 셀이 있습니다.");
    seen.add(key);
    const row =
      matrix && own(matrix, change.row) ? matrix[change.row] : undefined;
    const current = row && own(row, change.col) ? row[change.col] : undefined;
    if (
      !current ||
      reason(current) ||
      cellFingerprint(current) !== change.fingerprint ||
      current.v !== change.before
    )
      throw new Error(
        `${cellAddress(change.row, change.col)}의 내용이 변경되었습니다. 다시 미리보기한 뒤 적용해 주세요.`,
      );
    if (typeof change.after !== "string" || change.after.length > MAX_CELL_TEXT)
      throw new Error("바꿀 셀 내용이 너무 길거나 올바르지 않습니다.");
    bounds.startRow = Math.min(bounds.startRow, change.row);
    bounds.endRow = Math.max(bounds.endRow, change.row);
    bounds.startColumn = Math.min(bounds.startColumn, change.col);
    bounds.endColumn = Math.max(bounds.endColumn, change.col);
  }
  const values: IObjectMatrixPrimitiveType<ICellData> = {};
  for (const change of changes)
    (values[change.row] ??= {})[change.col] = {
      v: change.after,
      t: 1,
      f: null,
      si: null,
      p: null,
    };
  return { values, bounds };
}
