import type { IWorkbookData, ICellData } from "@univerjs/core";
import { LocaleType } from "@univerjs/core";

export const LIMITS = {
  bytes: 10 * 1024 * 1024,
  cells: 100000,
  rows: 10000,
  cols: 200,
  sheets: 30,
};
const record = (value: unknown): value is Record<string, unknown> =>
  !!value &&
  typeof value === "object" &&
  !Array.isArray(value) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(value));
const own = (value: object, key: string) =>
  Object.prototype.hasOwnProperty.call(value, key);
const safeKey = (key: string) =>
  !["__proto__", "prototype", "constructor"].includes(key);
const index = (key: string, max: number) =>
  /^(0|[1-9]\d*)$/.test(key) && +key < max;
const integer = (value: unknown, min: number, max: number) =>
  typeof value === "number" &&
  Number.isInteger(value) &&
  value >= min &&
  value <= max;
function validateJsonTree(value: unknown) {
  const seen = new WeakSet<object>();
  let nodes = 0,
    textLength = 0;
  function visit(item: unknown, depth: number) {
    if (++nodes > 1000000 || depth > 50)
      throw new Error("통합문서 구조가 너무 크거나 깊습니다.");
    if (typeof item === "string") {
      textLength += item.length;
      if (textLength > LIMITS.bytes)
        throw new Error("통합문서 내용은 10MB 이내로 줄여 주세요.");
      return;
    }
    if (item == null || typeof item === "boolean") return;
    if (typeof item === "number") {
      if (!Number.isFinite(item))
        throw new Error("유효하지 않은 숫자가 있습니다.");
      return;
    }
    if (typeof item !== "object" || (!Array.isArray(item) && !record(item)))
      throw new Error("통합문서에 지원하지 않는 데이터가 있습니다.");
    if (seen.has(item)) throw new Error("통합문서에 순환 참조가 있습니다.");
    seen.add(item);
    for (const [key, child] of Object.entries(item)) {
      if (!safeKey(key))
        throw new Error("안전하지 않은 통합문서 키가 있습니다.");
      textLength += key.length;
      visit(child, depth + 1);
    }
    seen.delete(item);
  }
  visit(value, 0);
  if (new TextEncoder().encode(JSON.stringify(value)).length > LIMITS.bytes)
    throw new Error("통합문서 백업 용량은 10MB 이내여야 합니다.");
}
export function newBook(example = false): IWorkbookData {
  const id = crypto.randomUUID(),
    sheetId = crypto.randomUUID();
  const cells: Record<number, Record<number, ICellData>> = {};
  if (example) {
    const rows = [
      ["날짜", "분류", "내용", "수입", "지출", "잔액"],
      ["2026-11-01", "수입", "예시 급여", 3000000, 0, "=D2-E2"],
      ["2026-11-02", "식비", "예시 장보기", 0, 65000, "=F2+D3-E3"],
      ["2026-11-03", "교통", "예시 교통비", 0, 55000, "=F3+D4-E4"],
      ["2026-11-04", "생활", "예시 통신비", 0, 45000, "=F4+D5-E5"],
      ["", "", "합계", "=SUM(D2:D5)", "=SUM(E2:E5)", "=D6-E6"],
    ];
    rows.forEach((row, r) => {
      cells[r] = {};
      row.forEach((v, c) => {
        cells[r][c] = {
          ...(typeof v === "string" && v.startsWith("=") ? { f: v } : { v }),
          s:
            r === 0
              ? { bl: 1, bg: { rgb: "#18794e" }, cl: { rgb: "#ffffff" } }
              : {
                  ...(c >= 3 ? { n: { pattern: "#,##0" } } : {}),
                  ...(r === 5 ? { bl: 1, bg: { rgb: "#eaf5ef" } } : {}),
                },
        };
      });
    });
  }
  return {
    id,
    name: example ? "예시 가계부" : "새 통합문서",
    appVersion: "1.3.1",
    locale: LocaleType.KO_KR,
    styles: {},
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: {
        id: sheetId,
        name: example ? "가계부" : "시트1",
        rowCount: 200,
        columnCount: 26,
        defaultColumnWidth: 120,
        defaultRowHeight: 30,
        cellData: cells,
        columnData: { 2: { w: 180 } },
      },
    },
  };
}
export function validateBook(value: unknown): asserts value is IWorkbookData {
  if (!record(value)) throw new Error("통합문서 형식이 올바르지 않습니다.");
  validateJsonTree(value);
  const b = value as unknown as IWorkbookData;
  if (
    typeof b.id !== "string" ||
    !b.id ||
    b.id.length > 200 ||
    !safeKey(b.id) ||
    typeof b.name !== "string" ||
    !b.name.trim() ||
    b.name.length > 200 ||
    !Array.isArray(b.sheetOrder) ||
    !b.sheetOrder.length ||
    b.sheetOrder.length > LIMITS.sheets ||
    !record(b.sheets) ||
    !record(b.styles) ||
    new Set(b.sheetOrder).size !== b.sheetOrder.length ||
    Object.keys(b.sheets).length !== b.sheetOrder.length
  )
    throw new Error("통합문서 구조나 시트 수를 확인해 주세요.");
  let count = 0;
  const sheetNames = new Set<string>();
  for (const id of b.sheetOrder) {
    if (
      typeof id !== "string" ||
      !id ||
      id.length > 200 ||
      !safeKey(id) ||
      !own(b.sheets, id)
    )
      throw new Error("시트 식별자가 올바르지 않습니다.");
    const s = b.sheets[id];
    if (
      !record(s) ||
      typeof s.name !== "string" ||
      !s.name.trim() ||
      s.name.length > 31 ||
      /[\[\]:*?\/\\]/.test(s.name) ||
      /^'|'$/.test(s.name) ||
      sheetNames.has(s.name.toLowerCase())
    )
      throw new Error("시트 이름은 중복 없이 1~31자로 지정해 주세요.");
    sheetNames.add(s.name.toLowerCase());
    if (
      !integer(s.rowCount, 1, LIMITS.rows) ||
      !integer(s.columnCount, 1, LIMITS.cols)
    )
      throw new Error("지원 범위는 시트당 10,000행·200열입니다.");
    if (s.cellData != null && !record(s.cellData))
      throw new Error("셀 데이터 구조가 올바르지 않습니다.");
    for (const [r, row] of Object.entries(s.cellData ?? {})) {
      if (!index(r, s.rowCount!) || !record(row))
        throw new Error("셀 주소가 올바르지 않습니다.");
      for (const [c, cell] of Object.entries(row)) {
        if (!index(c, s.columnCount!) || (cell != null && !record(cell)))
          throw new Error("셀 데이터가 올바르지 않습니다.");
        if (++count > LIMITS.cells)
          throw new Error("내용이 있는 셀은 최대 100,000개까지 지원합니다.");
        if (!cell) continue;
        if (
          cell.v != null &&
          !["string", "number", "boolean"].includes(typeof cell.v)
        )
          throw new Error("셀 값은 문자·숫자·논리값만 사용할 수 있습니다.");
        if (typeof cell.v === "string" && cell.v.length > 32767)
          throw new Error("셀 하나의 글자 수는 32,767자 이내여야 합니다.");
        if (
          cell.f != null &&
          cell.f !== "" &&
          (typeof cell.f !== "string" ||
            !cell.f.startsWith("=") ||
            cell.f.length > 8192)
        )
          throw new Error("수식 형식이나 길이를 확인해 주세요.");
        if (cell.s != null && typeof cell.s !== "string" && !record(cell.s))
          throw new Error("셀 서식이 올바르지 않습니다.");
      }
    }
    for (const [key, max, sizeKey] of [
      ["rowData", s.rowCount!, "h"],
      ["columnData", s.columnCount!, "w"],
    ] as const) {
      const dimensions = s[key];
      if (dimensions == null) continue;
      if (!record(dimensions))
        throw new Error("행·열 정보가 올바르지 않습니다.");
      for (const [key, dimension] of Object.entries(dimensions)) {
        if (!index(key, max) || (dimension != null && !record(dimension)))
          throw new Error("행·열 주소가 올바르지 않습니다.");
        if (
          dimension &&
          dimension[sizeKey] != null &&
          (typeof dimension[sizeKey] !== "number" ||
            dimension[sizeKey] < 0 ||
            dimension[sizeKey] > 10000)
        )
          throw new Error("행 높이·열 너비가 지원 범위를 넘습니다.");
      }
    }
    if (s.mergeData != null) {
      if (!Array.isArray(s.mergeData) || s.mergeData.length > LIMITS.cells)
        throw new Error("병합 정보가 올바르지 않습니다.");
      const merged = new Set<number>();
      for (const m of s.mergeData) {
        if (
          !record(m) ||
          !integer(m.startRow, 0, s.rowCount! - 1) ||
          !integer(m.endRow, m.startRow, s.rowCount! - 1) ||
          !integer(m.startColumn, 0, s.columnCount! - 1) ||
          !integer(m.endColumn, m.startColumn, s.columnCount! - 1)
        )
          throw new Error("병합 범위가 올바르지 않습니다.");
        if (
          (m.endRow - m.startRow + 1) * (m.endColumn - m.startColumn + 1) +
            merged.size >
          LIMITS.cells
        )
          throw new Error("병합한 전체 영역은 100,000셀 이내여야 합니다.");
        for (let row = m.startRow; row <= m.endRow; row++)
          for (let col = m.startColumn; col <= m.endColumn; col++) {
            const key = row * LIMITS.cols + col;
            if (merged.has(key))
              throw new Error("서로 겹치는 병합 범위가 있습니다.");
            merged.add(key);
          }
      }
    }
  }
}
export function cellText(
  cell: ICellData | null | undefined,
): string | number | boolean {
  if (!cell) return "";
  if (cell.p?.body?.dataStream)
    return cell.p.body.dataStream.replace(/[\r\n]+$/, "");
  return cell.v ?? "";
}
