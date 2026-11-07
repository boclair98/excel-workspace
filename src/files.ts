import ExcelJS from "exceljs";
import Papa from "papaparse";
import type {
  IWorkbookData,
  ICellData,
  IStyleData,
  IColorStyle,
  IRowData,
  IColumnData,
} from "@univerjs/core";
import { newBook, LIMITS, validateBook, cellText } from "./model";

const rgb = (c?: Partial<ExcelJS.Color>) =>
  c?.argb ? "#" + c.argb.slice(-6) : undefined;
const argb = (c?: IColorStyle | null | void) =>
  typeof c?.rgb === "string" && c.rgb.startsWith("#")
    ? "FF" + c.rgb.slice(1).toUpperCase()
    : undefined;
const borderStyles: ExcelJS.BorderStyle[] = [
  "thin",
  "thin",
  "hair",
  "dotted",
  "dashed",
  "dashDot",
  "dashDotDot",
  "double",
  "medium",
  "mediumDashed",
  "mediumDashDot",
  "mediumDashDotDot",
  "slantDashDot",
  "thick",
];
const borderSides = { top: "t", right: "r", bottom: "b", left: "l" } as const;
function fromBorders(cell: ExcelJS.Cell): IStyleData["bd"] {
  const result: NonNullable<IStyleData["bd"]> = {};
  for (const [side, key] of Object.entries(borderSides)) {
    const border = cell.border?.[side as keyof typeof borderSides];
    if (border?.style)
      result[key] = {
        s: Math.max(1, borderStyles.indexOf(border.style)),
        cl: { rgb: rgb(border.color) ?? "#000000" },
      };
  }
  return Object.keys(result).length ? result : undefined;
}
// Inspect bounded ZIP metadata before ExcelJS allocates expanded XML.
function validateXlsxArchive(bytes: ArrayBuffer) {
  const view = ArrayBuffer.isView(bytes)
    ? new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    : new DataView(bytes);
  let end = -1;
  for (
    let pos = bytes.byteLength - 22;
    pos >= Math.max(0, bytes.byteLength - 65557);
    pos--
  )
    if (
      view.getUint32(pos, true) === 0x06054b50 &&
      pos + 22 + view.getUint16(pos + 20, true) === bytes.byteLength
    ) {
      end = pos;
      break;
    }
  if (end < 0) throw new Error("올바른 XLSX 압축 파일이 아닙니다.");
  const entries = view.getUint16(end + 10, true),
    directorySize = view.getUint32(end + 12, true),
    start = view.getUint32(end + 16, true);
  if (
    view.getUint16(end + 4, true) ||
    view.getUint16(end + 6, true) ||
    !entries ||
    entries > 2000 ||
    start + directorySize > end
  )
    throw new Error("XLSX 압축 구조 또는 파일 수가 지원 범위를 넘습니다.");
  let position = start,
    expanded = 0;
  for (let i = 0; i < entries; i++) {
    if (
      position + 46 > start + directorySize ||
      view.getUint32(position, true) !== 0x02014b50
    )
      throw new Error("XLSX 압축 정보가 손상되었습니다.");
    if (view.getUint16(position + 8, true) & 1)
      throw new Error("암호화된 파일은 지원하지 않습니다.");
    expanded += view.getUint32(position + 24, true);
    if (expanded > 64 * 1024 * 1024)
      throw new Error(
        "압축을 푼 내용이 64MB를 넘습니다. 불필요한 시트·서식을 줄여 주세요.",
      );
    position +=
      46 +
      view.getUint16(position + 28, true) +
      view.getUint16(position + 30, true) +
      view.getUint16(position + 32, true);
    if (position > start + directorySize)
      throw new Error("XLSX 압축 정보가 손상되었습니다.");
  }
}
function fromStyle(cell: ExcelJS.Cell): IStyleData {
  const f = cell.font ?? {},
    a = cell.alignment ?? {};
  const bg = cell.fill?.type === "pattern" ? rgb(cell.fill.fgColor) : undefined;
  return {
    bl: f.bold ? 1 : 0,
    it: f.italic ? 1 : 0,
    fs: f.size,
    ff: f.name,
    ...(f.underline ? { ul: { s: 1 } } : {}),
    ...(f.strike ? { st: { s: 1 } } : {}),
    ...(rgb(f.color) ? { cl: { rgb: rgb(f.color) } } : {}),
    ...(bg ? { bg: { rgb: bg } } : {}),
    ...(cell.numFmt ? { n: { pattern: cell.numFmt } } : {}),
    bd: fromBorders(cell),
    ht:
      a.horizontal === "center"
        ? 2
        : a.horizontal === "right"
          ? 3
          : a.horizontal === "left"
            ? 1
            : 0,
    vt: a.vertical === "top" ? 1 : a.vertical === "bottom" ? 3 : 2,
    tb: a.wrapText ? 3 : 1,
  };
}
export async function importXlsx(
  bytes: ArrayBuffer,
  name: string,
): Promise<IWorkbookData> {
  if (bytes.byteLength > LIMITS.bytes)
    throw new Error("파일은 10MB 이하로 선택해 주세요.");
  validateXlsxArchive(bytes);
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(bytes);
  if (!x.worksheets.length || x.worksheets.length > LIMITS.sheets)
    throw new Error("시트는 1~30개까지 지원합니다.");
  const book = newBook();
  book.name =
    name
      .replace(/\.xlsx$/i, "")
      .trim()
      .slice(0, 200) || "가져온 통합문서";
  book.sheets = {};
  book.sheetOrder = [];
  let count = 0;
  for (const ws of x.worksheets) {
    if (ws.rowCount > LIMITS.rows || ws.columnCount > LIMITS.cols)
      throw new Error("지원 범위는 시트당 10,000행·200열입니다.");
    const id = crypto.randomUUID();
    const cellData: Record<number, Record<number, ICellData>> = {};
    const rowData: Record<number, IRowData> = {},
      columnData: Record<number, IColumnData> = {};
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      if (row.height || row.hidden)
        rowData[r - 1] = {
          ...(row.height ? { h: (row.height * 4) / 3 } : {}),
          ...(row.hidden ? { hd: 1 } : {}),
        };
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        if (cell.value == null && !Object.keys(cell.style ?? {}).length) return;
        if (++count > LIMITS.cells)
          throw new Error("내용이 있는 셀은 최대 100,000개까지 지원합니다.");
        if (cell.isMerged && cell.master.address !== cell.address) return;
        let v: ICellData["v"] = null;
        let f: string | undefined;
        const raw = cell.value;
        // Univer and exported workbooks use the 1900 date system. ExcelJS has
        // already converted a source 1904 date into the correct absolute Date.
        if (raw instanceof Date) v = raw.getTime() / 86400000 + 25569;
        else if (raw && typeof raw === "object") {
          if ("formula" in raw || "sharedFormula" in raw) {
            f = "=" + cell.formula;
            v =
              cell.result instanceof Date
                ? cell.result.getTime() / 86400000 + 25569
                : typeof cell.result === "object"
                  ? String((cell.result as { error?: string })?.error ?? "")
                  : (cell.result ?? null);
          } else if ("richText" in raw)
            v = raw.richText.map((t) => t.text).join("");
          else if ("text" in raw) v = raw.text;
          else if ("error" in raw) v = raw.error;
        } else v = raw ?? null;
        (cellData[r - 1] ??= {})[c - 1] = {
          v,
          ...(f ? { f } : {}),
          ...(typeof v === "string"
            ? { t: 1 }
            : typeof v === "boolean"
              ? { t: 3 }
              : typeof v === "number"
                ? { t: 2 }
                : {}),
          s: fromStyle(cell),
        };
      });
    });
    ws.columns.forEach((c, i) => {
      if (c.width || c.hidden)
        columnData[i] = {
          ...(c.width ? { w: Math.round(c.width * 7 + 5) } : {}),
          ...(c.hidden ? { hd: 1 } : {}),
        };
    });
    const mergeData = (ws.model.merges ?? []).map((ref) => {
      const parse = (s: string) => {
        const m = /^([A-Z]+)(\d+)$/.exec(s)!;
        return {
          row: +m[2] - 1,
          col: [...m[1]].reduce((v, c) => v * 26 + c.charCodeAt(0) - 64, 0) - 1,
        };
      };
      const [a, b] = ref.split(":").map(parse);
      return {
        startRow: a.row,
        endRow: b.row,
        startColumn: a.col,
        endColumn: b.col,
      };
    });
    book.sheetOrder.push(id);
    book.sheets[id] = {
      id,
      name: ws.name,
      tabColor: /^#[\da-f]{6}$/i.test(rgb(ws.properties.tabColor) ?? "")
        ? rgb(ws.properties.tabColor)!
        : "",
      hidden: ws.state === "veryHidden" ? 2 : ws.state === "hidden" ? 1 : 0,
      rowCount: Math.max(200, ws.rowCount + 20),
      columnCount: Math.max(26, ws.columnCount + 2),
      defaultRowHeight: 30,
      defaultColumnWidth: 120,
      cellData,
      rowData,
      columnData,
      mergeData,
    };
    book.sheets[id].rowCount = Math.min(LIMITS.rows, book.sheets[id].rowCount!);
    book.sheets[id].columnCount = Math.min(
      LIMITS.cols,
      book.sheets[id].columnCount!,
    );
  }
  validateBook(book);
  return book;
}
export function importCsv(text: string, name: string): IWorkbookData {
  if (new TextEncoder().encode(text).length > LIMITS.bytes)
    throw new Error("CSV는 10MB 이하로 선택해 주세요.");
  const parsed = Papa.parse<string[]>(text.replace(/^\uFEFF/, ""), {
    skipEmptyLines: "greedy",
  });
  if (parsed.errors.some((e) => e.code !== "UndetectableDelimiter"))
    throw new Error("CSV의 따옴표나 구분자를 확인해 주세요.");
  const book = newBook();
  book.name =
    name
      .replace(/\.csv$/i, "")
      .trim()
      .slice(0, 200) || "가져온 CSV";
  const sheet = book.sheets[book.sheetOrder[0]];
  sheet.cellData = {};
  let count = 0;
  parsed.data.forEach((row, r) => {
    if (
      r >= LIMITS.rows ||
      row.length > LIMITS.cols ||
      (count += row.length) > LIMITS.cells
    )
      throw new Error("CSV의 행·열 또는 셀 수가 지원 범위를 넘습니다.");
    sheet.cellData![r] = {};
    row.forEach((v, c) => {
      sheet.cellData![r][c] = {
        v:
          /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(v) && v.length < 16
            ? Number(v)
            : v,
        t: /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(v) && v.length < 16 ? 2 : 1,
      };
    });
  });
  sheet.rowCount = Math.min(
    LIMITS.rows,
    Math.max(200, parsed.data.length + 20),
  );
  sheet.columnCount = Math.min(
    LIMITS.cols,
    Math.max(26, ...parsed.data.map((r) => r.length)),
  );
  validateBook(book);
  return book;
}
export async function exportXlsx(book: IWorkbookData): Promise<ArrayBuffer> {
  validateBook(book);
  const x = new ExcelJS.Workbook();
  x.creator = "오픈소스 엑셀";
  x.calcProperties.fullCalcOnLoad = true;
  for (const id of book.sheetOrder) {
    const sheet = book.sheets[id],
      ws = x.addWorksheet(sheet.name ?? "시트");
    ws.state =
      sheet.hidden === 2
        ? "veryHidden"
        : sheet.hidden === 1
          ? "hidden"
          : "visible";
    if (/^#[\da-f]{6}$/i.test(sheet.tabColor ?? ""))
      ws.properties.tabColor = {
        argb: "FF" + sheet.tabColor!.slice(1).toUpperCase(),
      };
    for (const [r, row] of Object.entries(sheet.cellData ?? {}))
      for (const [c, data] of Object.entries(
        row as Record<string, ICellData>,
      )) {
        if (!data) continue;
        const cell = ws.getCell(+r + 1, +c + 1);
        const value = cellText(data);
        cell.value = data.f
          ? {
              formula: data.f.replace(/^=/, ""),
              ...(data.v != null ? { result: value } : {}),
            }
          : value;
        const s =
          (typeof data.s === "string" ? book.styles[data.s] : data.s) ?? {};
        cell.font = {
          name: s.ff ?? "Arial",
          size: s.fs ?? 11,
          bold: s.bl === 1,
          italic: s.it === 1,
          underline: s.ul?.s === 1,
          strike: s.st?.s === 1,
          ...(argb(s.cl) ? { color: { argb: argb(s.cl) } } : {}),
        };
        if (argb(s.bg))
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: argb(s.bg) },
          };
        if (s.n?.pattern) cell.numFmt = s.n.pattern;
        cell.alignment = {
          horizontal:
            s.ht === 2
              ? "center"
              : s.ht === 3
                ? "right"
                : s.ht === 1
                  ? "left"
                  : undefined,
          vertical: s.vt === 1 ? "top" : s.vt === 3 ? "bottom" : "middle",
          wrapText: s.tb === 3,
        };
        for (const [side, key] of Object.entries(borderSides)) {
          const border = s.bd?.[key];
          if (border?.s)
            cell.border = {
              ...cell.border,
              [side]: {
                style: borderStyles[border.s] ?? "thin",
                color: { argb: argb(border.cl) ?? "FF000000" },
              },
            };
        }
      }
    for (const [c, col] of Object.entries(sheet.columnData ?? {})) {
      if (col?.w) ws.getColumn(+c + 1).width = Math.max(1, (col.w - 5) / 7);
      if (col?.hd) ws.getColumn(+c + 1).hidden = true;
    }
    for (const [r, row] of Object.entries(sheet.rowData ?? {})) {
      if (row?.h) ws.getRow(+r + 1).height = (row.h * 3) / 4;
      if (row?.hd) ws.getRow(+r + 1).hidden = true;
    }
    for (const m of sheet.mergeData ?? [])
      ws.mergeCells(
        m.startRow + 1,
        m.startColumn + 1,
        m.endRow + 1,
        m.endColumn + 1,
      );
  }
  const bytes = await x.xlsx.writeBuffer();
  return ArrayBuffer.isView(bytes)
    ? Uint8Array.from(bytes as unknown as Uint8Array).buffer
    : (bytes as ArrayBuffer);
}
export function exportCsv(book: IWorkbookData, sheetId: string): string {
  validateBook(book);
  const s = book.sheets[sheetId];
  if (!s) throw new Error("시트를 선택해 주세요.");
  const data = s.cellData ?? {};
  let maxRow = 0,
    maxCol = 0;
  for (const [r, row] of Object.entries(data))
    for (const [c, value] of Object.entries(row as Record<string, ICellData>))
      if (value && (value.v != null || value.f || value.p)) {
        maxRow = Math.max(maxRow, +r);
        maxCol = Math.max(maxCol, +c);
      }
  if ((maxRow + 1) * (maxCol + 1) > 1000000)
    throw new Error("CSV 범위가 너무 큽니다. 불필요한 먼 셀을 지워 주세요.");
  const rows = Array.from({ length: maxRow + 1 }, (_, r) =>
    Array.from({ length: maxCol + 1 }, (_, c) => cellText(data[r]?.[c])),
  );
  return "\uFEFF" + Papa.unparse(rows, { escapeFormulae: true });
}
export function download(data: BlobPart, name: string, type: string) {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name.replace(/[<>:"/\\|?*]/g, "_");
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
