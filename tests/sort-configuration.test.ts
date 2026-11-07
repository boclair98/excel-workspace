import test from "node:test";
import assert from "node:assert/strict";
import type { IWorksheetData } from "@univerjs/core";
import {
  assertSortUnchanged,
  createSortConfiguration,
  rangeLabel,
  sortColumns,
  sortRangeFingerprint,
  validateSortRange,
} from "../src/sort-configuration";

function sheet(): Partial<IWorksheetData> {
  return {
    id: "sheet-1",
    name: "정렬 표",
    rowCount: 200,
    columnCount: 26,
    cellData: {
      1: { 1: { v: "부서" }, 2: { v: "점수" }, 3: { v: "이름" } },
      2: { 1: { v: "개발" }, 2: { v: 80 }, 3: { v: "첫 번째" } },
      3: { 1: { v: "기획" }, 2: { v: 95 }, 3: { v: "두 번째" } },
      4: { 1: { v: "개발" }, 2: { v: 90 }, 3: { v: "세 번째" } },
    },
  };
}
const selected = { startRow: 1, endRow: 4, startColumn: 1, endColumn: 3 };
test("non-A range retains relative criterion columns and excludes the header row", () => {
  const original = sheet(),
    before = structuredClone(original);
  const config = createSortConfiguration(original, selected, true, [
    { column: 0, ascending: true },
    { column: 1, ascending: false },
  ]);
  assert.deepEqual(config.criteria, [
    { column: 0, ascending: true },
    { column: 1, ascending: false },
  ]);
  assert.equal(config.columns[0].absoluteColumn, 1);
  assert.equal(config.columns[1].label, "C열 · 점수");
  assert.equal(rangeLabel(config.dataRange), "B3:D5");
  assert.equal(config.rows, 3);
  assert.equal(config.width, 3);
  assert.equal(config.populatedRows, 3);
  assert.deepEqual(original, before);
  assert.equal(
    rangeLabel(
      createSortConfiguration(original, selected, false, [
        { column: 2, ascending: false },
      ]).dataRange,
    ),
    "B2:D5",
  );
});
test("criteria limits, duplicate columns and invalid directions are rejected", () => {
  assert.throws(() => createSortConfiguration(sheet(), selected, true, []));
  assert.throws(() =>
    createSortConfiguration(sheet(), selected, true, [
      { column: 0, ascending: true },
      { column: 0, ascending: false },
    ]),
  );
  assert.throws(() =>
    createSortConfiguration(sheet(), selected, true, [
      { column: 3, ascending: true },
    ]),
  );
  assert.throws(() =>
    createSortConfiguration(sheet(), selected, true, [
      { column: 1, ascending: "asc" as unknown as boolean },
    ]),
  );
  assert.throws(() =>
    createSortConfiguration(
      sheet(),
      selected,
      true,
      Array.from({ length: 4 }, (_, column) => ({ column, ascending: true })),
    ),
  );
});
test("empty, single-row, oversized, invalid and merged ranges fail before engine mutation", () => {
  assert.throws(() => validateSortRange(sheet(), { ...selected, endRow: 1 }));
  assert.throws(() =>
    createSortConfiguration(sheet(), { ...selected, endRow: 2 }, true, [
      { column: 0, ascending: true },
    ]),
  );
  assert.throws(
    () =>
      createSortConfiguration(
        sheet(),
        { startRow: 10, endRow: 15, startColumn: 1, endColumn: 3 },
        true,
        [{ column: 0, ascending: true }],
      ),
    /데이터가 없습니다/,
  );
  assert.throws(
    () =>
      validateSortRange(
        { ...sheet(), rowCount: 10000, columnCount: 200 },
        { startRow: 0, endRow: 9999, startColumn: 0, endColumn: 199 },
      ),
    /100,000/,
  );
  assert.throws(() =>
    validateSortRange(sheet(), { ...selected, startColumn: -1 }),
  );
  assert.throws(
    () =>
      validateSortRange(
        {
          ...sheet(),
          mergeData: [{ startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 }],
        },
        selected,
      ),
    /병합/,
  );
});
test("hidden rows, array formulas and empty key columns are blocked", () => {
  assert.throws(
    () =>
      validateSortRange({ ...sheet(), rowData: { 3: { hd: 1 } } }, selected),
    /숨겨진/,
  );
  const array = sheet();
  array.cellData![3][2] = { f: "=SEQUENCE(2)", ref: "C4:C5" };
  assert.throws(() => validateSortRange(array, selected), /배열/);
  const emptyKey = sheet();
  for (let r = 2; r <= 4; r++) delete emptyKey.cellData![r][2];
  assert.throws(
    () =>
      createSortConfiguration(emptyKey, selected, true, [
        { column: 1, ascending: true },
      ]),
    /기준 열/,
  );
});
test("header labels are optional and bounded without changing the stored title", () => {
  const data = sheet();
  data.cellData![1][1] = { v: "한".repeat(100) };
  const labels = sortColumns(data, selected, true);
  assert.equal(labels[0].label.length, 66);
  assert.equal(data.cellData![1][1].v, "한".repeat(100));
  assert.deepEqual(
    sortColumns(data, selected, false).map((c) => c.label),
    ["B열", "C열", "D열"],
  );
});
test("stale content, styles or sheet structure invalidate a captured sort; outside cells do not", () => {
  const data = sheet(),
    fingerprint = sortRangeFingerprint(data, selected);
  assertSortUnchanged(data, selected, fingerprint);
  data.cellData![8] = { 8: { v: "범위 밖" } };
  assertSortUnchanged(data, selected, fingerprint);
  data.cellData![2][1].s = { bl: 1 };
  assert.throws(
    () => assertSortUnchanged(data, selected, fingerprint),
    /바뀌었습니다/,
  );
  const changed = sheet();
  changed.rowCount = 201;
  assert.throws(() => assertSortUnchanged(changed, selected, fingerprint));
  const formula = sheet();
  formula.cellData![3][2] = { f: "=1+2", v: 3 };
  assert.throws(() => assertSortUnchanged(formula, selected, fingerprint));
});
