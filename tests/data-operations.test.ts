import test from "node:test";
import assert from "node:assert/strict";
import type { ICellData } from "@univerjs/core";
import {
  cellAddress,
  cellFingerprint,
  createSparseTextPatch,
  findCells,
  planReplacement,
  planTrim,
} from "../src/data-operations";

const example = () =>
  ({
    0: {
      0: { v: "Apple apple", t: 1 },
      1: { v: "apple", t: 1 },
      2: { v: 123, t: 2 },
      3: { v: "apple", f: "=A1", t: 1 },
    },
    1: {
      0: { v: "  한글  문장  ", t: 1, s: { bl: 1 } },
      1: { v: "00123", t: 4 },
      2: { v: "apple", si: "shared-1" },
      3: {
        v: " apple ",
        p: {
          id: "rich-text-fixture",
          documentStyle: {},
          body: { dataStream: " apple \r\n" },
        },
      },
    },
    5: { 4: { v: "APPLE", t: 1 }, 5: { v: false, t: 3 } },
  }) as Record<number, Record<number, ICellData>>;

test("find is literal, supports case/whole-cell options, sorted addresses and frozen selection", () => {
  const matrix = example();
  assert.deepEqual(
    findCells(matrix, "apple").map((c) => c.address),
    ["A1", "B1", "D1", "C2", "D2", "E6"],
  );
  assert.deepEqual(
    findCells(matrix, "apple", { caseSensitive: true, wholeCell: true }).map(
      (c) => c.address,
    ),
    ["B1", "D1", "C2"],
  );
  assert.deepEqual(
    findCells(matrix, "apple", {
      bounds: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    }).map((c) => c.address),
    ["A1", "B1"],
  );
  matrix[2] = { 0: { v: "a.b $1 [a] .*" } };
  assert.equal(findCells(matrix, ".*").length, 1);
  assert.equal(findCells(matrix, "a.b").length, 1);
  assert.equal(findCells(matrix, "[a]").length, 1);
  matrix[3] = { 0: { v: "apple\n" } };
  assert.ok(
    !findCells(matrix, "apple", { wholeCell: true }).some(
      (cell) => cell.address === "A4",
    ),
  );
  assert.equal(cellAddress(9999, 199), "GR10000");
  assert.throws(() => cellAddress(0, 200));
});
test("replace never touches formula/shared-formula/rich-text/numeric/boolean cells", () => {
  const matrix = example(),
    before = structuredClone(matrix);
  const found = findCells(matrix, "apple"),
    plan = planReplacement(found, "apple", "수정");
  assert.equal(plan.protectedCount, 3);
  assert.deepEqual(
    plan.changes.map((c) => c.address),
    ["A1", "B1", "E6"],
  );
  assert.equal(plan.changes[0].after, "수정 수정");
  assert.equal(
    planReplacement(findCells(matrix, "123"), "123", "456").protectedCount,
    1,
  );
  assert.equal(
    planReplacement(findCells(matrix, "false"), "false", "true").changes.length,
    0,
  );
  assert.deepEqual(matrix, before);
});
test("sparse patch stores literal text, preserves unrelated cells/style and uses absolute coordinates", () => {
  const matrix = example();
  const found = findCells(matrix, "APPLE", {
    wholeCell: true,
    caseSensitive: true,
  });
  const plan = planReplacement(found, "APPLE", "=SUM(A1:A2)", {
    wholeCell: true,
    caseSensitive: true,
  });
  const patch = createSparseTextPatch(matrix, plan.changes);
  assert.deepEqual(Object.keys(patch.values), ["5"]);
  assert.deepEqual(Object.keys(patch.values[5]), ["4"]);
  assert.equal(patch.values[5][4].v, "=SUM(A1:A2)");
  assert.equal(patch.values[5][4].t, 1);
  assert.equal(patch.values[5][4].f, null);
  assert.equal(patch.values[5][4].s, undefined);
  assert.deepEqual(patch.bounds, {
    startRow: 5,
    endRow: 5,
    startColumn: 4,
    endColumn: 4,
  });
});
test("replacement strings have no regex expansion and cannot grow a cell without bound", () => {
  const matrix = { 0: { 0: { v: "a a a" } } };
  assert.equal(
    planReplacement(findCells(matrix, "a"), "a", "$& $1").changes[0].after,
    "$& $1 $& $1 $& $1",
  );
  const huge = { 0: { 0: { v: "a".repeat(32767) } } };
  assert.throws(
    () => planReplacement(findCells(huge, "a"), "a", "bb"),
    /32,767/,
  );
  const many: Record<number, Record<number, ICellData>> = {};
  for (let row = 0; row < 6000; row++) many[row] = { 0: { v: "a" } };
  assert.throws(
    () => planReplacement(findCells(many, "a"), "a", "b".repeat(2000)),
    /10MB/,
  );
  assert.throws(() => findCells(matrix, ""));
  assert.throws(() => findCells(matrix, "a".repeat(501)));
});
test("trim preview strips edge whitespace only and keeps numeric-looking strings as strings", () => {
  const matrix = example();
  matrix[7] = {
    0: { v: "\t 00123\u00a0" },
    1: { v: "  =SUM(A1:A2)  " },
    2: { v: "   " },
  };
  const plan = planTrim(matrix);
  assert.deepEqual(
    plan.changes.map((c) => c.address),
    ["A2", "A8", "B8", "C8"],
  );
  assert.equal(plan.changes[0].after, "한글  문장");
  assert.equal(plan.changes[1].after, "00123");
  assert.equal(plan.changes[2].after, "=SUM(A1:A2)");
  assert.equal(plan.changes[3].after, "");
  const patch = createSparseTextPatch(matrix, plan.changes);
  assert.equal(patch.values[7][0].t, 1);
  assert.equal(patch.values[7][1].f, null);
  assert.equal(
    planTrim(matrix, { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 })
      .changes.length,
    1,
  );
});
test("stale plans fail as a whole before generating a patch; formatting-only changes stay safe", () => {
  const matrix = example();
  const plan = planReplacement(findCells(matrix, "apple"), "apple", "orange");
  matrix[0][0].s = { bl: 1 };
  assert.ok(createSparseTextPatch(matrix, plan.changes));
  matrix[5][4].v = "someone else edited";
  assert.throws(() => createSparseTextPatch(matrix, plan.changes), /E6/);
  assert.equal(matrix[0][0].v, "Apple apple");
  const original = example(),
    snapshot = planReplacement(
      findCells(original, "Apple"),
      "Apple",
      "changed",
    );
  original[0][0].f = "=A9";
  assert.throws(() => createSparseTextPatch(original, snapshot.changes));
  assert.notEqual(
    cellFingerprint({ v: "123", t: 1 }),
    cellFingerprint({ v: "123", t: 2 }),
  );
});
test("invalid or duplicate coordinates and forged protected-cell changes are rejected", () => {
  const matrix = example(),
    plan = planReplacement(findCells(matrix, "Apple"), "Apple", "changed");
  assert.throws(
    () => createSparseTextPatch(matrix, [plan.changes[0], plan.changes[0]]),
    /중복/,
  );
  assert.throws(() =>
    createSparseTextPatch(matrix, [{ ...plan.changes[0], row: -1 }]),
  );
  assert.throws(() => findCells({ "-1": { 0: { v: "a" } } }, "a"));
  assert.throws(() =>
    findCells(matrix, "a", {
      bounds: { startRow: 5, endRow: 0, startColumn: 0, endColumn: 1 },
    }),
  );
  const protectedCell = matrix[0][3];
  assert.throws(() =>
    createSparseTextPatch(matrix, [
      {
        row: 0,
        col: 3,
        address: "D1",
        fingerprint: cellFingerprint(protectedCell),
        before: "apple",
        after: "changed",
      },
    ]),
  );
});
