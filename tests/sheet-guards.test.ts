import test from "node:test";
import assert from "node:assert/strict";
import { newBook, LIMITS } from "../src/model";
import {
  prepareSheetDuplicate,
  requireRemovableSheet,
  validateSheetName,
} from "../src/sheet-guards";

test("duplicate preflight preserves source, formulas, style and insertion order with unique engine name", () => {
  const book = newBook(true),
    id = book.sheetOrder[0],
    original = structuredClone(book);
  const first = prepareSheetDuplicate(book, id);
  assert.deepEqual(book, original);
  assert.equal(first.name, "가계부 (복사본)");
  assert.equal(first.candidate.sheetOrder[0], id);
  const copiedId = first.candidate.sheetOrder[1];
  assert.notEqual(copiedId, id);
  assert.deepEqual(
    first.candidate.sheets[copiedId].cellData,
    book.sheets[id].cellData,
  );
  const second = prepareSheetDuplicate(first.candidate, id);
  assert.equal(second.name, "가계부 (복사본2)");
  assert.equal(second.candidate.sheetOrder[2], copiedId);
});

test("duplicate preflight rejects sheet, populated-cell and byte limits before any mutation", () => {
  const book = newBook(),
    id = book.sheetOrder[0];
  for (let index = 1; index < LIMITS.sheets; index++) {
    const key = `sheet-${index}`;
    book.sheetOrder.push(key);
    book.sheets[key] = {
      ...structuredClone(book.sheets[id]),
      id: key,
      name: `시트 ${index}`,
    };
  }
  assert.throws(() => prepareSheetDuplicate(book, id), /30개/);
  const dense = newBook(),
    denseId = dense.sheetOrder[0],
    sheet = dense.sheets[denseId];
  sheet.rowCount = 300;
  sheet.columnCount = 200;
  sheet.cellData = {};
  for (let row = 0; row < 300; row++) {
    sheet.cellData[row] = {};
    for (let col = 0; col < 200; col++) sheet.cellData[row][col] = { v: 1 };
  }
  assert.throws(() => prepareSheetDuplicate(dense, denseId), /100,000/);
  const large = newBook(),
    largeId = large.sheetOrder[0];
  large.sheets[largeId].cellData = {};
  for (let row = 0; row < 190; row++)
    large.sheets[largeId].cellData![row] = { 0: { v: "x".repeat(30000) } };
  assert.throws(() => prepareSheetDuplicate(large, largeId), /10MB/);
  assert.equal(large.sheetOrder.length, 1);
});

test("sheet names reject ambiguous duplicates, forbidden syntax and overlong copy names", () => {
  const book = newBook(),
    id = book.sheetOrder[0];
  assert.equal(validateSheetName(book, id, "  원본  "), "원본");
  for (const invalid of [
    "",
    " ",
    "'인용",
    "인용'",
    "a/b",
    "a:b",
    "a\\b",
    "[시트]",
    "a*",
    "a?",
    "가".repeat(32),
  ])
    assert.throws(() => validateSheetName(book, id, invalid));
  book.sheets[id].name = "Sales";
  const next = prepareSheetDuplicate(book, id).candidate;
  assert.throws(
    () => validateSheetName(next, id, "sales (복사본)"),
    /같은 이름/,
  );
  book.sheets[id].name = "가".repeat(31);
  assert.throws(() => prepareSheetDuplicate(book, id), /이름이 길어요/);
  assert.throws(() => prepareSheetDuplicate(book, "missing"), /없어졌어요/);
});

test("last visible sheet stays protected while hidden sheets remain removable", () => {
  const book = newBook(),
    id = book.sheetOrder[0];
  assert.throws(() => requireRemovableSheet(book, id), /마지막/);
  const candidate = prepareSheetDuplicate(book, id).candidate,
    copyId = candidate.sheetOrder[1];
  assert.doesNotThrow(() => requireRemovableSheet(candidate, id));
  candidate.sheets[copyId].hidden = 1;
  assert.throws(() => requireRemovableSheet(candidate, id), /마지막/);
  assert.doesNotThrow(() => requireRemovableSheet(candidate, copyId));
  candidate.sheets[copyId].hidden = 2;
  assert.throws(() => requireRemovableSheet(candidate, id), /마지막/);
  assert.doesNotThrow(() => requireRemovableSheet(candidate, copyId));
  assert.throws(
    () => requireRemovableSheet(candidate, "missing"),
    /변경되었어요/,
  );
});
