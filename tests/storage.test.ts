import "fake-indexeddb/auto";
import test from "node:test";
import assert from "node:assert/strict";
import { newBook, LIMITS } from "../src/model";
import {
  saveBook,
  listBooks,
  deleteBook,
  createCheckpoint,
  listCheckpoints,
  deleteCheckpoint,
  StorageConflictError,
  type SavedBook,
} from "../src/storage";

test("v1 books migrate to v2 unchanged and gain a separate checkpoints store", async () => {
  const book = newBook(true);
  const original: SavedBook = {
    id: book.id,
    name: book.name,
    updatedAt: 1700000000000,
    book,
  };
  let legacy: IDBDatabase | undefined;
  await new Promise<void>((resolve, reject) => {
    const request = indexedDB.open("excel-workspace-v1", 1);
    request.onupgradeneeded = () =>
      request.result.createObjectStore("books", { keyPath: "id" });
    request.onerror = () => reject(request.error);
    request.onsuccess = () => {
      const database = request.result;
      legacy = database;
      const tx = database.transaction("books", "readwrite");
      tx.objectStore("books").put(original);
      tx.oncomplete = () => resolve();
      tx.onerror = () => {
        database.close();
        reject(tx.error);
      };
    };
  });
  await assert.rejects(listBooks(), /이전 버전/);
  legacy!.close();
  assert.deepEqual(
    (await listBooks()).find((row) => row.id === book.id),
    original,
  );
  const checkpoint = await createCheckpoint(book, "v1 이전본");
  assert.equal(checkpoint.bookId, book.id);
  assert.deepEqual(checkpoint.book, book);
  await new Promise<void>((resolve, reject) => {
    const request = indexedDB.open("excel-workspace-v1");
    request.onsuccess = () => {
      assert.equal(request.result.version, 2);
      assert.ok(request.result.objectStoreNames.contains("books"));
      assert.ok(request.result.objectStoreNames.contains("checkpoints"));
      request.result.close();
      resolve();
    };
    request.onerror = () => reject(request.error);
  });
  await deleteBook(book.id);
  assert.deepEqual(await listCheckpoints(book.id), []);
});
test("local library saves, updates without duplicates, restores and deletes", async () => {
  const book = newBook(true);
  await saveBook(book);
  book.name = "수정한 이름";
  await saveBook(book);
  const rows = await listBooks();
  assert.equal(rows.length, 1);
  assert.equal(rows[0].name, book.name);
  assert.equal(
    rows[0].book.sheets[book.sheetOrder[0]].cellData![1][3].v,
    3000000,
  );
  await deleteBook(book.id);
  assert.equal((await listBooks()).length, 0);
});
test("invalid documents cannot overwrite a valid saved workbook", async () => {
  const book = newBook();
  await saveBook(book);
  book.sheetOrder = [];
  await assert.rejects(saveBook(book));
  assert.equal(
    (await listBooks()).find((b) => b.id === book.id)?.book.sheetOrder.length,
    1,
  );
  await deleteBook(book.id);
});

test("50-document limit rejects new entries atomically but permits updating an existing document", async () => {
  const books = Array.from({ length: 50 }, () => newBook());
  for (const book of books) await saveBook(book);
  const overflow = newBook();
  await assert.rejects(saveBook(overflow), /50개/);
  assert.equal((await listBooks()).length, 50);
  books[0].name = "한도에서도 수정 가능";
  await saveBook(books[0]);
  assert.equal(
    (await listBooks()).find((b) => b.id === books[0].id)?.name,
    "한도에서도 수정 가능",
  );
  for (const book of books) await deleteBook(book.id);
  assert.equal((await listBooks()).length, 0);
});

test("conditional concurrent saves reject the stale writer without overwriting the winner", async () => {
  const book = newBook();
  const original = await saveBook(book, { expectedUpdatedAt: null });
  const first = structuredClone(book),
    second = structuredClone(book);
  first.name = "첫 번째 탭";
  second.name = "두 번째 탭";
  const results = await Promise.allSettled([
    saveBook(first, { expectedUpdatedAt: original.updatedAt }),
    saveBook(second, { expectedUpdatedAt: original.updatedAt }),
  ]);
  const successes = results.filter((result) => result.status === "fulfilled");
  const failures = results.filter((result) => result.status === "rejected");
  assert.equal(successes.length, 1);
  assert.equal(failures.length, 1);
  assert.ok(failures[0].reason instanceof StorageConflictError);
  assert.equal(failures[0].reason.bookId, book.id);
  assert.equal(failures[0].reason.expectedUpdatedAt, original.updatedAt);
  assert.equal(
    failures[0].reason.actualUpdatedAt,
    successes[0].value.updatedAt,
  );
  assert.deepEqual(
    (await listBooks()).find((row) => row.id === book.id),
    successes[0].value,
  );
  await assert.rejects(
    saveBook(book, { expectedUpdatedAt: null }),
    StorageConflictError,
  );
  await deleteBook(book.id);
  await assert.rejects(
    saveBook(book, { expectedUpdatedAt: original.updatedAt }),
    (error: unknown) =>
      error instanceof StorageConflictError && error.actualUpdatedAt === null,
  );
});

test("updatedAt is monotonic and legacy unconditional saves return an isolated SavedBook", async (t) => {
  t.mock.method(Date, "now", () => 1000);
  const book = newBook();
  const pending = saveBook(book);
  book.name = "호출 뒤 바뀐 이름";
  const first = await pending;
  assert.notEqual(first.name, book.name);
  assert.equal(first.updatedAt, 1000);
  const second = await saveBook(book, { expectedUpdatedAt: first.updatedAt });
  assert.equal(second.updatedAt, 1001);
  const third = await saveBook(book);
  assert.equal(third.updatedAt, 1002);
  third.book.name = "반환 객체 수정";
  assert.equal(
    (await listBooks()).find((row) => row.id === book.id)?.book.name,
    book.name,
  );
  await assert.rejects(saveBook(book, { expectedUpdatedAt: NaN }), /시각/);
  await assert.rejects(saveBook(book, { expectedUpdatedAt: -1 }), /시각/);
  await deleteBook(book.id);
});

test("checkpoints keep the newest ten per document, preserve their snapshots and delete independently", async () => {
  const book = newBook(true),
    other = newBook();
  await saveBook(book);
  await saveBook(other);
  const otherPoint = await createCheckpoint(other, "다른 문서");
  const pointIds: string[] = [];
  for (let i = 0; i < 12; i++) {
    book.name = `문서 ${i}`;
    const point = await createCheckpoint(book, `  지점 ${i}  `);
    pointIds.push(point.id);
    assert.equal(point.name, `지점 ${i}`);
  }
  const points = await listCheckpoints(book.id);
  assert.equal(points.length, 10);
  assert.deepEqual(
    points.map((point) => point.name),
    Array.from({ length: 10 }, (_, i) => `지점 ${11 - i}`),
  );
  assert.equal(points[0].book.name, "문서 11");
  assert.equal(points[9].book.name, "문서 2");
  assert.ok(
    points.every((point, i) => !i || points[i - 1].createdAt > point.createdAt),
  );
  assert.ok(
    !points.some(
      (point) => point.id === pointIds[0] || point.id === pointIds[1],
    ),
  );
  assert.equal((await listCheckpoints(other.id))[0].id, otherPoint.id);
  // A checkpoint does not implicitly replace the saved document's revision.
  assert.equal(
    (await listBooks()).find((row) => row.id === book.id)?.name,
    "예시 가계부",
  );
  await deleteCheckpoint(points[0].id);
  assert.equal((await listCheckpoints(book.id)).length, 9);
  await deleteCheckpoint(points[0].id);
  assert.equal((await listCheckpoints(book.id)).length, 9);
  await deleteBook(book.id);
  assert.equal((await listCheckpoints(book.id)).length, 0);
  assert.equal((await listCheckpoints(other.id)).length, 1);
  await deleteBook(other.id);
});

test("checkpoint validation rejects unowned, invalid, unnamed and oversized snapshots without changing history", async () => {
  const book = newBook();
  await assert.rejects(createCheckpoint(book, "저장 전"), /먼저 보관/);
  await saveBook(book);
  const first = await createCheckpoint(book, "정상");
  await assert.rejects(createCheckpoint(book, "  "), /이름/);
  await assert.rejects(createCheckpoint(book, "가".repeat(81)), /이름/);
  const invalid = structuredClone(book);
  invalid.sheetOrder = [];
  await assert.rejects(createCheckpoint(invalid, "손상"));
  const oversized = structuredClone(book) as typeof book & {
    oversized: string;
  };
  oversized.oversized = "x".repeat(LIMITS.bytes + 1);
  await assert.rejects(createCheckpoint(oversized, "과대"), /10MB/);
  const nearLimit = structuredClone(book) as typeof book & { padding: string };
  nearLimit.padding = "";
  nearLimit.padding = "x".repeat(
    LIMITS.bytes -
      new TextEncoder().encode(JSON.stringify(nearLimit)).length -
      64,
  );
  await assert.rejects(
    createCheckpoint(nearLimit, "메타데이터 포함 한도"),
    /복구 지점 하나/,
  );
  assert.deepEqual(
    (await listCheckpoints(book.id)).map((point) => point.id),
    [first.id],
  );
  await deleteBook(book.id);
});

test("failed checkpoint pruning and cascade deletion roll back every store mutation", async (t) => {
  const book = newBook();
  await saveBook(book);
  for (let i = 0; i < 10; i++) await createCheckpoint(book, `복구 ${i}`);
  const before = await listCheckpoints(book.id),
    oldest = before.at(-1)!;
  const nativeDelete = IDBObjectStore.prototype.delete;
  t.mock.method(
    IDBObjectStore.prototype,
    "delete",
    function (this: IDBObjectStore, key: IDBValidKey | IDBKeyRange) {
      if (this.name === "checkpoints" && key === oldest.id)
        throw new DOMException(
          "simulated storage failure",
          "QuotaExceededError",
        );
      return nativeDelete.call(this, key);
    },
  );
  await assert.rejects(
    createCheckpoint(book, "실패한 새 지점"),
    /simulated storage failure/,
  );
  assert.deepEqual(await listCheckpoints(book.id), before);
  await assert.rejects(deleteBook(book.id), /simulated storage failure/);
  assert.ok((await listBooks()).some((row) => row.id === book.id));
  assert.deepEqual(await listCheckpoints(book.id), before);
  t.mock.restoreAll();
  await deleteBook(book.id);
});
