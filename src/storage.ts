import { openDB, type DBSchema, type IDBPDatabase } from "idb";
import type { IWorkbookData } from "@univerjs/core";
import { LIMITS, validateBook } from "./model";
export interface SavedBook {
  id: string;
  name: string;
  updatedAt: number;
  book: IWorkbookData;
}
export interface Checkpoint {
  id: string;
  bookId: string;
  name: string;
  createdAt: number;
  book: IWorkbookData;
}
export interface SaveBookOptions {
  /** Omit for a legacy unconditional save. null expects no existing document. */
  expectedUpdatedAt?: number | null;
}
export class StorageConflictError extends Error {
  readonly name = "StorageConflictError";
  constructor(
    readonly bookId: string,
    readonly expectedUpdatedAt: number | null,
    readonly actualUpdatedAt: number | null,
  ) {
    super(
      "다른 탭에서 이 문서가 변경되었습니다. 최신 보관본을 확인하거나 현재 내용을 사본으로 저장해 주세요.",
    );
  }
}
interface WorkspaceDB extends DBSchema {
  books: { key: string; value: SavedBook };
  checkpoints: {
    key: string;
    value: Checkpoint;
    indexes: { byBookId: string };
  };
}
const DATABASE_NAME = "excel-workspace-v1";
const CHECKPOINT_LIMIT = 10;
let databasePromise: Promise<IDBPDatabase<WorkspaceDB>> | undefined;
function openWorkspace() {
  return new Promise<IDBPDatabase<WorkspaceDB>>((resolve, reject) => {
    let rejected = false;
    const opening = openDB<WorkspaceDB>(DATABASE_NAME, 2, {
      upgrade(database, oldVersion) {
        if (oldVersion < 1)
          database.createObjectStore("books", { keyPath: "id" });
        if (oldVersion < 2) {
          const checkpoints = database.createObjectStore("checkpoints", {
            keyPath: "id",
          });
          checkpoints.createIndex("byBookId", "bookId");
        }
      },
      blocked() {
        rejected = true;
        reject(
          new Error(
            "이전 버전의 엑셀 탭을 모두 닫은 뒤 다시 시도해 주세요. 기존 문서는 그대로 보관됩니다.",
          ),
        );
      },
      blocking() {
        const connection = databasePromise;
        databasePromise = undefined;
        void connection?.then((database) => database.close());
      },
      terminated() {
        databasePromise = undefined;
      },
    });
    void opening.then((database) => {
      if (rejected) database.close();
      else resolve(database);
    }, reject);
  });
}
const db = () =>
  (databasePromise ??= openWorkspace().catch((error) => {
    databasePromise = undefined;
    throw error;
  }));
function validateId(id: string) {
  if (typeof id !== "string" || !id || id.length > 200)
    throw new Error("문서 또는 복구 지점 식별자가 올바르지 않습니다.");
}
async function abortTransaction(tx: {
  abort: () => void;
  done: Promise<unknown>;
}) {
  try {
    tx.abort();
  } catch {
    /* Transaction may already have aborted. */
  }
  await tx.done.catch(() => {});
}
export async function saveBook(
  book: IWorkbookData,
  options: SaveBookOptions = {},
): Promise<SavedBook> {
  validateBook(book);
  const expected = options.expectedUpdatedAt;
  if (
    expected !== undefined &&
    expected !== null &&
    (!Number.isSafeInteger(expected) || expected < 0)
  )
    throw new Error("예상 저장 시각이 올바르지 않습니다.");
  const snapshot = structuredClone(book);
  const database = await db();
  const tx = database.transaction("books", "readwrite");
  try {
    const previous = await tx.store.get(snapshot.id);
    const actual = previous?.updatedAt ?? null;
    if (expected !== undefined && expected !== actual)
      throw new StorageConflictError(snapshot.id, expected, actual);
    if (!previous && (await tx.store.count()) >= 50)
      throw new Error("보관함이 가득 찼습니다. 50개까지 보관할 수 있습니다.");
    const saved: SavedBook = {
      id: snapshot.id,
      name: snapshot.name,
      updatedAt: Math.max(Date.now(), (previous?.updatedAt ?? 0) + 1),
      book: snapshot,
    };
    await tx.store.put(saved);
    await tx.done;
    return saved;
  } catch (error) {
    await abortTransaction(tx);
    throw error;
  }
}
export async function listBooks(): Promise<SavedBook[]> {
  return (await (await db()).getAll("books")).sort(
    (a, b) => b.updatedAt - a.updatedAt,
  );
}
export async function deleteBook(id: string): Promise<void> {
  validateId(id);
  const database = await db();
  const tx = database.transaction(["books", "checkpoints"], "readwrite");
  try {
    await tx.objectStore("books").delete(id);
    const checkpoints = tx.objectStore("checkpoints");
    const keys = await checkpoints.index("byBookId").getAllKeys(id);
    for (const key of keys) await checkpoints.delete(key);
    await tx.done;
  } catch (error) {
    await abortTransaction(tx);
    throw error;
  }
}
export async function createCheckpoint(
  book: IWorkbookData,
  name: string,
): Promise<Checkpoint> {
  validateBook(book);
  if (typeof name !== "string" || !name.trim() || name.trim().length > 80)
    throw new Error("복구 지점 이름은 1~80자로 입력해 주세요.");
  const snapshot = structuredClone(book);
  const database = await db();
  const tx = database.transaction(["books", "checkpoints"], "readwrite");
  try {
    if (!(await tx.objectStore("books").getKey(snapshot.id)))
      throw new Error("문서를 먼저 보관한 뒤 복구 지점을 만들어 주세요.");
    const store = tx.objectStore("checkpoints");
    const previous = await store.index("byBookId").getAll(snapshot.id);
    previous.sort(
      (a, b) => b.createdAt - a.createdAt || a.id.localeCompare(b.id),
    );
    const checkpoint: Checkpoint = {
      id: crypto.randomUUID(),
      bookId: snapshot.id,
      name: name.trim(),
      createdAt: Math.max(Date.now(), (previous[0]?.createdAt ?? 0) + 1),
      book: snapshot,
    };
    if (
      new TextEncoder().encode(JSON.stringify(checkpoint)).length > LIMITS.bytes
    )
      throw new Error(
        "복구 지점 하나는 10MB 이내여야 합니다. 문서 내용을 줄여 주세요.",
      );
    await store.add(checkpoint);
    for (const expired of previous.slice(CHECKPOINT_LIMIT - 1))
      await store.delete(expired.id);
    await tx.done;
    return checkpoint;
  } catch (error) {
    await abortTransaction(tx);
    throw error;
  }
}
export async function listCheckpoints(bookId: string): Promise<Checkpoint[]> {
  validateId(bookId);
  return (
    await (await db()).getAllFromIndex("checkpoints", "byBookId", bookId)
  ).sort((a, b) => b.createdAt - a.createdAt || a.id.localeCompare(b.id));
}
export async function deleteCheckpoint(id: string): Promise<void> {
  validateId(id);
  await (await db()).delete("checkpoints", id);
}
