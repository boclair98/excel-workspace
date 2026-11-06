import { openDB } from "idb";
import type { IWorkbookData } from "@univerjs/core";
import { validateBook } from "./model";
export interface SavedBook {
  id: string;
  name: string;
  updatedAt: number;
  book: IWorkbookData;
}
let databasePromise: ReturnType<typeof openDB> | undefined;
const db = () =>
  (databasePromise ??= openDB("excel-workspace-v1", 1, {
    upgrade(database) {
      database.createObjectStore("books", { keyPath: "id" });
    },
    terminated() {
      databasePromise = undefined;
    },
  }).catch((error) => {
    databasePromise = undefined;
    throw error;
  }));
export async function saveBook(book: IWorkbookData) {
  validateBook(book);
  const database = await db();
  const tx = database.transaction("books", "readwrite");
  try {
    const count = await tx.store.count();
    if (count >= 50 && !(await tx.store.get(book.id)))
      throw new Error("보관함이 가득 찼습니다. 50개까지 보관할 수 있습니다.");
    await tx.store.put({
      id: book.id,
      name: book.name,
      updatedAt: Date.now(),
      book,
    } satisfies SavedBook);
    await tx.done;
  } catch (error) {
    try {
      tx.abort();
    } catch {
      /* Transaction may already have aborted. */
    }
    await tx.done.catch(() => {});
    throw error;
  }
}
export async function listBooks(): Promise<SavedBook[]> {
  return (await (await db()).getAll("books")).sort(
    (a, b) => b.updatedAt - a.updatedAt,
  );
}
export async function deleteBook(id: string) {
  await (await db()).delete("books", id);
}
