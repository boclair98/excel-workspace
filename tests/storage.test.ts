import 'fake-indexeddb/auto';
import test from 'node:test';
import assert from 'node:assert/strict';
import { newBook } from '../src/model';
import { saveBook, listBooks, deleteBook } from '../src/storage';
test('local library saves, updates without duplicates, restores and deletes', async () => {
  const book = newBook(true); await saveBook(book); book.name = '수정한 이름'; await saveBook(book);
  const rows = await listBooks(); assert.equal(rows.length, 1); assert.equal(rows[0].name, book.name); assert.equal(rows[0].book.sheets[book.sheetOrder[0]].cellData![1][3].v, 3000000);
  await deleteBook(book.id); assert.equal((await listBooks()).length, 0);
});
test('invalid documents cannot overwrite a valid saved workbook', async () => {
  const book = newBook(); await saveBook(book); book.sheetOrder = [];
  await assert.rejects(saveBook(book)); assert.equal((await listBooks()).find(b => b.id === book.id)?.book.sheetOrder.length, 1); await deleteBook(book.id);
});

test('50-document limit rejects new entries atomically but permits updating an existing document', async () => {
  const books = Array.from({ length: 50 }, () => newBook());
  for (const book of books) await saveBook(book);
  const overflow = newBook(); await assert.rejects(saveBook(overflow), /50개/);
  assert.equal((await listBooks()).length, 50);
  books[0].name = '한도에서도 수정 가능'; await saveBook(books[0]);
  assert.equal((await listBooks()).find(b => b.id === books[0].id)?.name, '한도에서도 수정 가능');
  for (const book of books) await deleteBook(book.id);
  assert.equal((await listBooks()).length, 0);
});
