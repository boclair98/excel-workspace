import type { IWorkbookData } from "@univerjs/core";
import { LIMITS, validateBook } from "./model";

export function validateSheetName(
  book: IWorkbookData,
  sheetId: string,
  value: string,
): string {
  if (!book.sheetOrder.includes(sheetId) || !book.sheets[sheetId])
    throw new Error("시트가 변경되었어요. 시트 관리 창을 다시 열어 주세요.");
  const name = value.trim();
  if (
    !name ||
    name.length > 31 ||
    /[\[\]:*?\/\\]/.test(name) ||
    /^'|'$/.test(name)
  )
    throw new Error(
      "시트 이름은 1~31자로 입력하고 : \\ / ? * [ ] 및 양끝의 작은따옴표는 제외해 주세요.",
    );
  if (
    book.sheetOrder.some(
      (id) =>
        id !== sheetId &&
        book.sheets[id].name?.toLowerCase() === name.toLowerCase(),
    )
  )
    throw new Error("같은 이름의 시트가 있어요. 다른 이름을 입력해 주세요.");
  return name;
}

/** Preflight mirrors Univer's pinned Korean copy naming; the actual copy uses its undoable command. */
export function prepareSheetDuplicate(
  book: IWorkbookData,
  sheetId: string,
): { candidate: IWorkbookData; name: string } {
  validateBook(book);
  const source = book.sheets[sheetId];
  if (!book.sheetOrder.includes(sheetId) || !source)
    throw new Error(
      "복제할 시트가 없어졌어요. 시트 관리 창을 다시 열어 주세요.",
    );
  if (book.sheetOrder.length >= LIMITS.sheets)
    throw new Error(
      "시트는 30개까지 사용할 수 있어요. 불필요한 시트를 정리한 뒤 복제해 주세요.",
    );
  const names = new Set(
    book.sheetOrder.map((id) => book.sheets[id].name?.toLowerCase()),
  );
  let name = `${source.name} (복사본)`,
    count = 2;
  while (names.has(name.toLowerCase()))
    name = `${source.name} (복사본${count++})`;
  if (name.length > 31)
    throw new Error(
      "복제할 시트의 이름이 길어요. 원본 이름을 짧게 바꾼 뒤 다시 복제해 주세요. 복제본 이름도 31자 이내여야 합니다.",
    );
  const candidate = structuredClone(book),
    id = crypto.randomUUID();
  candidate.sheets[id] = { ...structuredClone(source), id, name };
  candidate.sheetOrder.splice(candidate.sheetOrder.indexOf(sheetId) + 1, 0, id);
  validateBook(candidate);
  return { candidate, name };
}

export function requireRemovableSheet(
  book: IWorkbookData,
  sheetId: string,
): void {
  if (!book.sheetOrder.includes(sheetId) || !book.sheets[sheetId])
    throw new Error("시트가 변경되었어요. 시트 관리 창을 다시 열어 주세요.");
  if (
    book.sheetOrder.length <= 1 ||
    (!book.sheets[sheetId].hidden &&
      book.sheetOrder.filter((id) => !book.sheets[id].hidden).length <= 1)
  )
    throw new Error(
      "마지막으로 표시된 시트는 숨기거나 삭제할 수 없어요. 다른 시트를 먼저 표시해 주세요.",
    );
}
