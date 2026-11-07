import { test, expect, type Page, type TestInfo } from "@playwright/test";
import { readFile } from "node:fs/promises";
import type { IWorkbookData } from "@univerjs/core";

async function ready(page: Page) {
  await expect(page.locator("#loading")).toBeHidden();
  await expect(page.locator("#sheet-tabs button").first()).toBeVisible();
}
async function cell(page: Page, address: string, value?: string) {
  await page.locator("#address").fill(address);
  await page.locator("#address").press("Enter");
  if (value !== undefined) {
    await page.locator("#cell-input").fill(value);
    await page.locator('#cell-form button[type="submit"]').click();
    await expect(page.locator("#save-state")).toHaveText(
      "이 브라우저에 보관됨",
    );
  }
}
async function tab(page: Page, name: string) {
  await page
    .locator("#sheet-tabs button")
    .filter({
      hasText: new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`),
    })
    .click();
}
async function manage(page: Page, name?: string) {
  await page.locator("#sheet-manage").click();
  await expect(page.locator("#sheet-manager-dialog")).toBeVisible();
  if (name)
    await page
      .locator("#sheet-manager-list .sheet-manager-item")
      .filter({ has: page.getByText(name, { exact: true }) })
      .click();
}
async function action(page: Page, id: string) {
  await page.locator(`#sheet-manager-${id}`).click();
  await expect(page.locator("#sheet-manager-dialog")).toHaveAttribute(
    "aria-busy",
    "false",
  );
  await expect(page.locator("#sheet-manager-error")).toBeHidden();
}
async function close(page: Page) {
  await page.locator("#sheet-manager-close").click();
  await expect(page.locator("#sheet-manager-dialog")).toHaveCount(0);
}
async function rename(page: Page, name: string) {
  await page.locator("#sheet-manager-name").fill(name);
  await action(page, "rename");
  await expect(page.locator("#sheet-manager-selected")).toHaveText(name);
}
async function backup(
  page: Page,
  info: TestInfo,
  name: string,
): Promise<IWorkbookData> {
  if (!(await page.locator("#sidebar").isVisible())) {
    await page.locator("#library-toggle").click();
  }
  const event = page.waitForEvent("download");
  await page.locator("#backup").click();
  const path = info.outputPath(name);
  await (await event).saveAs(path);
  return JSON.parse(await readFile(path, "utf8")).book;
}
function sheet(book: IWorkbookData, name: string) {
  const match = Object.values(book.sheets).find((value) => value.name === name);
  expect(match, `snapshot includes ${name}`).toBeTruthy();
  return match!;
}

test("engine sheet commands preserve formulas through rename, copy, order, deletion and undo", async ({
  page,
}, info) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/");
  await ready(page);
  await manage(page);
  await rename(page, "원본");
  await close(page);
  await cell(page, "A1", "21");
  await cell(page, "B1", "=A1*2");
  await page.locator("#sheet-add").click();
  await manage(page);
  await rename(page, "요약");
  await close(page);
  await cell(page, "A1", "='원본'!B1");
  await expect(page.locator("#selection-status")).toContainText("합계 42");

  await manage(page, "원본");
  await rename(page, "계산 원본");
  await close(page);
  let book = await backup(page, info, "renamed-reference.json");
  expect(sheet(book, "요약").cellData![0][0].f).toContain("계산 원본");
  expect(sheet(book, "요약").cellData![0][0].v).toBe(42);

  await manage(page, "계산 원본");
  await action(page, "duplicate");
  const copy = "계산 원본 (복사본)";
  await expect(page.locator("#sheet-manager-selected")).toHaveText(copy);
  await close(page);
  await expect(page.locator("#sheet-tabs button")).toHaveCount(3);
  await page.locator("#undo").click();
  await expect(page.locator("#sheet-tabs button")).toHaveCount(2);
  await page.locator("#redo").click();
  await expect(page.locator("#sheet-tabs button")).toHaveCount(3);
  await tab(page, copy);
  await cell(page, "A1", "7");
  await cell(page, "B1");
  await expect(page.locator("#cell-input")).toHaveValue("=A1*2");
  await expect(page.locator("#selection-status")).toContainText("합계 14");
  await tab(page, "요약");
  await cell(page, "A1");
  await expect(page.locator("#selection-status")).toContainText("합계 42");

  await manage(page, copy);
  await action(page, "left");
  await action(page, "color-2563eb");
  await close(page);
  book = await backup(page, info, "copied-and-reordered.json");
  expect(book.sheets[book.sheetOrder[0]].name).toBe(copy);
  expect(sheet(book, copy).tabColor).toBe("#2563eb");
  expect(sheet(book, copy).cellData![0][0].v).toBe(7);
  expect(sheet(book, "계산 원본").cellData![0][0].v).toBe(21);
  expect(sheet(book, "요약").cellData![0][0].v).toBe(42);

  await manage(page, copy);
  await action(page, "visibility");
  await action(page, "confirm-apply");
  await expect(page.locator("#sheet-manager-state")).toHaveText("숨긴 시트");
  await close(page);
  await expect(page.locator("#sheet-tabs button")).toHaveCount(2);
  await manage(page, copy);
  await action(page, "visibility");
  await close(page);
  await expect(page.locator("#sheet-tabs button")).toHaveCount(3);

  await manage(page, "계산 원본");
  await action(page, "delete");
  await expect(
    page.locator("#sheet-manager-confirm-description"),
  ).toContainText("#REF!");
  await action(page, "confirm-apply");
  await close(page);
  await expect(page.locator("#sheet-tabs button")).toHaveCount(2);
  book = await backup(page, info, "deleted-reference.json");
  expect(sheet(book, "요약").cellData![0][0].f).toContain("#REF!");
  await page.locator("#undo").click();
  await expect(page.locator("#sheet-tabs button")).toHaveCount(3);
  await tab(page, "요약");
  await cell(page, "A1");
  await expect(page.locator("#selection-status")).toContainText("합계 42");
  book = await backup(page, info, "undo-restored-reference.json");
  expect(sheet(book, "요약").cellData![0][0].f).toContain("계산 원본");
  expect(sheet(book, "계산 원본").cellData![0][0].v).toBe(21);
  await page.screenshot({ path: info.outputPath("sheet-command-undo.png") });
  expect(errors).toEqual([]);
});

test("hidden sheets remain manageable and the last visible sheet is protected at every viewport", async ({
  page,
}, info) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/");
  await ready(page);
  await manage(page);
  await expect(page.locator("#sheet-manager-delete")).toBeDisabled();
  await expect(page.locator("#sheet-manager-visibility")).toBeDisabled();
  await action(page, "duplicate");
  await action(page, "visibility");
  await page.keyboard.press("Escape");
  await expect(page.locator("#sheet-manager-confirm")).toBeHidden();
  await expect(page.locator("#sheet-manager-visibility")).toBeFocused();
  await action(page, "visibility");
  await action(page, "confirm-apply");
  await expect(page.locator("#sheet-manager-open")).toBeDisabled();
  await page
    .locator("#sheet-manager-list .sheet-manager-item")
    .filter({ has: page.getByText("시트1", { exact: true }) })
    .click();
  await expect(page.locator("#sheet-manager-delete")).toBeDisabled();
  await expect(page.locator("#sheet-manager-visibility")).toBeDisabled();
  for (const viewport of [
    { width: 360, height: 800 },
    { width: 390, height: 844 },
    { width: 768, height: 1024 },
    { width: 1440, height: 900 },
  ]) {
    await page.setViewportSize(viewport);
    await expect
      .poll(() =>
        page.evaluate(() => document.documentElement.scrollWidth <= innerWidth),
      )
      .toBe(true);
    await expect
      .poll(() =>
        page
          .locator("#sheet-manager-dialog")
          .evaluate((node) => node.scrollWidth <= node.clientWidth),
      )
      .toBe(true);
    for (const selector of [
      "#sheet-manager-name",
      "#sheet-manager-close",
      "#sheet-manager-open",
      "#sheet-manager-left",
      ".sheet-manager-color",
    ]) {
      const box = await page.locator(selector).first().boundingBox();
      expect(box?.height).toBeGreaterThanOrEqual(44);
    }
    await page.screenshot({
      path: info.outputPath(`sheet-manager-${viewport.width}.png`),
    });
  }
  await page.keyboard.press("Escape");
  await expect(page.locator("#sheet-manager-dialog")).toHaveCount(0);
  await expect(page.locator("#sheet-manage")).toBeFocused();
  await page.setViewportSize({ width: 1440, height: 900 });
  await expect(page.locator("#sheet-tabs button")).toHaveCount(1);
  await manage(page, "시트1 (복사본)");
  await action(page, "visibility");
  await close(page);
  await expect(page.locator("#sheet-tabs button")).toHaveCount(2);
  expect(errors).toEqual([]);
});
