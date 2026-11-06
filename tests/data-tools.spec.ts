import { test, expect, type Page, type TestInfo } from "@playwright/test";
import { createRequire } from "node:module";
const ExcelJS = createRequire(import.meta.url)(
  "exceljs",
) as typeof import("exceljs");

async function ready(page: Page) {
  await expect(page.locator("#loading")).toBeHidden();
  await expect(page.locator("#editor canvas:visible").first()).toBeVisible();
}
async function select(page: Page, address: string) {
  await page.locator("#address").fill(address);
  await page.locator("#address").press("Enter");
  await expect(page.locator("#selection-status")).toContainText(address);
}
async function fixture(page: Page, info: TestInfo) {
  const source = new ExcelJS.Workbook(),
    sheet = source.addWorksheet("데이터 도구 검증");
  sheet.getCell("A1").value = "Apple apple";
  sheet.getCell("A2").value = "apple";
  sheet.getCell("A3").value = "APPLE";
  sheet.getCell("B1").value = 123;
  sheet.getCell("B2").value = { formula: "LOWER(A3)", result: "apple" };
  sheet.getCell("C1").value = " 00123 ";
  sheet.getCell("C2").value = " =SUM(B1:B1) ";
  sheet.getCell("C3").value = " 앞  뒤 ";
  sheet.getCell("D1").value = " untouched ";
  sheet.getCell("F1").value = "apple";
  sheet.getCell("F1").font = { bold: true };
  const path = info.outputPath("data-tools-source.xlsx");
  await source.xlsx.writeFile(path);
  await page.goto("/");
  await ready(page);
  await page.locator("#open").click();
  const chooser = page.waitForEvent("filechooser");
  await page.locator("#modal-ok").click();
  await (await chooser).setFiles(path);
  await expect(page.locator("#toast")).toContainText("파일을 열었습니다");
  await ready(page);
  await expect(page.locator("#save-state")).toHaveText("이 브라우저에 보관됨");
}
async function exportBook(page: Page, info: TestInfo, name: string) {
  const event = page.waitForEvent("download");
  await page.locator("#export").click();
  const download = await event;
  const path = info.outputPath(name);
  await download.saveAs(path);
  const book = new ExcelJS.Workbook();
  await book.xlsx.readFile(path);
  return book;
}
async function openFind(page: Page) {
  await page.locator("#tab-2").click();
  await page.locator("#find").click();
  await expect(page.locator("#data-tools-dialog")).toBeVisible();
}
const action = (page: Page, name: string) =>
  page.locator(`#data-tools-dialog [data-action="${name}"]`);
async function close(page: Page) {
  await page.locator("#data-tools-dialog").press("Escape");
  await expect(page.locator("#data-tools-dialog")).toHaveCount(0);
}

test("find options, result navigation, reviewed replace-all and one-step undo preserve formulas and style", async ({
  page,
}, info) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await fixture(page, info);
  await openFind(page);
  await page.locator("#data-find-query").fill("apple");
  await page
    .locator("#data-find-form")
    .getByRole("button", { name: "찾기", exact: true })
    .click();
  await expect(page.locator("#data-find-count")).toContainText("5개 셀");
  await page.locator("#data-case").check();
  await page.locator("#data-whole").check();
  await expect(page.locator("#data-find-count")).toContainText("3개 셀");
  await action(page, "next").click();
  await expect(page.locator("#selection-status")).toContainText("A2");
  await action(page, "next").click();
  await expect(page.locator("#selection-status")).toContainText("B2");
  await expect(action(page, "replace-one")).toBeDisabled();
  await action(page, "previous").click();
  await expect(page.locator("#selection-status")).toContainText("A2");
  await page.locator("#data-case").uncheck();
  await page.locator("#data-whole").uncheck();
  await page.locator("#data-replacement").fill("=SUM(1,2)");
  await action(page, "replace-all").click();
  await expect(page.locator("#data-confirm-title")).toHaveText("4개 셀 바꾸기");
  await expect(page.locator("#data-confirm-caption")).toContainText(
    "보호된 1개 셀",
  );
  await action(page, "apply-plan").click();
  await expect(page.locator("#toast")).toContainText("4개 텍스트 셀");
  await close(page);
  await expect(page.locator("#find")).toBeFocused();
  const changed = (await exportBook(page, info, "replaced.xlsx")).worksheets[0];
  expect(changed.getCell("A1").value).toBe("=SUM(1,2) =SUM(1,2)");
  expect(changed.getCell("A2").value).toBe("=SUM(1,2)");
  expect(changed.getCell("A2").formula).toBeUndefined();
  expect(changed.getCell("B2").formula).toBe("LOWER(A3)");
  expect(changed.getCell("B1").value).toBe(123);
  expect(changed.getCell("F1").font.bold).toBe(true);
  await page.locator("#tab-0").click();
  await page.locator("#undo").click();
  const restored = (await exportBook(page, info, "one-undo.xlsx"))
    .worksheets[0];
  expect(restored.getCell("A1").value).toBe("Apple apple");
  expect(restored.getCell("A2").value).toBe("apple");
  expect(restored.getCell("A3").value).toBe("APPLE");
  expect(restored.getCell("F1").value).toBe("apple");
  expect(restored.getCell("B2").formula).toBe("LOWER(A3)");
  expect(errors).toEqual([]);
});

test("captured selection stays fixed during navigation and replace-one keeps leading zero text", async ({
  page,
}, info) => {
  await fixture(page, info);
  await select(page, "A1:A2");
  await openFind(page);
  await page.locator("#data-scope").selectOption("selection");
  await page.locator("#data-find-query").fill("apple");
  await page
    .locator("#data-find-form")
    .getByRole("button", { name: "찾기", exact: true })
    .click();
  await expect(page.locator("#data-find-count")).toContainText("2개 셀");
  await action(page, "next").click();
  await expect(page.locator("#selection-status")).toContainText("A2");
  await expect(page.locator("#data-scope option:checked")).toHaveText(
    "열 때 선택한 범위 (A1:A2)",
  );
  await page.locator("#data-replacement").fill("00123");
  await action(page, "replace-one").click();
  await expect(page.locator("#data-confirm-title")).toHaveText("1개 셀 바꾸기");
  await action(page, "apply-plan").click();
  await close(page);
  const changed = (await exportBook(page, info, "one-cell.xlsx")).worksheets[0];
  expect(changed.getCell("A1").value).toBe("Apple apple");
  expect(changed.getCell("A2").value).toBe("00123");
  expect(changed.getCell("A3").value).toBe("APPLE");
  expect(changed.getCell("B2").formula).toBe("LOWER(A3)");
});

test("mobile trim preview applies only captured text range and keeps formula-like text literal", async ({
  page,
}, info) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await fixture(page, info);
  await select(page, "C1:C3");
  await openFind(page);
  await page.locator("#data-tab-trim").click();
  await page.locator("#data-scope").selectOption("selection");
  await action(page, "preview-trim").click();
  await expect(page.locator("#data-confirm-title")).toHaveText(
    "3개 셀 공백 정리",
  );
  await page.screenshot({
    path: info.outputPath("data-tools-mobile-trim.png"),
  });
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= window.innerWidth,
    ),
  ).toBe(true);
  const box = await page.locator("#data-tools-dialog").boundingBox();
  expect(box!.width).toBeLessThanOrEqual(360);
  await action(page, "apply-plan").click();
  await expect(page.locator("#toast")).toContainText("3개 텍스트 셀");
  await close(page);
  const changed = (await exportBook(page, info, "trimmed.xlsx")).worksheets[0];
  expect(changed.getCell("C1").value).toBe("00123");
  expect(changed.getCell("C2").value).toBe("=SUM(B1:B1)");
  expect(changed.getCell("C2").formula).toBeUndefined();
  expect(changed.getCell("C3").value).toBe("앞  뒤");
  expect(changed.getCell("D1").value).toBe(" untouched ");
});
