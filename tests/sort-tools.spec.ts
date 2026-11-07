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
async function fixture(page: Page, info: TestInfo, merge = false) {
  const source = new ExcelJS.Workbook(),
    sheet = source.addWorksheet("다중 기준 정렬");
  const headers = ["부서", "점수", "항목", "수량", "단가", "금액"];
  headers.forEach((value, i) => {
    const cell = sheet.getCell(2, i + 2);
    cell.value = value;
    cell.font = { bold: true };
  });
  const rows = [
    ["B", 70, "B70", 2, 10],
    ["A", 80, "A80-low", 3, 20],
    ["A", 95, "A95", 4, 30],
    ["B", 90, "B90-high", 5, 40],
    ["A", 80, "A80-high", 9, 50],
    ["B", 90, "B90-low", 1, 60],
  ];
  rows.forEach((values, index) => {
    const row = index + 3;
    sheet.getCell(row, 1).value = `고정-${row}`;
    values.forEach((value, col) => (sheet.getCell(row, col + 2).value = value));
    sheet.getCell(row, 7).value = {
      formula: `E${row}*F${row}`,
      result: Number(values[3]) * Number(values[4]),
    };
  });
  sheet.getCell("I1").value = "범위 밖";
  if (merge) sheet.mergeCells("B3:C3");
  const path = info.outputPath(merge ? "merged.xlsx" : "sort-source.xlsx");
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
async function openSort(page: Page, address: string) {
  await select(page, address);
  await page.locator("#tab-2").click();
  await page.locator("#sort").click();
  await expect(page.locator("#sort-tools-dialog")).toBeVisible();
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
const rowNames = (sheet: import("exceljs").Worksheet) =>
  Array.from({ length: 6 }, (_, i) => sheet.getCell(i + 3, 4).value);

test("non-A range sorts by three relative criteria, preserves title and row formulas, then undoes once", async ({
  page,
}, info) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await fixture(page, info);
  await openSort(page, "B2:G8");
  await expect(page.locator("#sort-has-header")).toBeChecked();
  await expect(page.locator("#sort-column-0 option:checked")).toHaveText(
    "B열 · 부서",
  );
  await expect(page.locator("#sort-target-summary")).toContainText("B3:G8");
  await page.locator("#sort-add-criterion").click();
  await page.locator("#sort-column-1").selectOption("1");
  await page.locator("#sort-direction-1").selectOption("desc");
  await page.locator("#sort-add-criterion").click();
  await page.locator("#sort-column-2").selectOption("3");
  await page.locator("#sort-direction-2").selectOption("desc");
  await expect(page.locator("#sort-add-criterion")).toBeDisabled();
  await expect(page.locator("#sort-order-summary")).toContainText(
    "C열 · 점수 — 내림차순",
  );
  await expect(page.locator("#sort-order-summary")).toContainText(
    "E열 · 수량 — 내림차순",
  );
  await page.screenshot({
    path: info.outputPath("sort-three-criteria-desktop.png"),
  });
  await page.locator("#sort-apply").click();
  await expect(page.locator("#sort-tools-dialog")).toHaveCount(0);
  await expect(page.locator("#toast")).toContainText("6개 행을 정렬했습니다");
  await expect(page.locator("#sort")).toBeFocused();
  await expect(page.locator("#save-state")).toHaveText("이 브라우저에 보관됨");
  const sorted = (await exportBook(page, info, "sorted.xlsx")).worksheets[0];
  expect(rowNames(sorted)).toEqual([
    "A95",
    "A80-high",
    "A80-low",
    "B90-high",
    "B90-low",
    "B70",
  ]);
  expect(sorted.getCell("B2").value).toBe("부서");
  expect(sorted.getCell("G2").value).toBe("금액");
  expect(sorted.getCell("B2").font.bold).toBe(true);
  for (let row = 3; row <= 8; row++) {
    expect(sorted.getCell(row, 1).value).toBe(`고정-${row}`);
    expect(sorted.getCell(row, 7).formula).toBe(`E${row}*F${row}`);
    expect(sorted.getCell(row, 7).result).toBe(
      Number(sorted.getCell(row, 5).value) *
        Number(sorted.getCell(row, 6).value),
    );
  }
  expect(sorted.getCell("I1").value).toBe("범위 밖");
  await page.locator("#tab-0").click();
  await page.locator("#undo").click();
  await expect(page.locator("#save-state")).toHaveText("이 브라우저에 보관됨");
  const restored = (await exportBook(page, info, "sort-undone.xlsx"))
    .worksheets[0];
  expect(rowNames(restored)).toEqual([
    "B70",
    "A80-low",
    "A95",
    "B90-high",
    "A80-high",
    "B90-low",
  ]);
  for (let row = 3; row <= 8; row++)
    expect(restored.getCell(row, 7).formula).toBe(`E${row}*F${row}`);
  expect(errors).toEqual([]);
});

test("header-free sorting includes first selected row and criterion remove preserves a usable configuration", async ({
  page,
}, info) => {
  await fixture(page, info);
  await openSort(page, "B3:G8");
  await page.locator("#sort-has-header").uncheck();
  await expect(page.locator("#sort-column-0 option:checked")).toHaveText("B열");
  await page.locator("#sort-column-0").selectOption("1");
  await page.locator("#sort-direction-0").selectOption("desc");
  await page.locator("#sort-add-criterion").click();
  await page
    .getByRole("button", { name: "2번째 정렬 기준 삭제", exact: true })
    .click();
  await expect(page.locator(".sort-criterion")).toHaveCount(1);
  await expect(page.locator("#sort-target-summary")).toContainText("6개 행");
  await page.locator("#sort-apply").click();
  await expect(page.locator("#sort-tools-dialog")).toHaveCount(0);
  const sorted = (await exportBook(page, info, "header-free.xlsx"))
    .worksheets[0];
  expect(
    Array.from({ length: 6 }, (_, i) => sorted.getCell(i + 3, 3).value),
  ).toEqual([95, 90, 90, 80, 80, 70]);
  expect(sorted.getCell("B2").value).toBe("부서");
  expect(sorted.getCell("D8").value).toBe("B70");
});

test("mobile sort controls fit and invalid or merged ranges remain unchanged", async ({
  page,
}, info) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await fixture(page, info);
  await openSort(page, "B2:G8");
  await page.locator("#sort-add-criterion").click();
  await page.locator("#sort-add-criterion").click();
  await page.screenshot({ path: info.outputPath("sort-mobile.png") });
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= window.innerWidth,
    ),
  ).toBe(true);
  const box = await page.locator("#sort-tools-dialog").boundingBox();
  expect(box!.width).toBeLessThanOrEqual(360);
  await page.locator("#sort-tools-dialog").press("Escape");
  await openSort(page, "B2:G2");
  await expect(page.locator("#sort-apply")).toBeDisabled();
  await expect(page.locator("#sort-error")).toContainText("두 줄 이상");
  await page.locator("#sort-tools-dialog").press("Escape");
  await fixture(page, info, true);
  await openSort(page, "B2:G8");
  await expect(page.locator("#sort-apply")).toBeDisabled();
  await expect(page.locator("#sort-error")).toContainText("병합된 셀");
  await page.locator("#sort-tools-dialog").press("Escape");
  const unchanged = (await exportBook(page, info, "merged-unchanged.xlsx"))
    .worksheets[0];
  expect(rowNames(unchanged)).toEqual([
    "B70",
    "A80-low",
    "A95",
    "B90-high",
    "A80-high",
    "B90-low",
  ]);
});
