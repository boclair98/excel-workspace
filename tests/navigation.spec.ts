import { test, expect, type Page } from "@playwright/test";

async function ready(page: Page) {
  await expect(page.locator("#loading")).toBeHidden();
  await expect(
    page.locator('#sheet-tabs [aria-selected="true"]'),
  ).toBeVisible();
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

test("sheet navigation, renaming, numeric summary and formula help work on desktop and mobile", async ({
  page,
}, info) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/");
  await ready(page);
  await cell(page, "A1", "10");
  await cell(page, "A2", "20");
  await cell(page, "A1:A2");
  await expect(page.locator("#selection-status")).toContainText(
    "숫자 2개 · 합계 30 · 평균 15",
  );
  await page.locator("#cell-input").fill("=SUM(A1:A2)");
  await expect(page.locator("#formula-help")).toContainText("SUM(범위)");
  // Leave input without submitting; the existing source numbers stay unchanged.
  await page.locator("#address").focus();
  await expect(page.locator("#formula-help")).toBeHidden();
  const firstId = await page
    .locator("#sheet-tabs button")
    .first()
    .getAttribute("data-sheet");
  await page.locator("#sheet-add").click();
  await expect(page.locator("#sheet-tabs button")).toHaveCount(2);
  await page.locator("#sheet-rename").click();
  await page.locator("#sheet-name").fill("검토 자료");
  await page.locator("#modal-ok").click();
  await expect(page.locator('#sheet-tabs [aria-selected="true"]')).toHaveText(
    "검토 자료",
  );
  await cell(page, "A1", "두 번째 시트");
  await page.locator('#sheet-tabs [aria-selected="true"]').focus();
  await page.keyboard.press("ArrowLeft");
  await expect(
    page.locator('#sheet-tabs [aria-selected="true"]'),
  ).toBeFocused();
  await cell(page, "A1");
  await expect(page.locator("#cell-input")).toHaveValue("10");
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
    for (const id of ["#sheet-add", "#sheet-rename", "#sheet-zoom"]) {
      const box = await page.locator(id).boundingBox();
      expect(box?.height).toBeGreaterThanOrEqual(44);
    }
    if (viewport.width <= 800) {
      await expect(page.locator("#sheet-picker")).toBeVisible();
      await page.locator("#sheet-picker").selectOption({ label: "검토 자료" });
      await cell(page, "A1");
      await expect(page.locator("#cell-input")).toHaveValue("두 번째 시트");
      await page.locator("#sheet-zoom").selectOption("1.25");
      await expect(page.locator("#zoom")).toHaveValue("1.25");
      await page.locator("#sheet-picker").selectOption(firstId!);
      await cell(page, "A1");
      await expect(page.locator("#cell-input")).toHaveValue("10");
    }
    await page.screenshot({
      path: info.outputPath(`navigation-${viewport.width}.png`),
    });
  }
  await page.reload();
  await ready(page);
  await expect(page.locator("#sheet-tabs button")).toHaveCount(2);
  expect(errors).toEqual([]);
});
