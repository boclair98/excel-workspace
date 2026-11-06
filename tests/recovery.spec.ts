import { test, expect, type Page } from "@playwright/test";

async function ready(page: Page) {
  await expect(page.locator("#tab-0")).toBeVisible();
  await expect(page.locator("#loading")).toBeHidden();
  await expect(page.locator("#editor canvas:visible").first()).toBeVisible();
  await expect(page.locator("#history")).toBeVisible();
}

async function selectA1(page: Page) {
  await page.locator("#address").fill("A1");
  await page.locator("#address").press("Enter");
  await expect(page.locator("#selection-status")).toContainText("A1");
}

async function editA1(page: Page, value: string) {
  await selectA1(page);
  await page.locator("#cell-input").fill(value);
  await page.locator('#cell-form button[type="submit"]').click();
}

async function saveA1(page: Page, value: string) {
  await editA1(page, value);
  await page.locator("#save").click();
  await expect(page.locator("#save-state")).toHaveText("이 브라우저에 보관됨");
}

async function rename(page: Page, name: string) {
  await page.locator("#rename").click();
  await page.locator("#new-name").fill(name);
  await page.locator("#modal-ok").click();
  await expect(page.locator("#document-name")).toHaveText(name);
  await expect(page.locator("#save-state")).toHaveText("이 브라우저에 보관됨");
}

async function expectA1(page: Page, value: string) {
  await selectA1(page);
  await expect(page.locator("#cell-input")).toHaveValue(value);
}

async function openSavedBook(page: Page, name: string) {
  await page
    .locator("#library .book-open")
    .filter({
      has: page.getByText(name, { exact: true }),
    })
    .click();
  await expect(page.locator("#document-name")).toHaveText(name);
  await ready(page);
}

test("named checkpoint restores a persistent copy while retaining later original edits", async ({
  page,
}, info) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  const original = "복구 검증 문서";
  const restored = `${original} · 복구 사본`;
  await page.goto("/");
  await ready(page);
  await rename(page, original);
  await saveA1(page, "복구 지점의 내용");

  await page.locator("#history").click();
  await expect(page.locator("#history-empty")).toBeVisible();
  await page.locator("#history-name").fill("수정하기 전");
  await page.locator("#history-save").click();
  await expect(page.locator("#history-count")).toHaveText("1 / 10");
  await expect(page.locator("#history-list strong")).toHaveText("수정하기 전");
  await expect(page.locator("#history-error")).toBeHidden();
  await expect(page.locator("#history-save")).toBeEnabled();
  await page.keyboard.press("Escape");
  await expect(page.locator("#history-dialog")).toHaveCount(0);
  await expect(page.locator("#history")).toBeFocused();

  await saveA1(page, "원본의 나중 변경");
  await page.locator("#history").click();
  const restore = page.getByRole("button", {
    name: "수정하기 전 사본으로 복원",
    exact: true,
  });
  await restore.click();
  await expect(page.locator("#history-confirm-title")).toHaveText(
    "복구 사본을 열까요?",
  );
  await page.keyboard.press("Escape");
  await expect(page.locator("#history-confirm")).toBeHidden();
  await expect(page.locator("#history-dialog")).toBeVisible();
  await expect(restore).toBeFocused();
  await restore.click();
  await page.locator("#history-confirm-apply").click();
  await expect(page.locator("#history-dialog")).toHaveCount(0);
  await expect(page.locator("#document-name")).toHaveText(restored);
  await expect(page.locator("#save-state")).toHaveText("이 브라우저에 보관됨");
  await expectA1(page, "복구 지점의 내용");
  await expect(page.locator("#library .book-open")).toHaveCount(2);
  await page.screenshot({
    path: info.outputPath("checkpoint-restored-copy.png"),
  });

  // Reopening and reloading verifies durable snapshots, not only the current canvas.
  await page.reload();
  await ready(page);
  await expect(page.locator("#document-name")).toHaveText(restored);
  await expectA1(page, "복구 지점의 내용");
  await openSavedBook(page, original);
  await expectA1(page, "원본의 나중 변경");
  await page.locator("#history").click();
  await expect(page.locator("#history-count")).toHaveText("1 / 10");
  await expect(page.locator("#history-list strong")).toHaveText("수정하기 전");
  await page.keyboard.press("Escape");
  await openSavedBook(page, restored);
  await expectA1(page, "복구 지점의 내용");
  expect(errors).toEqual([]);
});

test("a stale tab cannot overwrite another save and preserves its edits as a separate copy", async ({
  page,
  context,
}, info) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  const original = "동시 편집 검증";
  const copy = `${original} 내 변경 사본`;
  await page.goto("/");
  await ready(page);
  await rename(page, original);
  await saveA1(page, "두 탭이 처음 연 내용");

  const stale = await context.newPage();
  stale.on("pageerror", (error) => errors.push(error.message));
  await stale.goto("/");
  await ready(stale);
  await expect(stale.locator("#document-name")).toHaveText(original);
  await expectA1(stale, "두 탭이 처음 연 내용");

  await saveA1(page, "먼저 보관한 탭의 내용");
  await editA1(stale, "충돌한 탭에서 보존할 내용");
  await stale.locator("#save").click();
  await expect(stale.locator("#save-conflict")).toBeVisible();
  await expect(stale.locator("#save-state")).toContainText("저장 충돌");
  await expectA1(stale, "충돌한 탭에서 보존할 내용");
  await stale.screenshot({ path: info.outputPath("stale-tab-conflict.png") });

  // The stale tab's failed save must leave the first tab's durable version intact.
  await page.reload();
  await ready(page);
  await expectA1(page, "먼저 보관한 탭의 내용");
  await stale.locator("#conflict-copy").click();
  await expect(stale.locator("#save-conflict")).toBeHidden();
  await expect(stale.locator("#document-name")).toHaveText(copy);
  await expect(stale.locator("#save-state")).toHaveText("이 브라우저에 보관됨");
  await expect(stale.locator("#library .book-open")).toHaveCount(2);
  await expectA1(stale, "충돌한 탭에서 보존할 내용");
  await stale.screenshot({
    path: info.outputPath("conflict-preserved-copy.png"),
  });

  await stale.reload();
  await ready(stale);
  // Open explicitly because both tabs share the browser's last-opened preference.
  await openSavedBook(stale, copy);
  await expectA1(stale, "충돌한 탭에서 보존할 내용");
  await openSavedBook(stale, original);
  await expectA1(stale, "먼저 보관한 탭의 내용");
  await expect(stale.locator("#save-conflict")).toBeHidden();
  expect(errors).toEqual([]);
  await stale.close();
});
