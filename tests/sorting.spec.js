const { expect, test } = require("@playwright/test");
const { cardTitles, gotoSignedIn, selectFlatList } = require("./helpers");

test.describe("sorting", function () {
  test.beforeEach(async function ({ page }) {
    await gotoSignedIn(page, {});
    await selectFlatList(page);
  });

  test("the sort field and direction reorder the list", async function ({ page }) {
    // Default: Updated, descending.
    await expect(cardTitles(page)).toHaveText(["Cordless drill", "Armchair", "Desk lamp"]);

    await page.locator("#sort-direction-btn").click();
    await expect(cardTitles(page)).toHaveText(["Desk lamp", "Armchair", "Cordless drill"]);

    await page.locator("#sort-field").selectOption("Name");
    await expect(cardTitles(page)).toHaveText(["Armchair", "Cordless drill", "Desk lamp"]);

    await page.locator("#sort-direction-btn").click();
    await expect(cardTitles(page)).toHaveText(["Desk lamp", "Cordless drill", "Armchair"]);
    await expect(page.locator("#sort-direction-icon")).toHaveText("↓");
    await expect(page.locator("#sort-direction-btn")).toHaveAttribute("aria-label", "Sort direction: descending");
  });

  test("the sort and group selection survive a reload", async function ({ page }) {
    await page.locator("#sort-field").selectOption("CreatedAt");
    await page.locator("#sort-direction-btn").click();

    await page.reload();
    await expect(page.locator("#sort-field")).toHaveValue("CreatedAt");
    await expect(page.locator("#sort-direction-btn")).toHaveAttribute("aria-label", "Sort direction: ascending");
    await expect(page.locator("#group-field")).toHaveValue("");
  });
});

test.describe("grouping", function () {
  test.beforeEach(async function ({ page }) {
    await gotoSignedIn(page, {});
  });

  test("groups render as collapsed panels", async function ({ page }) {
    const panels = page.locator("#items .category-panel");
    await expect(panels).toHaveCount(3);
    await expect(page.locator("#items .category-name")).toHaveText(["Tools", "Furniture", "Lighting"]);
    await expect(page.locator("#items .category-panel[open]")).toHaveCount(0);
  });

  test("expand all and collapse all work and are hidden in flat mode", async function ({ page }) {
    const panels = page.locator("#items .category-panel");
    await page.locator("#expand-all-btn").click();
    await expect(page.locator("#items .category-panel[open]")).toHaveCount(await panels.count());

    await page.locator("#collapse-all-btn").click();
    await expect(page.locator("#items .category-panel[open]")).toHaveCount(0);

    await selectFlatList(page);
    await expect(page.locator("#expand-all-btn")).toBeHidden();
    await expect(page.locator("#collapse-all-btn")).toBeHidden();
    await expect(page.locator("#items .cards")).toHaveCount(1);
  });

  test("the expanded state is kept per grouping mode", async function ({ page }) {
    await page.locator("#items .category-panel").first().locator("summary").click();
    await expect(page.locator("#items .category-panel[open]")).toHaveCount(1);

    await page.locator("#group-field").selectOption("Room");
    await expect(page.locator("#items .category-panel[open]")).toHaveCount(0);

    await page.locator("#group-field").selectOption("Category");
    await expect(page.locator("#items .category-panel[open]")).toHaveCount(1);
  });

  test("group counts show how much the filter hides", async function ({ page }) {
    await page.locator("#search").fill("drill");
    await expect(page.locator("#items .category-name")).toHaveText(["Tools"]);
    await expect(page.locator("#items .category-count")).toHaveText(["(1)"]);
  });
});
