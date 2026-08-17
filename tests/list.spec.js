const { expect, test } = require("@playwright/test");
const fixtures = require("./fixtures");
const { cardTitles, gotoSignedIn, openFilters, selectFlatList } = require("./helpers");

test.beforeEach(async function ({ page }) {
  await gotoSignedIn(page, {});
  await selectFlatList(page);
});

test("renders every row except the deleted ones", async function ({ page }) {
  await expect(cardTitles(page)).toHaveCount(fixtures.VISIBLE_NAMES.length);
  const titles = await cardTitles(page).allTextContents();
  expect(titles.sort()).toEqual(fixtures.VISIBLE_NAMES.slice().sort());
  await expect(page.locator("#items")).not.toContainText("Broken kettle");
});

test("search matches the searchable columns", async function ({ page }) {
  const search = page.locator("#search");

  await search.fill("drill");                 // Name
  await expect(cardTitles(page)).toHaveText(["Cordless drill"]);

  await search.fill("lighting");              // Category
  await expect(cardTitles(page)).toHaveText(["Desk lamp"]);

  await search.fill("living room");           // Room
  await expect(cardTitles(page)).toHaveText(["Armchair"]);

  await search.fill("adjustable");            // Notes
  await expect(cardTitles(page)).toHaveText(["Desk lamp"]);

  await search.fill("fragile");               // Tags
  await expect(cardTitles(page)).toHaveCount(2);
});

test("search does not match non-searchable columns", async function ({ page }) {
  const search = page.locator("#search");

  await search.fill("id-drill");              // ID
  await expect(cardTitles(page)).toHaveCount(0);
  await expect(page.locator("#empty")).toBeVisible();

  await search.fill("42");                    // Quantity
  await expect(cardTitles(page)).toHaveCount(0);

  await search.fill("2024-03-05");            // Purchased
  await expect(cardTitles(page)).toHaveCount(0);
});

test("category and room selects hold the distinct values and filter", async function ({ page }) {
  await openFilters(page);
  const category = page.locator("#filter-category");
  const room = page.locator("#filter-room");

  await expect(category.locator("option")).toHaveText([
    "All categories", "Furniture", "Lighting", "Tools"
  ]);
  await expect(room.locator("option")).toHaveText([
    "All rooms", "Living room", "Office"
  ]);

  await category.selectOption("Tools");
  await expect(cardTitles(page)).toHaveText(["Cordless drill"]);

  await category.selectOption("");
  await room.selectOption("Office");
  await expect(cardTitles(page)).toHaveCount(2);
});

test("tag toggles filter with AND semantics", async function ({ page }) {
  await openFilters(page);
  const tags = page.locator("#tag-filter-list .tag-toggle");
  await expect(tags).toHaveText(["fragile", "metal"]);

  await tags.filter({ hasText: "metal" }).click();
  await expect(cardTitles(page)).toHaveCount(2);

  await tags.filter({ hasText: "fragile" }).click();
  await expect(cardTitles(page)).toHaveText(["Desk lamp"]);

  await page.locator("#clear-tags-btn").click();
  await expect(cardTitles(page)).toHaveCount(3);
});

test("the filters badge counts the active filters", async function ({ page }) {
  await openFilters(page);
  const badge = page.locator("#filters-count");
  await expect(badge).toBeHidden();

  await page.locator("#filter-category").selectOption("Tools");
  await expect(badge).toHaveText("1");

  await page.locator("#filter-room").selectOption("Office");
  await expect(badge).toHaveText("2");

  await page.locator("#tag-filter-list .tag-toggle").filter({ hasText: "metal" }).click();
  await expect(badge).toHaveText("3");
});

test("clear all filters resets everything and is disabled when idle", async function ({ page }) {
  await openFilters(page);
  const clear = page.locator("#clear-filters-btn");
  await expect(clear).toBeDisabled();

  await page.locator("#search").fill("drill");
  await page.locator("#filter-category").selectOption("Tools");
  await page.locator("#tag-filter-list .tag-toggle").filter({ hasText: "metal" }).click();
  await expect(clear).toBeEnabled();

  await clear.click();
  await expect(page.locator("#filter-category")).toHaveValue("");
  await expect(page.locator("#filter-room")).toHaveValue("");
  await expect(page.locator("#tag-filter-list .tag-toggle[aria-pressed='true']")).toHaveCount(0);
  await expect(clear).toBeDisabled();
  await expect(page.locator("#filters-count")).toBeHidden();
  // The search box is deliberately not part of the filter count, so it keeps
  // its text and stays applied.
  await expect(page.locator("#search")).toHaveValue("drill");
  await expect(cardTitles(page)).toHaveText(["Cordless drill"]);
});

test("an empty result shows the empty message", async function ({ page }) {
  await page.locator("#search").fill("nothing matches this");
  await expect(cardTitles(page)).toHaveCount(0);
  await expect(page.locator("#empty")).toBeVisible();
});

test("the filters popover closes on outside click and Escape", async function ({ page }) {
  await openFilters(page);
  await page.locator("h1.brand").click();
  await expect(page.locator("#filters-panel")).toBeHidden();

  await openFilters(page);
  await page.keyboard.press("Escape");
  await expect(page.locator("#filters-panel")).toBeHidden();
});

test("the account popover closes on outside click and Escape", async function ({ page }) {
  await page.locator("#account-btn").click();
  await expect(page.locator("#account-panel")).toBeVisible();
  await page.locator("h1.brand").click();
  await expect(page.locator("#account-panel")).toBeHidden();

  await page.locator("#account-btn").click();
  await page.keyboard.press("Escape");
  await expect(page.locator("#account-panel")).toBeHidden();
});
