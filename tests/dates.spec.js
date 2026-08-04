const { expect, test } = require("@playwright/test");
const fixtures = require("./fixtures");
const { gotoSignedIn, selectFlatList } = require("./helpers");

// A zone west of UTC would shift a date-only value to the previous day if it
// were parsed as UTC, so it is the interesting case to pin down.
test.use({ timezoneId: "America/Los_Angeles" });

const DATE_ROW = Object.assign({}, fixtures.ROWS[0], {
  ID: "id-dates",
  Name: "Dated item",
  Notes: "not a date",
  Purchased: "2024-03-05",
  CreatedAt: "2024-03-05T08:30:45.000Z",
  UpdatedAt: "2024-03-05T08:30:45.000Z"
});

function detail(page, label) {
  return page.locator("#items .card dt", { hasText: new RegExp("^" + label + "$") }).locator("+ dd");
}

test("a date-only value is not shifted by a day", async function ({ page }) {
  await gotoSignedIn(page, { rows: [DATE_ROW] });
  await selectFlatList(page);
  await expect(detail(page, "Purchased")).toHaveText("05.03.2024");
});

test("timestamps render as dd.MM.yyyy HH:mm:ss in the local zone", async function ({ page }) {
  await gotoSignedIn(page, { rows: [DATE_ROW] });
  await selectFlatList(page);
  // 2024-03-05T08:30:45Z is 00:30:45 in America/Los_Angeles (UTC-8).
  await expect(detail(page, "CreatedAt")).toHaveText("05.03.2024 00:30:45");
  await expect(detail(page, "UpdatedAt")).toHaveText("05.03.2024 00:30:45");
});

test("non-date values pass through unchanged", async function ({ page }) {
  await gotoSignedIn(page, { rows: [DATE_ROW] });
  await selectFlatList(page);
  await expect(detail(page, "Notes")).toHaveText("not a date");
  await expect(detail(page, "Quantity")).toHaveText(DATE_ROW.Quantity);
});
