const { expect, test } = require("@playwright/test");
const fixtures = require("./fixtures");
const { setup, statusTexts } = require("./helpers");

test.describe("startup", function () {
  test("unconfigured config.js shows the not-configured message", async function ({ page }) {
    await setup(page, { config: fixtures.PLACEHOLDER_CONFIG });
    await page.goto("/");
    await expect(page.locator("#status")).toContainText("Not configured yet");
    await expect(page.locator("#status")).toHaveClass(/error/);
    await expect(page.locator("#toolbar")).toBeHidden();
  });

  test("missing CONFIG shows the config-file error", async function ({ page }) {
    await setup(page, { config: null });
    await page.goto("/");
    await expect(page.locator("#status")).toContainText("Configuration file (config.js) could not be loaded");
    await expect(page.locator("#toolbar")).toBeHidden();
  });

  test("missing msal global shows the MSAL-load error", async function ({ page }) {
    await setup(page, { msal: false });
    await page.goto("/");
    await expect(page.locator("#status")).toContainText("Microsoft sign-in library (MSAL) could not be loaded");
    await expect(page.locator("#toolbar")).toBeHidden();
  });

  test("signed out shows only the sign-in button", async function ({ page }) {
    await setup(page, { signedIn: false });
    await page.goto("/");
    await expect(page.locator("#signin-btn")).toBeVisible();
    await expect(page.locator("#status")).toContainText("Sign in with your Microsoft account");
    await expect(page.locator("#account-btn")).toBeHidden();
    await expect(page.locator("#toolbar")).toBeHidden();
    await expect(page.locator("#items")).toBeHidden();
    await expect(page.locator("#filters-panel")).toBeHidden();

    await page.locator("#signin-btn").click();
    await expect.poll(function () {
      return page.evaluate(function () { return window.__msalCalls; });
    }).toContain("loginRedirect");
  });

  test("a spreadsheet without a Status column shows an error and no rows", async function ({ page }) {
    const columns = fixtures.COLUMNS.filter(function (name) { return name !== "Status"; });
    await setup(page, { columns: columns });
    await page.goto("/");
    await expect(page.locator("#toolbar")).toBeVisible();
    await expect(page.locator("#items .card")).toHaveCount(0);
    // Characterized as-is: the message is shown, but the render that follows
    // immediately replaces it with the "no items yet" hint.
    await expect.poll(function () {
      return statusTexts(page);
    }).toEqual(expect.arrayContaining([
      expect.stringContaining('missing a "Status" column')
    ]));
  });

  test("signed in shows the account button with initials", async function ({ page }) {
    await setup(page, {});
    await page.goto("/");
    await expect(page.locator("#account-btn")).toBeVisible();
    await expect(page.locator("#account-initials")).toHaveText("AL");
    await expect(page.locator("#signin-btn")).toBeHidden();
    await page.locator("#account-btn").click();
    await expect(page.locator("#account-name")).toHaveText(fixtures.ACCOUNT.name);
    await expect(page.locator("#account-email")).toHaveText(fixtures.ACCOUNT.username);
    // No profile photo: the initials avatar stays.
    await expect(page.locator("#account-photo")).toBeHidden();
  });
});
