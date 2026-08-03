const { expect, test } = require("@playwright/test");
const { setup } = require("./helpers");

test("the toggle switches the theme, its state and its icon", async function ({ page }) {
  await setup(page, {});
  await page.emulateMedia({ colorScheme: "light" });
  await page.goto("/");

  const toggle = page.locator("#theme-toggle-btn");
  await expect(page.locator("html")).toHaveAttribute("data-theme", "light");
  await expect(toggle).toHaveAttribute("aria-pressed", "false");
  await expect(page.locator("#theme-toggle-icon")).toHaveText("☾");

  await toggle.click();
  await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  await expect(toggle).toHaveAttribute("aria-pressed", "true");
  await expect(toggle).toHaveAttribute("aria-label", "Switch to light theme");
  await expect(page.locator("#theme-toggle-icon")).toHaveText("☀");
});

test("the choice survives a reload", async function ({ page }) {
  await setup(page, {});
  await page.emulateMedia({ colorScheme: "light" });
  await page.goto("/");
  await page.locator("#theme-toggle-btn").click();

  await page.reload();
  await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  await expect(page.locator("#theme-toggle-btn")).toHaveAttribute("aria-pressed", "true");
});

test("without a choice the app follows the OS preference", async function ({ page }) {
  await setup(page, {});
  await page.emulateMedia({ colorScheme: "dark" });
  await page.goto("/");
  await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");

  await page.emulateMedia({ colorScheme: "light" });
  await expect(page.locator("html")).toHaveAttribute("data-theme", "light");
});
