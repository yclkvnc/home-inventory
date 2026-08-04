// Playwright configuration for the smoke suite.
//
// The app has no build step, so the tests run against the repository served as
// plain static files.
const { defineConfig, devices } = require("@playwright/test");

const PORT = 4173;

module.exports = defineConfig({
  testDir: "./tests",
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? [["html", { open: "never" }], ["list"]] : "list",
  use: {
    baseURL: "http://localhost:" + PORT,
    trace: "on-first-retry"
  },
  projects: [
    { name: "chromium", use: Object.assign({}, devices["Desktop Chrome"]) }
  ],
  webServer: {
    command: "python3 -m http.server " + PORT,
    url: "http://localhost:" + PORT + "/index.html",
    reuseExistingServer: !process.env.CI
  }
});
