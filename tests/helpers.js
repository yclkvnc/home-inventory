// Shared setup for the Playwright smoke suite.
//
// The app is served as plain static files, so everything it depends on is
// injected by intercepting requests:
//
//   * `config.js`  — replaced with the test CONFIG, since the checked-in file
//                    is not meant to be edited by the tests.
//   * MSAL         — the vendored bundle is replaced with a stub that returns a
//                    fake account and a dummy token. Real redirect flows are
//                    out of scope.
//   * Graph        — every `graph.microsoft.com` call is answered from fixtures.

const { expect } = require("@playwright/test");
const fixtures = require("./fixtures");

const GRAPH_GLOB = "**graph.microsoft.com/**";

function configScript(config) {
  return "var CONFIG = " + JSON.stringify(config) + ";\n" +
    "window.CONFIG = CONFIG;\n";
}

// Minimal stand-in for the vendored MSAL bundle. Only the members app.js uses
// are implemented.
function msalScript(options) {
  const account = options.account;
  return "window.msal = (function () {\n" +
    "  var account = " + JSON.stringify(account) + ";\n" +
    "  function PublicClientApplication() {}\n" +
    "  PublicClientApplication.prototype.initialize = function () { return Promise.resolve(); };\n" +
    "  PublicClientApplication.prototype.handleRedirectPromise = function () { return Promise.resolve(null); };\n" +
    "  PublicClientApplication.prototype.getAllAccounts = function () { return account ? [account] : []; };\n" +
    "  PublicClientApplication.prototype.setActiveAccount = function () {};\n" +
    "  PublicClientApplication.prototype.acquireTokenSilent = function () {\n" +
    "    return Promise.resolve({ accessToken: 'test-token' });\n" +
    "  };\n" +
    "  PublicClientApplication.prototype.acquireTokenRedirect = function () { return Promise.resolve(); };\n" +
    "  PublicClientApplication.prototype.loginRedirect = function () {\n" +
    "    window.__msalCalls.push('loginRedirect');\n" +
    "    return Promise.resolve();\n" +
    "  };\n" +
    "  PublicClientApplication.prototype.logoutRedirect = function () {\n" +
    "    window.__msalCalls.push('logoutRedirect');\n" +
    "    return Promise.resolve();\n" +
    "  };\n" +
    "  return { PublicClientApplication: PublicClientApplication };\n" +
    "}());\n";
}

function json(route, body, status) {
  return route.fulfill({
    status: status || 200,
    contentType: "application/json",
    body: JSON.stringify(body)
  });
}

/**
 * Installs the config, MSAL and Graph stubs on `page`.
 *
 * Options:
 *   config       — CONFIG object to serve, or null to serve an empty config.js
 *   msal         — false to serve an empty MSAL bundle (library failed to load)
 *   signedIn     — whether MSAL reports an existing account (default true)
 *   columns      — Excel column names (default fixtures.COLUMNS)
 *   rows         — Excel rows as objects (default fixtures.ROWS)
 *   accountPhoto — whether /me/photo/$value returns an image (default false)
 *   photoStatus  — status for item photo downloads (default 200)
 *   writeStatus  — status for rows/add and rows/itemAt (default 200)
 *
 * Returns a handle exposing the Graph requests the app issued.
 */
async function setup(page, options) {
  const opts = options || {};
  const config = opts.config === undefined ? fixtures.CONFIG : opts.config;
  const columns = opts.columns || fixtures.COLUMNS;
  const rows = opts.rows || fixtures.ROWS;
  const signedIn = opts.signedIn === undefined ? true : opts.signedIn;
  const graphRequests = [];

  await page.addInitScript(function () {
    window.__msalCalls = [];
    // Some status messages are replaced right away by the next render, so the
    // suite records every text the status line ever showed.
    window.__statusTexts = [];
    document.addEventListener("DOMContentLoaded", function () {
      var status = document.getElementById("status");
      if (!status) return;
      window.__statusTexts.push(status.textContent);
      new MutationObserver(function () {
        window.__statusTexts.push(status.textContent);
      }).observe(status, { childList: true, characterData: true, subtree: true });
    });
  });

  await page.route("**/config.js", function (route) {
    route.fulfill({
      status: 200,
      contentType: "application/javascript",
      body: config ? configScript(config) : ""
    });
  });

  await page.route("**/vendor/msal-browser.min.js", function (route) {
    route.fulfill({
      status: 200,
      contentType: "application/javascript",
      body: opts.msal === false ? "" : msalScript({ account: signedIn ? fixtures.ACCOUNT : null })
    });
  });

  await page.route(GRAPH_GLOB, function (route) {
    const request = route.request();
    const url = request.url();
    let body = null;
    try {
      body = request.postDataJSON();
    } catch (e) { /* raw upload body */ }
    graphRequests.push({ method: request.method(), url: url, body: body });

    if (/\/workbook\/tables\/[^/]+\/columns/.test(url)) {
      return json(route, fixtures.toGraphColumns(columns));
    }
    if (/\/rows\/add$/.test(url) || /\/rows\/itemAt\(index=\d+\)$/.test(url)) {
      if (opts.writeStatus && opts.writeStatus !== 200) {
        return json(route, { error: { message: "Write rejected." } }, opts.writeStatus);
      }
      return json(route, {});
    }
    if (/\/workbook\/tables\/[^/]+\/rows/.test(url)) {
      return json(route, fixtures.toGraphRows(columns, rows));
    }
    if (/\/me\/photo\/\$value$/.test(url)) {
      if (!opts.accountPhoto) return json(route, { error: { message: "Not found." } }, 404);
      return route.fulfill({ status: 200, contentType: "image/png", body: fixtures.PHOTO_BYTES });
    }
    if (/\/content$/.test(url)) {
      const status = opts.photoStatus || 200;
      if (status !== 200) return json(route, { error: { message: "Item not found." } }, status);
      if (request.method() === "PUT") return json(route, { name: "uploaded" });
      return route.fulfill({ status: 200, contentType: "image/png", body: fixtures.PHOTO_BYTES });
    }
    return json(route, { error: { message: "Unexpected Graph call: " + url } }, 500);
  });

  return {
    requests: graphRequests,
    find: function (pattern, method) {
      return graphRequests.filter(function (entry) {
        return pattern.test(entry.url) && (!method || entry.method === method);
      });
    }
  };
}

// Navigates to the app and waits for the item list to be populated.
async function gotoSignedIn(page, options) {
  const graph = await setup(page, options);
  await page.goto("/");
  await expect(page.locator("#toolbar")).toBeVisible();
  await expect(page.locator("#items")).toBeVisible();
  return graph;
}

function cardTitles(page) {
  return page.locator("#items .card h3");
}

// Every text the status line has shown since load, newest last.
function statusTexts(page) {
  return page.evaluate(function () { return window.__statusTexts || []; });
}

// Grouping is on by default and every panel starts collapsed, so the flat mode
// is the simplest ground for the filtering and sorting assertions.
async function selectFlatList(page) {
  await page.locator("#group-field").selectOption("");
}

async function openFilters(page) {
  await page.locator("#filters-btn").click();
  await expect(page.locator("#filters-panel")).toBeVisible();
}

module.exports = {
  cardTitles,
  gotoSignedIn,
  openFilters,
  selectFlatList,
  setup,
  statusTexts
};
