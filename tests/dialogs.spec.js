const { expect, test } = require("@playwright/test");
const fixtures = require("./fixtures");
const { gotoSignedIn, setup } = require("./helpers");

const EDITABLE_FIELDS = ["Name", "Category", "Room", "Quantity", "Notes", "Tags", "Purchased"];

const UPLOAD = {
  name: "new-photo.png",
  mimeType: "image/png",
  buffer: fixtures.PHOTO_BYTES
};

function fieldLabels(page) {
  return page.locator("#form-fields .form-row > label");
}

test("the add dialog is built from the columns, without the automatic ones", async function ({ page }) {
  await gotoSignedIn(page, {});
  await page.locator("#add-btn").click();
  await expect(page.locator("#item-dialog")).toBeVisible();
  await expect(page.locator("#dialog-title")).toHaveText("Add Item");
  await expect(fieldLabels(page)).toHaveText(EDITABLE_FIELDS);
  await expect(page.locator("#field-Quantity")).toHaveValue("1");
  await expect(page.locator("#field-Name")).toHaveValue("");
});

test("the edit dialog is prefilled from the selected row", async function ({ page }) {
  await gotoSignedIn(page, {});
  await page.locator("#group-field").selectOption("");
  await page.locator("#items .card").filter({ hasText: "Cordless drill" })
    .getByRole("button", { name: "Edit" }).click();

  await expect(page.locator("#dialog-title")).toHaveText("Edit Item");
  await expect(page.locator("#field-Name")).toHaveValue("Cordless drill");
  await expect(page.locator("#field-Category")).toHaveValue("Tools");
  await expect(page.locator("#field-Room")).toHaveValue("Office");
  await expect(page.locator("#field-Notes")).toHaveValue("18V battery");
  await expect(page.locator("#form-fields .tag > span")).toHaveText(["metal"]);
});

test("adding an item posts rows/add with the values in column order", async function ({ page }) {
  const graph = await gotoSignedIn(page, {});
  await page.locator("#add-btn").click();
  await page.locator("#field-Name").fill("Watering can");
  await page.locator("#field-Category").fill("Garden");
  await page.locator("#field-Room").fill("Balcony");
  await page.locator("#field-Quantity").fill("3");
  await page.locator("#save-btn").click();

  await expect(page.locator("#item-dialog")).toBeHidden();
  await expect.poll(function () {
    return graph.find(/\/rows\/add$/, "POST").length;
  }).toBe(1);

  const values = graph.find(/\/rows\/add$/, "POST")[0].body.values[0];
  expect(values.length).toBe(fixtures.COLUMNS.length);
  expect(values[fixtures.COLUMNS.indexOf("Name")]).toBe("Watering can");
  expect(values[fixtures.COLUMNS.indexOf("Category")]).toBe("Garden");
  expect(values[fixtures.COLUMNS.indexOf("Room")]).toBe("Balcony");
  expect(values[fixtures.COLUMNS.indexOf("Quantity")]).toBe("3");
  expect(values[fixtures.COLUMNS.indexOf("ID")]).not.toBe("");
  expect(values[fixtures.COLUMNS.indexOf("CreatedAt")]).not.toBe("");
  expect(values[fixtures.COLUMNS.indexOf("UpdatedAt")]).not.toBe("");
});

test("editing an item patches rows/itemAt at the row index", async function ({ page }) {
  const graph = await gotoSignedIn(page, {});
  await page.locator("#group-field").selectOption("");
  await page.locator("#items .card").filter({ hasText: "Armchair" })
    .getByRole("button", { name: "Edit" }).click();
  await page.locator("#field-Name").fill("Armchair (repaired)");
  await page.locator("#save-btn").click();

  const index = fixtures.ROWS.findIndex(function (row) { return row.Name === "Armchair"; });
  await expect.poll(function () {
    return graph.find(new RegExp("/rows/itemAt\\(index=" + index + "\\)$"), "PATCH").length;
  }).toBe(1);

  const patch = graph.find(new RegExp("/rows/itemAt\\(index=" + index + "\\)$"), "PATCH")[0];
  const values = patch.body.values[0];
  expect(values[fixtures.COLUMNS.indexOf("Name")]).toBe("Armchair (repaired)");
  expect(values[fixtures.COLUMNS.indexOf("ID")]).toBe("id-chair");
});

test("a failed save shows the error with a working retry and dismiss", async function ({ page }) {
  const graph = await gotoSignedIn(page, { writeStatus: 500 });
  await page.locator("#add-btn").click();
  await page.locator("#field-Name").fill("Doomed item");
  await page.locator("#save-btn").click();

  const error = page.locator("#form-error");
  await expect(error).toBeVisible();
  await expect(page.locator("#form-error-text")).toContainText("Save failed:");
  await expect(page.locator("#retry-btn")).toBeVisible();

  await page.locator("#retry-btn").click();
  await expect(error).toBeVisible();
  await expect.poll(function () {
    return graph.find(/\/rows\/add$/, "POST").length;
  }).toBe(2);

  await page.locator("#dismiss-error-btn").click();
  await expect(error).toBeHidden();
  await expect(page.locator("#item-dialog")).toBeVisible();
});

test("a missing name is refused before any Graph call", async function ({ page }) {
  const graph = await gotoSignedIn(page, {});
  await page.locator("#add-btn").click();
  // The form is submitted through the retry path so the browser's own required
  // validation does not swallow the click.
  await page.locator("#field-Name").fill("");
  await page.evaluate(function () {
    document.getElementById("item-form").dispatchEvent(new Event("submit", { cancelable: true }));
  });
  await expect(page.locator("#form-error-text")).toHaveText("Name is required.");
  expect(graph.find(/\/rows\/add$/, "POST").length).toBe(0);
});

test("a photo upload rejected with 404 explains the Photos subfolder", async function ({ page }) {
  await gotoSignedIn(page, { photoStatus: 404 });
  await page.locator("#add-btn").click();
  await page.locator("#field-Name").fill("Photographed item");
  await page.locator("#photo-input").setInputFiles(UPLOAD);
  await page.locator("#save-btn").click();

  await expect(page.locator("#form-error-text")).toContainText('"Photos" subfolder');
});

test("removing the photo hides the preview", async function ({ page }) {
  await gotoSignedIn(page, { rows: [fixtures.ROW_WITH_PHOTO] });
  await page.locator("#group-field").selectOption("");
  await page.locator("#items .card").getByRole("button", { name: "Edit" }).click();

  await expect(page.locator("#photo-preview")).toBeVisible();
  await page.locator("#remove-photo-btn").click();
  await expect(page.locator("#photo-preview")).toBeHidden();
  await expect(page.locator("#photo-current")).toBeHidden();
});

test("the photo dialog opens from a card thumbnail and closes", async function ({ page }) {
  await gotoSignedIn(page, { rows: [fixtures.ROW_WITH_PHOTO] });
  await page.locator("#group-field").selectOption("");
  await page.locator("#items .card .thumb").click();

  await expect(page.locator("#photo-dialog")).toBeVisible();
  await expect(page.locator("#photo-full")).toHaveAttribute("alt", fixtures.ROW_WITH_PHOTO.Name);
  await page.locator("#photo-close-btn").click();
  await expect(page.locator("#photo-dialog")).toBeHidden();
});

test("cancel closes the dialog without touching Graph", async function ({ page }) {
  const graph = await gotoSignedIn(page, {});
  await page.locator("#add-btn").click();
  await page.locator("#cancel-btn").click();
  await expect(page.locator("#item-dialog")).toBeHidden();
  expect(graph.find(/\/rows\//, "POST").length).toBe(0);
});

test("the sign-in button triggers no Graph call while signed out", async function ({ page }) {
  const graph = await setup(page, { signedIn: false });
  await page.goto("/");
  await expect(page.locator("#signin-btn")).toBeVisible();
  expect(graph.requests.length).toBe(0);
});
