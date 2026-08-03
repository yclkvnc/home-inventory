# Home Inventory

A simple, secure web app to keep an inventory of household items. The whole family
can view and edit the same list, and each item can have a photo.

- **No framework** — plain HTML + CSS + vanilla JavaScript.
- **No backend** — a fully static site, hosted for free on GitHub Pages.
- **Real sign-in** — Microsoft OAuth via MSAL.js, no shared passwords.
- **Single source of truth** — one Excel file in a single OneDrive folder.
- **Cost** — $0.

## How it works

```
Browser ──sign in (MSAL.js)──▶ Microsoft Identity Platform
   │
   └──Microsoft Graph API──▶ OneDrive /HomeInventory/
                               ├── inventory.xlsx   ← the database
                               └── Photos/
                                   └── photo-*.jpg  ← item photos
```

The Excel file and the folder are referenced by their fixed OneDrive item IDs, so
everyone who signs in opens exactly the same data.

## Files

| File | Purpose |
|---|---|
| `index.html` | Single-page app shell |
| `styles.css` | All styling (responsive, CSS Grid) |
| `app.js` | Auth, Microsoft Graph calls, rendering |
| `config.js` | Your Client ID, Excel file ID, folder ID, photo subfolder name |
| `vendor/msal-browser.min.js` | MSAL.js browser library, vendored locally |
| `tests/` | Playwright end-to-end tests (development only) |
| `playwright.config.js` | Test runner configuration |

### Vendored MSAL

MSAL.js is not loaded from a CDN — it is committed to the repo at
`vendor/msal-browser.min.js` and served from the same origin as the site. This
removes an external point of failure and keeps the app working behind ad blockers
or firewalls that block auth-related domains.

Pinned version: **`@azure/msal-browser` 2.38.3** (v2 API — `app.js` depends on it).

To update it, download `lib/msal-browser.min.js` from the npm package and replace
the file, for example:

```sh
npm pack @azure/msal-browser@<version>
tar xzf azure-msal-browser-<version>.tgz package/lib/msal-browser.min.js
cp package/lib/msal-browser.min.js vendor/msal-browser.min.js
```

Do not commit the `.map` file. Staying on v2 is intentional: v3+ contains breaking
changes that would require rewriting the auth layer in `app.js`.

## Tests

End-to-end tests run with [Playwright](https://playwright.dev/). They are a
development-only dependency: the app itself still ships as plain HTML, CSS and
JavaScript with no build step, and nothing under `node_modules/` is deployed.

```sh
npm install
npx playwright install --with-deps chromium
npm test
```

The runner serves the repository with `python3 -m http.server` and stubs
everything the app talks to — `config.js`, MSAL and Microsoft Graph are all
answered from `tests/fixtures.js`, so no real account or spreadsheet is needed.
The same suite runs on every pull request through `.github/workflows/tests.yml`.

## Setup

These steps are one-time and must be done manually by the repo owner.

### 1. Register the app in Azure (free, no subscription required)

1. Go to <https://portal.azure.com> and sign in with your personal Microsoft account.
2. **App registrations → New registration**.
3. Name: `Home Inventory`.
4. Supported account types: **Personal Microsoft accounts only**.
5. Redirect URI: platform **Single-page application (SPA)**, value
   `https://<your-github-user>.github.io/home-inventory/`.
6. After registering, copy the **Application (client) ID** from the Overview page —
   this is `clientId` in `config.js`.
7. **API permissions → Add a permission → Microsoft Graph → Delegated**, add:
   - `Files.ReadWrite`
   - `User.Read`
8. To limit who can sign in: **Enterprise applications → Home Inventory → Properties**,
   set **User assignment required** to **Yes**, then add the allowed users under
   **Users and groups**. Everyone else is blocked at the Microsoft sign-in screen.

### 2. Prepare OneDrive

1. Create a folder named `HomeInventory` in your OneDrive.
2. Inside it, create an Excel workbook named `inventory.xlsx`.
3. Put these headers in the first row:
   `ID`, `Name`, `Category`, `Room`, `Quantity`, `Notes`, `Tags`, `PhotoName`, `CreatedAt`,
   `UpdatedAt`, `Status`.
4. Select the header row and choose **Insert → Table** (with headers), then rename the
   table to `Table1` (Table Design → Table Name).
5. Inside `HomeInventory`, create a subfolder named `Photos` — all item photos are
   stored there, so the database and the images stay separate. The app never creates
   this folder; if it is missing, uploading or opening a photo fails with a message
   telling you to create it.
6. Get the OneDrive **item IDs** of the folder and of `inventory.xlsx`. The easiest way is
   [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer): sign in and run
   `GET https://graph.microsoft.com/v1.0/me/drive/root:/HomeInventory:/children` —
   the `id` of each entry is what you need. For the folder itself, run
   `GET https://graph.microsoft.com/v1.0/me/drive/root:/HomeInventory`.
7. Share the `HomeInventory` folder with your family members ("Can edit"). One share
   covers both the database and all photos.

### 3. Configure the app

Edit `config.js` and replace the placeholders:

| Setting | Where to get it |
|---|---|
| `clientId` | Azure App registration → Overview → Application (client) ID |
| `redirectUri` | Your GitHub Pages URL (must match the SPA redirect URI in Azure) |
| `excelFileId` | OneDrive item ID of `inventory.xlsx` |
| `folderId` | OneDrive item ID of the `HomeInventory` folder |
| `tableName` | The Excel table name (default `Table1`) |
| `photoFolder` | Name of the photo subfolder inside `HomeInventory` (default `Photos`) |

Until these are filled in, the app shows a "Not configured yet" message instead of
failing with a raw API error.

### 4. Publish on GitHub Pages

Repository **Settings → Pages → Build and deployment**, source **Deploy from a branch**,
branch `main`, folder `/ (root)`. The site is then live at the redirect URI you registered.

## Using the app

- **Sign in** with your Microsoft account.
- **Add Item** opens a form generated from the Excel column headers.
- **Category** and **Room** suggest values already used in the inventory, while still
  accepting any new value you type.
- **Edit** / **Delete** are available on every card (delete asks for confirmation).
- **Tags** let you label an item freely: pick a tag already used in the inventory from
  the suggestion list or type a new one and press Enter. Each tag becomes a chip that
  can be removed individually.
- **Search** filters live across Name, Category, Room, Notes and Tags.
- **Filters** opens a panel holding the Category, Room and tag filters. The button shows a
  badge with the number of active filters and **Clear all filters** resets them at once.
- **Tag filter** in the filters panel lists every tag in use; selecting several shows only
  the items carrying *all* of them, and **Clear tags** resets the tag selection.
- **Sort** drives both levels at once: pick `Name`, `CreatedAt` or `UpdatedAt` and toggle
  the direction with the ↓/↑ button. Items are sorted by the chosen field, groups by
  the newest value among their items (alphabetically when `Name` is chosen). Rows with an empty
  `UpdatedAt` fall back to their `CreatedAt`, and rows with neither sort last. The
  selection defaults to `UpdatedAt` descending and is remembered in `localStorage`.
- **Group by** picks how the list is arranged: `No group` shows one flat grid of cards,
  `By category` (the default) and `By room` put the items into collapsible panels — all
  collapsed on load; use **Expand all** / **Collapse all** above the list to open or close
  every panel at once. Items with an empty value land under `Uncategorized` or `No room`.
  Each mode remembers which of its panels are open, and the selection is remembered in
  `localStorage` under `homeInventory.groupBy`. Click a photo thumbnail to view it full size.
- **Theme**: the ☾/☀ button in the top bar switches between light and dark. Without an
  explicit choice the app follows the operating system setting (`prefers-color-scheme`);
  the choice is remembered in `localStorage` under `homeInventory.theme`.
- **Account menu**: when signed in the top bar shows only the theme toggle and a round
  avatar with your initials. Clicking it opens a panel with your profile photo (from
  Microsoft Graph, falling back to the initials when you have none), your full name, your
  email address and the **Sign out** button. The panel closes on an outside click or Escape.
- After saving an item, its group panel opens and the item is scrolled into view and
  briefly highlighted; the sort selection is left untouched.

### Dynamic columns

The app reads the column headers from the Excel table at runtime. Add a new column in
Excel (for example `Brand` or `Warranty Until`) and it automatically appears in the form
and on the cards — no code change needed. `ID`, `PhotoName`, `CreatedAt`, `UpdatedAt` and
`Status` are managed by the app and are not editable in the form. `CreatedAt` and
`UpdatedAt` are still shown read-only on the cards; `UpdatedAt` is written on every save,
edit and delete. All dates are displayed in your local timezone, in the fixed format
`dd.MM.yyyy HH:mm:ss` (`dd.MM.yyyy` for date-only values).

### Tags column

The optional `Tags` column stores all tags of an item in a single cell, separated by
semicolons (for example `electronics;living room;fragile`). Whitespace around a tag is
trimmed when reading and writing, and empty entries are ignored. If the column is missing,
the tag editor and the tag filter simply do not appear.

### Status column

The `Status` column controls whether a row is visible in the app:

| Value | Meaning |
|---|---|
| *(empty)* | Active — the item appears in the card list, filters and search results |
| `deleted` | Soft-deleted — the item is hidden everywhere in the app but the row stays in the spreadsheet |

Clicking **Delete** on a card marks the row as `deleted` instead of removing it, so the
data is always recoverable. Rows can be cleaned up (permanently removed) manually in Excel
whenever desired. The associated photo file in the `Photos` subfolder is never removed
by the app.

Excel stores only the photo's file name in `PhotoName`; the app always resolves it
against the `Photos` subfolder of the `HomeInventory` folder.

## Security notes

- Access tokens are only sent in the HTTP `Authorization` header, never in a URL.
- MSAL is configured to use `sessionStorage`; the app itself never writes tokens to
  `localStorage` and no credentials are stored in this repository.
- Photos are fetched through Graph with the auth header and displayed via an object URL,
  so no signed public links are created.
- `clientId`, `excelFileId` and `folderId` are not secrets, but they are user-specific —
  keep your own values in your fork.
- Photos and inventory data are never committed to the repository; they live only in OneDrive.

## Not included (possible future work)

Export to CSV, barcode/QR scanning, item history, offline mode.
