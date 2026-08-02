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
                               └── photo-*.jpg      ← item photos
```

The Excel file and the folder are referenced by their fixed OneDrive item IDs, so
everyone who signs in opens exactly the same data.

## Files

| File | Purpose |
|---|---|
| `index.html` | Single-page app shell |
| `styles.css` | All styling (responsive, CSS Grid) |
| `app.js` | Auth, Microsoft Graph calls, rendering |
| `config.js` | Your Client ID, Excel file ID, folder ID |
| `vendor/msal-browser.min.js` | MSAL.js browser library, vendored locally |

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
   `ID`, `Name`, `Category`, `Room`, `Quantity`, `Notes`, `PhotoName`, `CreatedAt`.
4. Select the header row and choose **Insert → Table** (with headers), then rename the
   table to `Table1` (Table Design → Table Name).
5. Get the OneDrive **item IDs** of the folder and of `inventory.xlsx`. The easiest way is
   [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer): sign in and run
   `GET https://graph.microsoft.com/v1.0/me/drive/root:/HomeInventory:/children` —
   the `id` of each entry is what you need. For the folder itself, run
   `GET https://graph.microsoft.com/v1.0/me/drive/root:/HomeInventory`.
6. Share the `HomeInventory` folder with your family members ("Can edit"). One share
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
- **Search** filters live across Name, Category, Room and Notes.
- **Filters** narrow the list by Category or Room.
- Items are grouped by Category; click a photo thumbnail to view it full size.

### Dynamic columns

The app reads the column headers from the Excel table at runtime. Add a new column in
Excel (for example `Brand` or `Warranty Until`) and it automatically appears in the form
and on the cards — no code change needed. `ID`, `PhotoName` and `CreatedAt` are managed
by the app and are not editable in the form.

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
