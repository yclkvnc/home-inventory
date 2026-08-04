// Fixture data for the Playwright smoke suite.
//
// The app reads the column names at runtime, so every fixture is expressed as a
// column list plus rows given as objects keyed by column name. `toGraphRows`
// turns them into the shape Microsoft Graph returns for an Excel table.

const CONFIG = {
  clientId: "test-client-id",
  authority: "https://login.microsoftonline.com/consumers",
  redirectUri: "http://localhost:4173/",
  excelFileId: "test-excel-file-id",
  folderId: "test-folder-id",
  tableName: "Table1",
  photoFolder: "Photos",
  scopes: ["Files.ReadWrite", "User.Read"]
};

const PLACEHOLDER_CONFIG = Object.assign({}, CONFIG, {
  clientId: "YOUR_AZURE_APP_CLIENT_ID",
  excelFileId: "YOUR_EXCEL_FILE_ONEDRIVE_ITEM_ID",
  folderId: "YOUR_HOMEINVENTORY_FOLDER_ONEDRIVE_ITEM_ID"
});

const ACCOUNT = {
  homeAccountId: "test-home-account-id",
  localAccountId: "test-local-account-id",
  username: "ada@example.com",
  name: "Ada Lovelace"
};

const COLUMNS = [
  "ID",
  "Name",
  "Category",
  "Room",
  "Quantity",
  "Notes",
  "Tags",
  "Purchased",
  "PhotoName",
  "Status",
  "CreatedAt",
  "UpdatedAt"
];

const ROWS = [
  {
    ID: "id-lamp",
    Name: "Desk lamp",
    Category: "Lighting",
    Room: "Office",
    Quantity: "1",
    Notes: "Adjustable arm",
    Tags: "metal fragile",
    Purchased: "2024-03-05",
    PhotoName: "",
    Status: "",
    CreatedAt: "2024-01-01T10:00:00.000Z",
    UpdatedAt: "2024-01-05T10:00:00.000Z"
  },
  {
    ID: "id-chair",
    Name: "Armchair",
    Category: "Furniture",
    Room: "Living room",
    Quantity: "2",
    Notes: "Green fabric",
    Tags: "fragile",
    Purchased: "2023-11-20",
    PhotoName: "",
    Status: "",
    CreatedAt: "2024-02-01T10:00:00.000Z",
    UpdatedAt: "2024-02-02T10:00:00.000Z"
  },
  {
    ID: "id-drill",
    Name: "Cordless drill",
    Category: "Tools",
    Room: "Office",
    Quantity: "42",
    Notes: "18V battery",
    Tags: "metal",
    Purchased: "2022-06-30",
    PhotoName: "",
    Status: "",
    CreatedAt: "2024-03-01T10:00:00.000Z",
    UpdatedAt: "2024-03-03T10:00:00.000Z"
  },
  {
    ID: "id-ghost",
    Name: "Broken kettle",
    Category: "Kitchen",
    Room: "Kitchen",
    Quantity: "1",
    Notes: "Thrown away",
    Tags: "metal",
    Purchased: "2021-01-01",
    PhotoName: "",
    Status: "deleted",
    CreatedAt: "2024-04-01T10:00:00.000Z",
    UpdatedAt: "2024-04-02T10:00:00.000Z"
  }
];

// Names of the rows the app is expected to show for the default fixture.
const VISIBLE_NAMES = ["Desk lamp", "Armchair", "Cordless drill"];

const ROW_WITH_PHOTO = {
  ID: "id-photo",
  Name: "Picture frame",
  Category: "Decoration",
  Room: "Hallway",
  Quantity: "1",
  Notes: "",
  Tags: "",
  Purchased: "",
  PhotoName: "photo-frame.jpg",
  Status: "",
  CreatedAt: "2024-05-01T10:00:00.000Z",
  UpdatedAt: "2024-05-01T10:00:00.000Z"
};

// Smallest possible PNG, served as the body of every photo download.
const PHOTO_BYTES = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=",
  "base64"
);

function toGraphColumns(columns) {
  return { value: columns.map(function (name) { return { name: name }; }) };
}

function toGraphRows(columns, rows) {
  return {
    value: rows.map(function (row, index) {
      return {
        index: index,
        values: [columns.map(function (name) {
          return row[name] === undefined ? "" : row[name];
        })]
      };
    })
  };
}

module.exports = {
  ACCOUNT,
  COLUMNS,
  CONFIG,
  PHOTO_BYTES,
  PLACEHOLDER_CONFIG,
  ROWS,
  ROW_WITH_PHOTO,
  VISIBLE_NAMES,
  toGraphColumns,
  toGraphRows
};
