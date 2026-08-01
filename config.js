// Home Inventory — configuration
//
// Replace every "YOUR_..." placeholder below with the real values.
// See README.md for step-by-step instructions on where to find each value.
//
// Do NOT put any secrets here. A client ID is public by design; this app is a
// public client (SPA) and never uses a client secret.

const CONFIG = {
  // Azure App Registration → Overview → Application (client) ID
  clientId: "6557da9d-9eaa-4b9d-b9f2-7820570f2abc",

  // Personal Microsoft accounts only
  authority: "https://login.microsoftonline.com/consumers",

  // Must exactly match the SPA redirect URI registered in Azure
  redirectUri: "https://yclkvnc.github.io/home-inventory/",

  // OneDrive item ID of /HomeInventory/inventory.xlsx
  excelFileId: "DC4CDEC5003A3E1B!s252c95f238434126a58b3487974183e8",

  // OneDrive item ID of the /HomeInventory/ folder
  folderId: "DC4CDEC5003A3E1B!sca675cafb9994226b6080d938144b6f5",

  // Name of the Excel table inside inventory.xlsx
  tableName: "Table1",

  scopes: ["Files.ReadWrite", "User.Read"]
};
