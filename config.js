// Home Inventory — configuration
//
// Replace every "YOUR_..." placeholder below with the real values.
// See README.md for step-by-step instructions on where to find each value.
//
// Do NOT put any secrets here. A client ID is public by design; this app is a
// public client (SPA) and never uses a client secret.

const CONFIG = {
  // Azure App Registration → Overview → Application (client) ID
  clientId: "YOUR_AZURE_APP_CLIENT_ID",

  // Personal Microsoft accounts only
  authority: "https://login.microsoftonline.com/consumers",

  // Must exactly match the SPA redirect URI registered in Azure
  redirectUri: "https://yclkvnc.github.io/home-inventory/",

  // OneDrive item ID of /HomeInventory/inventory.xlsx
  excelFileId: "YOUR_EXCEL_FILE_ONEDRIVE_ITEM_ID",

  // OneDrive item ID of the /HomeInventory/ folder
  folderId: "YOUR_HOMEINVENTORY_FOLDER_ONEDRIVE_ITEM_ID",

  // Name of the Excel table inside inventory.xlsx
  tableName: "Table1",

  scopes: ["Files.ReadWrite", "User.Read"]
};
