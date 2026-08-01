/* Home Inventory — vanilla JS app.
   Auth: MSAL.js (browser). Data: Microsoft Graph (Excel table + OneDrive files).
   No framework, no backend, no build step. */

(function () {
  "use strict";

  var GRAPH = "https://graph.microsoft.com/v1.0";

  var PLACEHOLDERS = [
    "YOUR_AZURE_APP_CLIENT_ID",
    "YOUR_EXCEL_FILE_ONEDRIVE_ITEM_ID",
    "YOUR_HOMEINVENTORY_FOLDER_ONEDRIVE_ITEM_ID"
  ];

  // Columns the app manages itself; they are not shown as editable form fields.
  var AUTO_COLUMNS = ["ID", "PhotoName", "CreatedAt"];
  var SEARCH_COLUMNS = ["Name", "Category", "Room", "Notes"];

  var el = {
    status: document.getElementById("status"),
    toolbar: document.getElementById("toolbar"),
    items: document.getElementById("items"),
    empty: document.getElementById("empty"),
    search: document.getElementById("search"),
    filterCategory: document.getElementById("filter-category"),
    filterRoom: document.getElementById("filter-room"),
    addBtn: document.getElementById("add-btn"),
    refreshBtn: document.getElementById("refresh-btn"),
    signIn: document.getElementById("signin-btn"),
    signOut: document.getElementById("signout-btn"),
    userName: document.getElementById("user-name"),
    dialog: document.getElementById("item-dialog"),
    dialogTitle: document.getElementById("dialog-title"),
    form: document.getElementById("item-form"),
    formFields: document.getElementById("form-fields"),
    formError: document.getElementById("form-error"),
    photoInput: document.getElementById("photo-input"),
    photoCurrent: document.getElementById("photo-current"),
    cancelBtn: document.getElementById("cancel-btn"),
    saveBtn: document.getElementById("save-btn"),
    photoDialog: document.getElementById("photo-dialog"),
    photoFull: document.getElementById("photo-full"),
    photoCloseBtn: document.getElementById("photo-close-btn")
  };

  var msalApp = null;
  var account = null;
  var headers = [];   // Excel column names, in sheet order
  var rows = [];      // [{ index: n, values: {Header: value} }]
  var editingId = null;
  var photoUrlCache = {}; // PhotoName -> object URL

  /* ---------------------------------------------------------------- helpers */

  function setStatus(message, isError) {
    el.status.textContent = message || "";
    el.status.hidden = !message;
    el.status.classList.toggle("error", !!isError);
  }

  function show(node, visible) {
    node.hidden = !visible;
  }

  function isConfigured() {
    if (typeof CONFIG === "undefined") return false;
    var values = [CONFIG.clientId, CONFIG.excelFileId, CONFIG.folderId];
    for (var i = 0; i < values.length; i++) {
      if (!values[i] || PLACEHOLDERS.indexOf(values[i]) !== -1) return false;
    }
    return true;
  }

  function newId() {
    return String(Date.now()) + "-" + Math.random().toString(36).slice(2, 8);
  }

  function fileExtension(name) {
    var match = /\.([a-z0-9]{1,5})$/i.exec(name || "");
    return match ? match[1].toLowerCase() : "jpg";
  }

  function errorMessage(error) {
    if (!error) return "Unknown error.";
    return error.message || String(error);
  }

  /* ------------------------------------------------------------------- auth */

  function initAuth() {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: CONFIG.authority,
        redirectUri: CONFIG.redirectUri
      },
      cache: {
        // Tokens are kept in memory only; nothing is written to localStorage.
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    });
    return msalApp.handleRedirectPromise().then(function (result) {
      if (result && result.account) return result.account;
      var accounts = msalApp.getAllAccounts();
      return accounts.length ? accounts[0] : null;
    });
  }

  function getToken() {
    var request = { scopes: CONFIG.scopes, account: account };
    return msalApp.acquireTokenSilent(request).then(function (result) {
      return result.accessToken;
    }, function () {
      return msalApp.acquireTokenPopup({ scopes: CONFIG.scopes }).then(function (result) {
        return result.accessToken;
      });
    });
  }

  function signIn() {
    setStatus("Signing in…");
    msalApp.loginPopup({ scopes: CONFIG.scopes }).then(function (result) {
      account = result.account;
      msalApp.setActiveAccount(account);
      onSignedIn();
    }, function (error) {
      setStatus("Sign-in failed: " + errorMessage(error), true);
    });
  }

  function signOut() {
    revokePhotoUrls();
    msalApp.logoutPopup({ account: account }).catch(function () { /* ignore */ });
    account = null;
    render();
    show(el.toolbar, false);
    show(el.items, false);
    show(el.empty, false);
    show(el.signOut, false);
    show(el.userName, false);
    show(el.signIn, true);
    setStatus("Signed out. Sign in to view the inventory.");
  }

  /* ------------------------------------------------------------ graph layer */

  function graph(path, options) {
    options = options || {};
    return getToken().then(function (token) {
      var init = {
        method: options.method || "GET",
        headers: Object.assign({ Authorization: "Bearer " + token }, options.headers || {})
      };
      if (options.body !== undefined) {
        init.body = options.isRaw ? options.body : JSON.stringify(options.body);
        if (!options.isRaw) init.headers["Content-Type"] = "application/json";
      }
      return fetch(GRAPH + path, init);
    }).then(function (response) {
      if (!response.ok) {
        return response.text().then(function (text) {
          var detail = text;
          try {
            var parsed = JSON.parse(text);
            if (parsed && parsed.error && parsed.error.message) detail = parsed.error.message;
          } catch (e) { /* keep raw text */ }
          throw new Error("Graph " + response.status + ": " + detail);
        });
      }
      if (options.asBlob) return response.blob();
      if (response.status === 204) return null;
      return response.text().then(function (text) {
        return text ? JSON.parse(text) : null;
      });
    });
  }

  function tablePath() {
    return "/me/drive/items/" + encodeURIComponent(CONFIG.excelFileId) +
      "/workbook/tables/" + encodeURIComponent(CONFIG.tableName || "Table1");
  }

  function filePath(name) {
    return "/me/drive/items/" + encodeURIComponent(CONFIG.folderId) +
      ":/" + encodeURIComponent(name) + ":";
  }

  function loadTable() {
    return graph(tablePath() + "/columns?$select=name").then(function (data) {
      headers = (data && data.value ? data.value : []).map(function (column) {
        return column.name;
      });
      return graph(tablePath() + "/rows?$select=index,values");
    }).then(function (data) {
      rows = (data && data.value ? data.value : []).map(function (row) {
        var values = row.values && row.values[0] ? row.values[0] : [];
        var record = {};
        headers.forEach(function (header, i) {
          var value = values[i];
          record[header] = value === null || value === undefined ? "" : String(value);
        });
        return { index: row.index, values: record };
      });
    });
  }

  function rowValuesArray(record) {
    return headers.map(function (header) {
      var value = record[header];
      return value === undefined || value === null ? "" : value;
    });
  }

  function addRow(record) {
    return graph(tablePath() + "/rows/add", {
      method: "POST",
      body: { index: null, values: [rowValuesArray(record)] }
    });
  }

  function updateRow(rowIndex, record) {
    return graph(tablePath() + "/rows/itemAt(index=" + rowIndex + ")", {
      method: "PATCH",
      body: { values: [rowValuesArray(record)] }
    });
  }

  function deleteRow(rowIndex) {
    return graph(tablePath() + "/rows/itemAt(index=" + rowIndex + ")/delete", {
      method: "POST"
    });
  }

  function uploadPhoto(file) {
    var name = "photo-" + newId() + "." + fileExtension(file.name);
    return graph(filePath(name) + "/content", {
      method: "PUT",
      body: file,
      isRaw: true,
      headers: { "Content-Type": file.type || "application/octet-stream" }
    }).then(function () {
      return name;
    });
  }

  function photoUrl(name) {
    if (photoUrlCache[name]) return Promise.resolve(photoUrlCache[name]);
    return graph(filePath(name) + "/content", { asBlob: true }).then(function (blob) {
      var url = URL.createObjectURL(blob);
      photoUrlCache[name] = url;
      return url;
    });
  }

  function revokePhotoUrls() {
    Object.keys(photoUrlCache).forEach(function (name) {
      URL.revokeObjectURL(photoUrlCache[name]);
    });
    photoUrlCache = {};
  }

  /* --------------------------------------------------------------- rendering */

  function activeHeaders() {
    return headers.filter(function (header) {
      return AUTO_COLUMNS.indexOf(header) === -1;
    });
  }

  function matchesSearch(record, term) {
    if (!term) return true;
    var haystack = headers.filter(function (header) {
      return SEARCH_COLUMNS.indexOf(header) !== -1;
    }).map(function (header) {
      return record[header] || "";
    }).join(" ").toLowerCase();
    return haystack.indexOf(term) !== -1;
  }

  function visibleRows() {
    var term = el.search.value.trim().toLowerCase();
    var category = el.filterCategory.value;
    var room = el.filterRoom.value;
    return rows.filter(function (row) {
      var record = row.values;
      if (!matchesSearch(record, term)) return false;
      if (category && (record.Category || "") !== category) return false;
      if (room && (record.Room || "") !== room) return false;
      return true;
    });
  }

  function fillFilter(select, column, allLabel) {
    var previous = select.value;
    var options = [];
    rows.forEach(function (row) {
      var value = row.values[column];
      if (value && options.indexOf(value) === -1) options.push(value);
    });
    options.sort();
    select.textContent = "";
    var all = document.createElement("option");
    all.value = "";
    all.textContent = allLabel;
    select.appendChild(all);
    options.forEach(function (value) {
      var option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });
    select.value = options.indexOf(previous) !== -1 ? previous : "";
    show(select, headers.indexOf(column) !== -1);
  }

  function buildCard(row) {
    var record = row.values;
    var card = document.createElement("article");
    card.className = "card";

    if (record.PhotoName) {
      var img = document.createElement("img");
      img.className = "thumb";
      img.alt = record.Name || "Item photo";
      img.addEventListener("click", function () {
        openPhoto(record.PhotoName, img.alt);
      });
      card.appendChild(img);
      photoUrl(record.PhotoName).then(function (url) {
        img.src = url;
      }, function () {
        img.remove();
      });
    }

    var title = document.createElement("h3");
    title.textContent = record.Name || "(no name)";
    card.appendChild(title);

    var list = document.createElement("dl");
    headers.forEach(function (header) {
      if (header === "Name" || header === "ID" || header === "PhotoName") return;
      var value = record[header];
      if (!value) return;
      var dt = document.createElement("dt");
      dt.textContent = header;
      var dd = document.createElement("dd");
      dd.textContent = value;
      list.appendChild(dt);
      list.appendChild(dd);
    });
    card.appendChild(list);

    var actions = document.createElement("div");
    actions.className = "card-actions";

    var editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "btn";
    editBtn.textContent = "Edit";
    editBtn.addEventListener("click", function () { openForm(record.ID); });

    var deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "btn btn-danger";
    deleteBtn.textContent = "Delete";
    deleteBtn.addEventListener("click", function () { confirmDelete(record); });

    actions.appendChild(editBtn);
    actions.appendChild(deleteBtn);
    card.appendChild(actions);
    return card;
  }

  function render() {
    el.items.textContent = "";
    if (!account) return;

    var visible = visibleRows();
    show(el.empty, visible.length === 0 && rows.length > 0);
    if (rows.length === 0) {
      setStatus("No items yet. Use “Add Item” to create the first one.");
    } else {
      setStatus("");
    }

    var groups = {};
    var order = [];
    visible.forEach(function (row) {
      var category = row.values.Category || "Uncategorized";
      if (!groups[category]) {
        groups[category] = [];
        order.push(category);
      }
      groups[category].push(row);
    });
    order.sort();

    order.forEach(function (category) {
      var heading = document.createElement("h2");
      heading.className = "category-title";
      heading.textContent = category;
      el.items.appendChild(heading);

      var grid = document.createElement("div");
      grid.className = "cards";
      groups[category].forEach(function (row) {
        grid.appendChild(buildCard(row));
      });
      el.items.appendChild(grid);
    });
  }

  /* ------------------------------------------------------------------- form */

  function fieldId(header) {
    return "field-" + header.replace(/[^A-Za-z0-9_-]/g, "_");
  }

  function buildForm(record) {
    el.formFields.textContent = "";
    activeHeaders().forEach(function (header) {
      var wrapper = document.createElement("div");
      wrapper.className = "form-row";

      var label = document.createElement("label");
      label.textContent = header;
      label.htmlFor = fieldId(header);

      var input;
      if (header === "Notes") {
        input = document.createElement("textarea");
        input.rows = 3;
      } else {
        input = document.createElement("input");
        input.type = header === "Quantity" ? "number" : "text";
        if (header === "Quantity") input.min = "0";
      }
      input.id = fieldId(header);
      input.name = header;
      input.value = record ? (record[header] || "") : (header === "Quantity" ? "1" : "");
      if (header === "Name") input.required = true;

      wrapper.appendChild(label);
      wrapper.appendChild(input);
      el.formFields.appendChild(wrapper);
    });
  }

  function findRowById(id) {
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].values.ID === id) return rows[i];
    }
    return null;
  }

  function openForm(id) {
    editingId = id || null;
    var row = id ? findRowById(id) : null;
    el.dialogTitle.textContent = row ? "Edit Item" : "Add Item";
    buildForm(row ? row.values : null);
    el.photoInput.value = "";
    var current = row && row.values.PhotoName;
    el.photoCurrent.textContent = current ? "Current photo: " + current + " (choose a file to replace it)" : "";
    show(el.photoCurrent, !!current);
    show(el.formError, false);
    el.dialog.showModal();
  }

  function readForm() {
    var record = {};
    activeHeaders().forEach(function (header) {
      var input = document.getElementById(fieldId(header));
      record[header] = input ? input.value.trim() : "";
    });
    return record;
  }

  function saveItem(event) {
    event.preventDefault();
    var input = readForm();
    if (headers.indexOf("Name") !== -1 && !input.Name) {
      el.formError.textContent = "Name is required.";
      show(el.formError, true);
      return;
    }

    var existing = editingId ? findRowById(editingId) : null;
    var file = el.photoInput.files && el.photoInput.files[0];

    el.saveBtn.disabled = true;
    show(el.formError, false);

    var uploaded = file ? uploadPhoto(file) : Promise.resolve(null);

    uploaded.then(function (photoName) {
      var record = Object.assign({}, existing ? existing.values : {}, input);
      if (photoName) record.PhotoName = photoName;
      if (!existing) {
        record.ID = newId();
        record.CreatedAt = new Date().toISOString();
      }
      return existing ? updateRow(existing.index, record) : addRow(record);
    }).then(function () {
      el.dialog.close();
      return refresh();
    }).catch(function (error) {
      el.formError.textContent = "Save failed: " + errorMessage(error);
      show(el.formError, true);
    }).then(function () {
      el.saveBtn.disabled = false;
    });
  }

  function confirmDelete(record) {
    if (!window.confirm('Delete "' + (record.Name || "this item") + '"? This cannot be undone.')) return;
    var row = findRowById(record.ID);
    if (!row) return;
    setStatus("Deleting…");
    deleteRow(row.index).then(function () {
      return refresh();
    }).catch(function (error) {
      setStatus("Delete failed: " + errorMessage(error), true);
    });
  }

  function openPhoto(name, alt) {
    photoUrl(name).then(function (url) {
      el.photoFull.src = url;
      el.photoFull.alt = alt || "Item photo";
      el.photoDialog.showModal();
    }, function (error) {
      setStatus("Could not load photo: " + errorMessage(error), true);
    });
  }

  /* ------------------------------------------------------------------- flow */

  function refresh() {
    setStatus("Loading items…");
    return loadTable().then(function () {
      fillFilter(el.filterCategory, "Category", "All categories");
      fillFilter(el.filterRoom, "Room", "All rooms");
      render();
    }, function (error) {
      setStatus("Could not load the inventory: " + errorMessage(error), true);
    });
  }

  function onSignedIn() {
    el.userName.textContent = account.name || account.username || "";
    show(el.userName, true);
    show(el.signIn, false);
    show(el.signOut, true);
    show(el.toolbar, true);
    show(el.items, true);
    refresh();
  }

  function wireEvents() {
    el.signIn.addEventListener("click", signIn);
    el.signOut.addEventListener("click", signOut);
    el.addBtn.addEventListener("click", function () { openForm(null); });
    el.refreshBtn.addEventListener("click", function () { refresh(); });
    el.search.addEventListener("input", render);
    el.filterCategory.addEventListener("change", render);
    el.filterRoom.addEventListener("change", render);
    el.form.addEventListener("submit", saveItem);
    el.cancelBtn.addEventListener("click", function () { el.dialog.close(); });
    el.photoCloseBtn.addEventListener("click", function () { el.photoDialog.close(); });
  }

  function start() {
    if (!isConfigured()) {
      setStatus("Not configured yet — open config.js and fill in the Client ID, Excel file ID and folder ID. See README for instructions.", true);
      return;
    }
    wireEvents();
    initAuth().then(function (existingAccount) {
      if (existingAccount) {
        account = existingAccount;
        msalApp.setActiveAccount(account);
        onSignedIn();
      } else {
        show(el.signIn, true);
        setStatus("Sign in with your Microsoft account to view the inventory.");
      }
    }, function (error) {
      setStatus("Authentication could not start: " + errorMessage(error), true);
    });
  }

  start();
})();
