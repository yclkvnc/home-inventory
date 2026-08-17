/* Home Inventory — orphan photo management.
   Auth and data access use Microsoft Graph directly from the browser. */

(function () {
  "use strict";

  var errorMessage = window.HomeInventoryLib.errorMessage;

  var GRAPH = "https://graph.microsoft.com/v1.0";
  var DEFAULT_PHOTO_FOLDER = "Photos";
  var THEME_STORAGE_KEY = "homeInventory.theme";
  var PLACEHOLDERS = [
    "YOUR_AZURE_APP_CLIENT_ID",
    "YOUR_EXCEL_FILE_ONEDRIVE_ITEM_ID",
    "YOUR_HOMEINVENTORY_FOLDER_ONEDRIVE_ITEM_ID"
  ];

  var el = {
    status: document.getElementById("status"),
    toolbar: document.getElementById("orphan-toolbar"),
    photos: document.getElementById("orphan-photos"),
    empty: document.getElementById("empty"),
    selectAll: document.getElementById("select-all"),
    selectionCount: document.getElementById("selection-count"),
    deleteSelected: document.getElementById("delete-selected-btn"),
    refresh: document.getElementById("refresh-btn"),
    signIn: document.getElementById("signin-btn"),
    signOut: document.getElementById("signout-btn"),
    themeToggleBtn: document.getElementById("theme-toggle-btn"),
    themeToggleIcon: document.getElementById("theme-toggle-icon")
  };

  var msalApp = null;
  var account = null;
  var photos = [];
  var selected = {};
  var photoUrls = {};
  var theme = null;
  var busy = false;

  function show(node, visible) {
    node.hidden = !visible;
  }

  function setStatus(message, isError) {
    el.status.textContent = message || "";
    el.status.hidden = !message;
    el.status.classList.toggle("error", !!isError);
  }

  function isConfigured() {
    if (typeof CONFIG === "undefined") return false;
    var values = [CONFIG.clientId, CONFIG.excelFileId, CONFIG.folderId];
    for (var i = 0; i < values.length; i++) {
      if (!values[i] || PLACEHOLDERS.indexOf(values[i]) !== -1) return false;
    }
    return true;
  }

  function prefersDark() {
    return !!(window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches);
  }

  function activeTheme() {
    return theme || (prefersDark() ? "dark" : "light");
  }

  function applyTheme() {
    var current = activeTheme();
    var label = current === "dark" ? "Switch to light theme" : "Switch to dark theme";
    document.documentElement.setAttribute("data-theme", current);
    el.themeToggleBtn.setAttribute("aria-pressed", current === "dark" ? "true" : "false");
    el.themeToggleBtn.setAttribute("aria-label", label);
    el.themeToggleBtn.title = label;
    el.themeToggleIcon.textContent = current === "dark" ? "☀" : "☾";
  }

  function loadTheme() {
    try {
      var stored = window.localStorage.getItem(THEME_STORAGE_KEY);
      if (stored === "light" || stored === "dark") theme = stored;
    } catch (e) { /* no or unusable storage — follow the OS preference */ }
    applyTheme();
    if (window.matchMedia) {
      var query = window.matchMedia("(prefers-color-scheme: dark)");
      var onChange = function () { if (!theme) applyTheme(); };
      if (query.addEventListener) query.addEventListener("change", onChange);
      else if (query.addListener) query.addListener(onChange);
    }
  }

  function toggleTheme() {
    theme = activeTheme() === "dark" ? "light" : "dark";
    try {
      window.localStorage.setItem(THEME_STORAGE_KEY, theme);
    } catch (e) { /* storage unavailable — the choice stays for this session */ }
    applyTheme();
  }

  function initAuth() {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: CONFIG.authority,
        redirectUri: CONFIG.redirectUri,
        navigateToLoginRequestUrl: true
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
      }
    });
    var init = msalApp.initialize ? msalApp.initialize() : Promise.resolve();
    return init.then(function () {
      return msalApp.handleRedirectPromise();
    }).then(function (result) {
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
      msalApp.acquireTokenRedirect({ scopes: CONFIG.scopes, account: account });
      return new Promise(function () { /* navigation in progress */ });
    });
  }

  function graph(path, options) {
    options = options || {};
    return getToken().then(function (token) {
      var init = {
        method: options.method || "GET",
        headers: { Authorization: "Bearer " + token }
      };
      return fetch(path.indexOf("https://") === 0 ? path : GRAPH + path, init);
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

  function photoFolderName() {
    return CONFIG.photoFolder || DEFAULT_PHOTO_FOLDER;
  }

  function allPages(path) {
    var values = [];
    function next(url) {
      return graph(url).then(function (data) {
        values = values.concat(data && data.value ? data.value : []);
        return data && data["@odata.nextLink"] ? next(data["@odata.nextLink"]) : values;
      });
    }
    return next(path);
  }

  function loadReferences() {
    return graph(tablePath() + "/columns?$select=name").then(function (data) {
      var headers = (data && data.value ? data.value : []).map(function (column) {
        return column.name;
      });
      var photoIndex = headers.indexOf("PhotoName");
      if (photoIndex === -1) {
        throw new Error('The spreadsheet is missing a "PhotoName" column.');
      }
      return allPages(tablePath() + "/rows?$select=values").then(function (rows) {
        var references = {};
        rows.forEach(function (row) {
          var values = row.values && row.values[0] ? row.values[0] : [];
          var name = String(values[photoIndex] === undefined || values[photoIndex] === null
            ? "" : values[photoIndex]).trim().toLowerCase();
          if (name) references[name] = true;
        });
        return references;
      });
    });
  }

  function loadPhotos() {
    var path = "/me/drive/items/" + encodeURIComponent(CONFIG.folderId) +
      ":/" + encodeURIComponent(photoFolderName()) +
      ":/children?$select=id,name,file&$top=200";
    return allPages(path);
  }

  function revokePhotoUrl(id) {
    if (!photoUrls[id]) return;
    URL.revokeObjectURL(photoUrls[id]);
    delete photoUrls[id];
  }

  function revokePhotoUrls() {
    Object.keys(photoUrls).forEach(revokePhotoUrl);
  }

  function syncSelection() {
    var count = Object.keys(selected).filter(function (id) {
      return selected[id];
    }).length;
    var allSelected = photos.length > 0 && count === photos.length;
    el.selectAll.checked = allSelected;
    el.selectAll.indeterminate = count > 0 && !allSelected;
    el.selectionCount.textContent = count + " selected";
    el.deleteSelected.disabled = busy || count === 0;
  }

  function loadPhoto(photo, img) {
    if (photoUrls[photo.id]) {
      img.src = photoUrls[photo.id];
      return;
    }
    graph("/me/drive/items/" + encodeURIComponent(photo.id) + "/content", { asBlob: true })
      .then(function (blob) {
        if (!document.contains(img)) return;
        var url = URL.createObjectURL(blob);
        photoUrls[photo.id] = url;
        img.src = url;
      }, function () {
        if (document.contains(img)) {
          img.replaceWith(document.createTextNode("Preview unavailable"));
        }
      });
  }

  function buildCard(photo) {
    var card = document.createElement("article");
    card.className = "orphan-card";

    var label = document.createElement("label");
    label.className = "orphan-select";
    var checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = !!selected[photo.id];
    checkbox.setAttribute("aria-label", "Select " + photo.name);
    checkbox.addEventListener("change", function () {
      selected[photo.id] = checkbox.checked;
      card.classList.toggle("orphan-card-selected", checkbox.checked);
      syncSelection();
    });
    label.appendChild(checkbox);

    var media = document.createElement("div");
    media.className = "orphan-media";
    var img = document.createElement("img");
    img.alt = "";
    media.appendChild(img);
    loadPhoto(photo, img);

    var name = document.createElement("p");
    name.className = "orphan-name";
    name.textContent = photo.name;
    name.title = photo.name;

    label.appendChild(media);
    label.appendChild(name);
    card.appendChild(label);
    card.classList.toggle("orphan-card-selected", checkbox.checked);
    return card;
  }

  function render() {
    el.photos.textContent = "";
    photos.forEach(function (photo) {
      el.photos.appendChild(buildCard(photo));
    });
    show(el.photos, photos.length > 0);
    show(el.empty, photos.length === 0);
    syncSelection();
  }

  function refresh() {
    if (busy) return Promise.resolve();
    busy = true;
    selected = {};
    el.refresh.disabled = true;
    el.deleteSelected.disabled = true;
    setStatus("Scanning the inventory and Photos folder…");
    revokePhotoUrls();
    return Promise.all([loadReferences(), loadPhotos()]).then(function (results) {
      var references = results[0];
      photos = results[1].filter(function (item) {
        return item.file && !references[String(item.name || "").trim().toLowerCase()];
      }).sort(function (a, b) {
        return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
      });
      render();
      setStatus(photos.length + (photos.length === 1 ? " orphan photo found." : " orphan photos found."));
    }).catch(function (error) {
      photos = [];
      render();
      setStatus("Could not scan for orphan photos: " + errorMessage(error), true);
    }).then(function () {
      busy = false;
      el.refresh.disabled = false;
      syncSelection();
    });
  }

  function deleteSelected() {
    if (busy) return;
    var chosen = photos.filter(function (photo) {
      return selected[photo.id];
    });
    if (chosen.length === 0) return;
    if (!window.confirm(
      "Permanently delete " + chosen.length + (chosen.length === 1 ? " photo" : " photos") +
      " from OneDrive? This cannot be undone."
    )) return;

    busy = true;
    el.refresh.disabled = true;
    syncSelection();
    setStatus("Deleting selected photos…");
    var deletions = chosen.map(function (photo) {
      return graph("/me/drive/items/" + encodeURIComponent(photo.id), { method: "DELETE" })
        .then(function () {
          return { photo: photo, deleted: true };
        }, function (error) {
          return { photo: photo, deleted: false, error: error };
        });
    });
    Promise.all(deletions).then(function (results) {
      var deleted = {};
      var failed = 0;
      results.forEach(function (result) {
        if (result.deleted) {
          deleted[result.photo.id] = true;
          revokePhotoUrl(result.photo.id);
        } else {
          failed++;
        }
      });
      photos = photos.filter(function (photo) {
        return !deleted[photo.id];
      });
      selected = {};
      busy = false;
      el.refresh.disabled = false;
      render();
      if (failed) {
        setStatus(
          (chosen.length - failed) + " deleted; " + failed +
          (failed === 1 ? " photo could" : " photos could") + " not be deleted.",
          true
        );
      } else {
        setStatus(chosen.length + (chosen.length === 1 ? " photo deleted." : " photos deleted."));
      }
    });
  }

  function signIn() {
    setStatus("Signing in…");
    msalApp.loginRedirect({ scopes: CONFIG.scopes }).catch(function (error) {
      setStatus("Sign-in failed: " + errorMessage(error), true);
    });
  }

  function signOut() {
    revokePhotoUrls();
    account = null;
    show(el.toolbar, false);
    show(el.photos, false);
    show(el.empty, false);
    show(el.signOut, false);
    show(el.signIn, true);
    setStatus("Signing out…");
    msalApp.logoutRedirect({ postLogoutRedirectUri: CONFIG.redirectUri })
      .catch(function () { /* ignore */ });
  }

  function wireEvents() {
    el.themeToggleBtn.addEventListener("click", toggleTheme);
    el.signIn.addEventListener("click", signIn);
    el.signOut.addEventListener("click", signOut);
    el.refresh.addEventListener("click", refresh);
    el.deleteSelected.addEventListener("click", deleteSelected);
    el.selectAll.addEventListener("change", function () {
      photos.forEach(function (photo) {
        selected[photo.id] = el.selectAll.checked;
      });
      render();
    });
    window.addEventListener("beforeunload", revokePhotoUrls);
  }

  function start() {
    loadTheme();
    if (typeof msal === "undefined") {
      setStatus("The Microsoft sign-in library (MSAL) could not be loaded. Check your network or ad blocker.", true);
      return;
    }
    if (!isConfigured()) {
      setStatus("Not configured yet — open config.js and fill in the Client ID, Excel file ID and folder ID.", true);
      return;
    }
    wireEvents();
    initAuth().then(function (existingAccount) {
      if (!existingAccount) {
        show(el.signIn, true);
        setStatus("Sign in with your Microsoft account to scan for orphan photos.");
        return;
      }
      account = existingAccount;
      msalApp.setActiveAccount(account);
      show(el.signOut, true);
      show(el.toolbar, true);
      refresh();
    }, function (error) {
      setStatus("Authentication could not start: " + errorMessage(error), true);
    });
  }

  start();
})();
