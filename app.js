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
  var AUTO_COLUMNS = ["ID", "PhotoName", "CreatedAt", "UpdatedAt", "Status"];
  // Columns hidden from the card detail list (distinct from AUTO_COLUMNS).
  var CARD_HIDDEN_COLUMNS = ["ID", "PhotoName", "Status", "Tags"];
  var SEARCH_COLUMNS = ["Name", "Category", "Room", "Notes", "Tags"];
  // Column holding the item tags, stored as one separator-joined string.
  var TAGS_COLUMN = "Tags";
  var TAG_SEPARATOR = " ";
  // Free-text columns that also offer existing values as datalist suggestions.
  var SUGGEST_COLUMNS = ["Category", "Room"];
  // Subfolder of the HomeInventory folder where item photos are stored.
  var DEFAULT_PHOTO_FOLDER = "Photos";
  // How long a success confirmation stays on screen, in milliseconds.
  var TOAST_TIMEOUT = 4000;
  // Upper bound for the toast exit transition, after which the node is removed
  // even if no transitionend arrived.
  var TOAST_EXIT_TIMEOUT = 500;
  // How long the card of a freshly saved item stays highlighted, in milliseconds.
  var HIGHLIGHT_TIMEOUT = 2500;
  var SAVE_LABEL = "Save";
  // Fields the toolbar sort control offers, for both items and categories. The
  // value is the column name used for sorting; the label is what the user sees.
  var SORT_OPTIONS = [
    { value: "Name", label: "Name" },
    { value: "CreatedAt", label: "Created" },
    { value: "UpdatedAt", label: "Updated" }
  ];
  var SORT_FIELDS = SORT_OPTIONS.map(function (option) { return option.value; });
  var DEFAULT_SORT_FIELD = "UpdatedAt";
  var DEFAULT_SORT_DIRECTION = "desc";
  var SORT_STORAGE_KEY = "homeInventory.sort";
  var THEME_STORAGE_KEY = "homeInventory.theme";
  var UNCATEGORIZED = "Uncategorized";
  var NO_ROOM = "No room";
  // Columns the toolbar can group the list by. An empty value renders one flat
  // grid; the placeholder is used when the column is empty on a row.
  var GROUP_OPTIONS = [
    { value: "", label: "No group", placeholder: "" },
    { value: "Category", label: "By category", placeholder: UNCATEGORIZED },
    { value: "Room", label: "By room", placeholder: NO_ROOM }
  ];
  var GROUP_FIELDS = GROUP_OPTIONS.map(function (option) { return option.value; });
  var DEFAULT_GROUP_FIELD = "Category";
  var GROUP_STORAGE_KEY = "homeInventory.groupBy";

  var el = {
    status: document.getElementById("status"),
    toolbar: document.getElementById("toolbar"),
    items: document.getElementById("items"),
    empty: document.getElementById("empty"),
    search: document.getElementById("search"),
    filtersBtn: document.getElementById("filters-btn"),
    filtersPanel: document.getElementById("filters-panel"),
    filtersCount: document.getElementById("filters-count"),
    clearFiltersBtn: document.getElementById("clear-filters-btn"),
    filterCategory: document.getElementById("filter-category"),
    filterRoom: document.getElementById("filter-room"),
    tagFilter: document.getElementById("tag-filter"),
    sortField: document.getElementById("sort-field"),
    groupField: document.getElementById("group-field"),
    sortDirectionBtn: document.getElementById("sort-direction-btn"),
    sortDirectionIcon: document.getElementById("sort-direction-icon"),
    listHeader: document.getElementById("list-header"),
    themeToggleBtn: document.getElementById("theme-toggle-btn"),
    themeToggleIcon: document.getElementById("theme-toggle-icon"),
    tagFilterList: document.getElementById("tag-filter-list"),
    clearTagsBtn: document.getElementById("clear-tags-btn"),
    addBtn: document.getElementById("add-btn"),
    refreshBtn: document.getElementById("refresh-btn"),
    expandAllBtn: document.getElementById("expand-all-btn"),
    collapseAllBtn: document.getElementById("collapse-all-btn"),
    signIn: document.getElementById("signin-btn"),
    signOut: document.getElementById("signout-btn"),
    accountBtn: document.getElementById("account-btn"),
    accountPanel: document.getElementById("account-panel"),
    accountInitials: document.getElementById("account-initials"),
    accountPhoto: document.getElementById("account-photo"),
    accountPanelInitials: document.getElementById("account-panel-initials"),
    accountPanelPhoto: document.getElementById("account-panel-photo"),
    accountName: document.getElementById("account-name"),
    accountEmail: document.getElementById("account-email"),
    dialog: document.getElementById("item-dialog"),
    dialogTitle: document.getElementById("dialog-title"),
    form: document.getElementById("item-form"),
    formFields: document.getElementById("form-fields"),
    formError: document.getElementById("form-error"),
    photoInput: document.getElementById("photo-input"),
    photoCurrent: document.getElementById("photo-current"),
    photoPreview: document.getElementById("photo-preview"),
    photoThumb: document.getElementById("photo-thumb"),
    removePhotoBtn: document.getElementById("remove-photo-btn"),
    cancelBtn: document.getElementById("cancel-btn"),
    saveBtn: document.getElementById("save-btn"),
    saveLabel: document.getElementById("save-label"),
    saveSpinner: document.getElementById("save-spinner"),
    formErrorText: document.getElementById("form-error-text"),
    retryBtn: document.getElementById("retry-btn"),
    dismissErrorBtn: document.getElementById("dismiss-error-btn"),
    toastRegion: document.getElementById("toast-region"),
    photoDialog: document.getElementById("photo-dialog"),
    photoFull: document.getElementById("photo-full"),
    photoCloseBtn: document.getElementById("photo-close-btn")
  };

  var msalApp = null;
  var account = null;
  var headers = [];   // Excel column names, in sheet order
  var rows = [];      // [{ index: n, values: {Header: value} }]
  var editingId = null;
  var selectedTags = [];  // tag filter selection on the main screen
  var formTags = [];      // tags being edited in the item dialog
  var formTagChips = null; // container holding the tag chips of the item dialog
  var removePhotoFlag = false;
  var saving = false;
  var toastTimer = null;
  var photoUrlCache = {}; // PhotoName -> object URL
  var accountPhotoUrl = null; // object URL of the signed-in user's Graph photo
  // Grouping mode -> { group name: true } for every expanded panel. Empty on
  // load, so every panel starts collapsed, and kept per mode so switching the
  // grouping does not carry the state of another mode over.
  var expandedGroups = {};
  // Sort selection shared by categories and items; restored from localStorage.
  var sortField = DEFAULT_SORT_FIELD;
  var sortDirection = DEFAULT_SORT_DIRECTION;
  // Column the list is grouped by, "" for a flat list; restored from localStorage.
  var groupField = DEFAULT_GROUP_FIELD;
  var highlightTimer = null;
  // "light" or "dark" once chosen explicitly, null while following the OS.
  var theme = null;

  /* ---------------------------------------------------------------- helpers */

  function setStatus(message, isError) {
    el.status.textContent = message || "";
    el.status.hidden = !message;
    el.status.classList.toggle("error", !!isError);
  }

  function show(node, visible) {
    node.hidden = !visible;
  }

  // Avatar fallback: first letters of the first and last word, at most two.
  function initials(name) {
    var words = String(name || "").trim().split(/\s+/).filter(Boolean);
    if (words.length === 0) return "?";
    var letters = words.length === 1
      ? words[0].charAt(0)
      : words[0].charAt(0) + words[words.length - 1].charAt(0);
    return letters.toUpperCase();
  }

  // Transient confirmation shown in a polite live region so screen readers
  // announce it; it disappears on its own after a short delay.
  function showToast(message) {
    if (!el.toastRegion) return;
    if (toastTimer) clearTimeout(toastTimer);
    dismissToasts();
    var toast = document.createElement("div");
    toast.className = "toast";
    var text = document.createElement("p");
    text.textContent = message;
    toast.appendChild(text);
    el.toastRegion.appendChild(toast);
    toastTimer = setTimeout(function () {
      dismissToast(toast);
      toastTimer = null;
    }, TOAST_TIMEOUT);
  }

  // The node has to stay in the DOM while the exit transition runs, so it is
  // removed on transitionend, with a timer in case the transition never fires
  // (reduced motion, or a browser that skips it).
  function dismissToast(toast) {
    if (!toast || toast.className.indexOf("toast-exit") !== -1) return;
    toast.className = "toast toast-exit";
    toast.setAttribute("aria-hidden", "true");
    var remove = function () {
      if (toast.parentNode) toast.parentNode.removeChild(toast);
    };
    toast.addEventListener("transitionend", remove);
    setTimeout(remove, TOAST_EXIT_TIMEOUT);
  }

  function dismissToasts() {
    var toasts = el.toastRegion.querySelectorAll(".toast");
    Array.prototype.forEach.call(toasts, dismissToast);
  }

  // Errors stay visible (role="alert") until the user dismisses or retries.
  function showFormError(message, canRetry) {
    el.formErrorText.textContent = message;
    show(el.retryBtn, !!canRetry);
    show(el.formError, true);
  }

  function clearFormError() {
    el.formErrorText.textContent = "";
    show(el.retryBtn, false);
    show(el.formError, false);
  }

  function setSaving(isSaving) {
    saving = isSaving;
    el.saveBtn.disabled = isSaving;
    el.saveBtn.setAttribute("aria-busy", isSaving ? "true" : "false");
    el.cancelBtn.disabled = isSaving;
    show(el.saveSpinner, isSaving);
    el.saveLabel.textContent = isSaving ? "Saving…" : SAVE_LABEL;
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

  /* ------------------------------------------------------------------ dates */

  var ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
  var ISO_DATE_TIME = /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?$/;

  // Parses the date values the app stores (ISO timestamps) and the plain dates a
  // user may type into a custom column. Date-only values are read as local dates
  // so they are not shifted by a day west of UTC.
  function parseDate(value) {
    var text = String(value === undefined || value === null ? "" : value).trim();
    var date = null;
    if (ISO_DATE.test(text)) {
      var parts = text.split("-");
      date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    } else if (ISO_DATE_TIME.test(text)) {
      date = new Date(text);
    }
    return date && !isNaN(date.getTime()) ? date : null;
  }

  function pad2(number) {
    return (number < 10 ? "0" : "") + number;
  }

  // Every date shown in the UI is rendered in the viewer's local timezone as
  // dd.MM.yyyy HH:mm:ss (dd.MM.yyyy for date-only values); any other value is
  // passed through unchanged.
  function formatValue(value) {
    var text = String(value === undefined || value === null ? "" : value);
    var date = parseDate(text);
    if (!date) return text;
    var day = pad2(date.getDate()) + "." + pad2(date.getMonth() + 1) + "." + date.getFullYear();
    if (ISO_DATE.test(text.trim())) return day;
    return day + " " + pad2(date.getHours()) + ":" + pad2(date.getMinutes()) + ":" + pad2(date.getSeconds());
  }

  /* ------------------------------------------------------------------- theme */

  function prefersDark() {
    return !!(window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches);
  }

  function activeTheme() {
    return theme || (prefersDark() ? "dark" : "light");
  }

  function applyTheme() {
    var current = activeTheme();
    document.documentElement.setAttribute("data-theme", current);
    if (!el.themeToggleBtn) return;
    var label = current === "dark" ? "Switch to light theme" : "Switch to dark theme";
    el.themeToggleBtn.setAttribute("aria-pressed", current === "dark" ? "true" : "false");
    el.themeToggleBtn.setAttribute("aria-label", label);
    el.themeToggleBtn.title = label;
    el.themeToggleIcon.textContent = current === "dark" ? "☀" : "☾";
  }

  function loadTheme() {
    var stored = null;
    try {
      stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    } catch (e) { /* no or unusable storage — follow the OS preference */ }
    if (stored === "light" || stored === "dark") theme = stored;
    applyTheme();
    el.themeToggleBtn.addEventListener("click", toggleTheme);
    // Without an explicit choice the app keeps following the OS setting.
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

  /* ----------------------------------------------------------------- sorting */

  // The sort control is built from SORT_OPTIONS so the value sent to the sorting
  // logic stays a column name while the user sees a short, friendly label.
  function fillSortOptions() {
    el.sortField.textContent = "";
    SORT_OPTIONS.forEach(function (item) {
      var option = document.createElement("option");
      option.value = item.value;
      option.textContent = "Sort: " + item.label;
      el.sortField.appendChild(option);
    });
  }

  function loadSort() {
    var stored = null;
    try {
      stored = JSON.parse(window.localStorage.getItem(SORT_STORAGE_KEY));
    } catch (e) { /* no or unusable storage — keep the defaults */ }
    if (stored && SORT_FIELDS.indexOf(stored.field) !== -1) sortField = stored.field;
    if (stored && (stored.direction === "asc" || stored.direction === "desc")) {
      sortDirection = stored.direction;
    }
  }

  function saveSort() {
    try {
      window.localStorage.setItem(SORT_STORAGE_KEY, JSON.stringify({
        field: sortField,
        direction: sortDirection
      }));
    } catch (e) { /* storage unavailable — the selection stays for this session */ }
  }

  // Timestamp a row sorts by, or null when it has no usable value. An empty
  // UpdatedAt falls back to CreatedAt for rows written before that column existed.
  function rowTime(record, field) {
    var date = parseDate(record[field]);
    if (!date && field === "UpdatedAt") date = parseDate(record.CreatedAt);
    return date ? date.getTime() : null;
  }

  function directed(comparison) {
    return sortDirection === "desc" ? -comparison : comparison;
  }

  function compareText(a, b) {
    return directed(a.toLowerCase().localeCompare(b.toLowerCase()));
  }

  // Rows without a value sort last whatever the direction is.
  function compareTime(a, b) {
    if (a === null && b === null) return 0;
    if (a === null) return 1;
    if (b === null) return -1;
    return directed(a < b ? -1 : (a > b ? 1 : 0));
  }

  function compareRows(a, b) {
    if (sortField === "Name") return compareText(a.values.Name || "", b.values.Name || "");
    return compareTime(rowTime(a.values, sortField), rowTime(b.values, sortField));
  }

  // A group sorts by the most recent value of its items; with Name selected
  // groups keep their natural alphabetical order instead.
  function groupTime(groupRows) {
    var best = null;
    groupRows.forEach(function (row) {
      var time = rowTime(row.values, sortField);
      if (time !== null && (best === null || time > best)) best = time;
    });
    return best;
  }

  function compareGroups(a, b, groups) {
    if (sortField === "Name") return compareText(a, b);
    return compareTime(groupTime(groups[a]), groupTime(groups[b]));
  }

  /* ---------------------------------------------------------------- grouping */

  // The group control mirrors the sort one: the value is a column name (or ""
  // for a flat list) while the user sees a short label.
  function fillGroupOptions() {
    el.groupField.textContent = "";
    GROUP_OPTIONS.forEach(function (item) {
      var option = document.createElement("option");
      option.value = item.value;
      option.textContent = item.label;
      el.groupField.appendChild(option);
    });
  }

  function loadGroup() {
    var stored = null;
    try {
      stored = window.localStorage.getItem(GROUP_STORAGE_KEY);
    } catch (e) { /* no or unusable storage — keep the default */ }
    if (stored !== null && GROUP_FIELDS.indexOf(stored) !== -1) groupField = stored;
  }

  function saveGroup() {
    try {
      window.localStorage.setItem(GROUP_STORAGE_KEY, groupField);
    } catch (e) { /* storage unavailable — the selection stays for this session */ }
  }

  // Placeholder shown for rows with no value in the grouping column.
  function groupPlaceholder(field) {
    var placeholder = "";
    GROUP_OPTIONS.forEach(function (option) {
      if (option.value === field) placeholder = option.placeholder;
    });
    return placeholder;
  }

  function groupKey(record, field) {
    return record[field] || groupPlaceholder(field);
  }

  // Buckets rows by the grouping column, keeping the incoming row order inside
  // each group and returning the group names in their sorted order.
  function groupRows(list, field) {
    var groups = {};
    var order = [];
    list.forEach(function (row) {
      var key = groupKey(row.values, field);
      if (!groups[key]) {
        groups[key] = [];
        order.push(key);
      }
      groups[key].push(row);
    });
    order.sort(function (a, b) {
      return compareGroups(a, b, groups);
    });
    return { groups: groups, order: order };
  }

  // Expanded panels of the active grouping mode, created on first use so each
  // mode keeps its own state.
  function expandedForGroup() {
    if (!expandedGroups[groupField]) expandedGroups[groupField] = {};
    return expandedGroups[groupField];
  }

  /* ------------------------------------------------------------------- tags */

  // Tags live in one cell as a separator-joined string; whitespace around each
  // tag is dropped and empty entries are ignored.
  function parseTags(value) {
    return String(value || "").split(TAG_SEPARATOR).map(function (tag) {
      return tag.trim();
    }).filter(function (tag) {
      return tag !== "";
    });
  }

  function formatTags(tags) {
    return tags.join(TAG_SEPARATOR);
  }

  function hasTag(tags, tag) {
    var wanted = tag.toLowerCase();
    return tags.some(function (existing) {
      return existing.toLowerCase() === wanted;
    });
  }

  // Every tag used by the visible (not deleted) rows, sorted alphabetically.
  function allTags() {
    var tags = [];
    rows.forEach(function (row) {
      if (row.values.Status === "deleted") return;
      parseTags(row.values[TAGS_COLUMN]).forEach(function (tag) {
        if (!hasTag(tags, tag)) tags.push(tag);
      });
    });
    tags.sort(function (a, b) {
      return a.toLowerCase().localeCompare(b.toLowerCase());
    });
    return tags;
  }

  function hasTagsColumn() {
    return headers.indexOf(TAGS_COLUMN) !== -1;
  }

  function tagChip(tag) {
    var chip = document.createElement("span");
    chip.className = "tag";
    chip.textContent = tag;
    return chip;
  }

  /* ------------------------------------------------------------------- auth */

  function initAuth() {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: CONFIG.authority,
        redirectUri: CONFIG.redirectUri,
        navigateToLoginRequestUrl: true
      },
      cache: {
        // localStorage survives the full-page redirect round trip; Chrome
        // partitions/blocks some sessionStorage + cookie access during auth.
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
      // Silent renewal failed — send the user through a full-page redirect.
      msalApp.acquireTokenRedirect({ scopes: CONFIG.scopes, account: account });
      return new Promise(function () { /* navigation in progress */ });
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
    revokeAccountPhoto();
    account = null;
    render();
    show(el.toolbar, false);
    show(el.listHeader, false);
    showFilters(false);
    show(el.items, false);
    show(el.empty, false);
    showAccountMenu(false);
    show(el.accountBtn, false);
    show(el.signIn, true);
    setStatus("Signing out…");
    msalApp.logoutRedirect({
      postLogoutRedirectUri: CONFIG.redirectUri
    }).catch(function () { /* ignore */ });
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

  function photoFolderName() {
    return CONFIG.photoFolder || DEFAULT_PHOTO_FOLDER;
  }

  // Photos live in a dedicated subfolder of the HomeInventory folder, so the
  // Excel file and the images are not mixed together. Only the file name is
  // stored in Excel; it is always resolved against that subfolder.
  function photoPath(name) {
    return "/me/drive/items/" + encodeURIComponent(CONFIG.folderId) +
      ":/" + encodeURIComponent(photoFolderName()) +
      "/" + encodeURIComponent(name) + ":";
  }

  // Graph returns 404 when either the photo or the Photos subfolder is absent.
  function photoError(error) {
    if (error && /Graph 404/.test(error.message || "")) {
      return new Error(
        "Not found in the \"" + photoFolderName() + "\" subfolder of your " +
        "HomeInventory folder. Create that subfolder in OneDrive (see README) " +
        "and make sure the photo is inside it."
      );
    }
    return error;
  }

  function loadTable() {
    return graph(tablePath() + "/columns?$select=name").then(function (data) {
      headers = (data && data.value ? data.value : []).map(function (column) {
        return column.name;
      });
      if (headers.length > 0 && headers.indexOf("Status") === -1) {
        setStatus(
          "The spreadsheet is missing a \"Status\" column. " +
          "Please add a \"Status\" column to Table1 in inventory.xlsx and reload.",
          true
        );
        rows = [];
        return null;
      }
      return graph(tablePath() + "/rows?$select=index,values");
    }).then(function (data) {
      if (!data) return;
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

  function uploadPhoto(file) {
    var name = "photo-" + newId() + "." + fileExtension(file.name);
    return graph(photoPath(name) + "/content", {
      method: "PUT",
      body: file,
      isRaw: true,
      headers: { "Content-Type": file.type || "application/octet-stream" }
    }).then(function () {
      return name;
    }, function (error) {
      throw photoError(error);
    });
  }

  function photoUrl(name) {
    if (photoUrlCache[name]) return Promise.resolve(photoUrlCache[name]);
    return graph(photoPath(name) + "/content", { asBlob: true }).then(function (blob) {
      var url = URL.createObjectURL(blob);
      photoUrlCache[name] = url;
      return url;
    }, function (error) {
      throw photoError(error);
    });
  }

  function revokePhotoUrls() {
    Object.keys(photoUrlCache).forEach(function (name) {
      URL.revokeObjectURL(photoUrlCache[name]);
    });
    photoUrlCache = {};
  }

  function revokeAccountPhoto() {
    if (accountPhotoUrl) URL.revokeObjectURL(accountPhotoUrl);
    accountPhotoUrl = null;
    el.accountPhoto.removeAttribute("src");
    el.accountPanelPhoto.removeAttribute("src");
    show(el.accountPhoto, false);
    show(el.accountPanelPhoto, false);
    show(el.accountInitials, true);
    show(el.accountPanelInitials, true);
  }

  // Graph answers 404 when the account has no profile photo; the initials
  // avatar rendered underneath then simply stays visible.
  function loadAccountPhoto() {
    return graph("/me/photo/$value", { asBlob: true }).then(function (blob) {
      if (!account) return;
      accountPhotoUrl = URL.createObjectURL(blob);
      el.accountPhoto.src = accountPhotoUrl;
      el.accountPanelPhoto.src = accountPhotoUrl;
      show(el.accountInitials, false);
      show(el.accountPanelInitials, false);
      show(el.accountPhoto, true);
      show(el.accountPanelPhoto, true);
    }, function () { /* no photo — keep the initials */ });
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

  function matchesTags(record) {
    if (selectedTags.length === 0) return true;
    var tags = parseTags(record[TAGS_COLUMN]);
    return selectedTags.every(function (tag) {
      return hasTag(tags, tag);
    });
  }

  function visibleRows() {
    var term = el.search.value.trim().toLowerCase();
    var category = el.filterCategory.value;
    var room = el.filterRoom.value;
    return rows.filter(function (row) {
      var record = row.values;
      if (record.Status === "deleted") return false;
      if (!matchesSearch(record, term)) return false;
      if (category && (record.Category || "") !== category) return false;
      if (room && (record.Room || "") !== room) return false;
      if (!matchesTags(record)) return false;
      return true;
    });
  }

  function distinctValues(column) {
    var values = [];
    rows.forEach(function (row) {
      if (row.values.Status === "deleted") return;
      var value = row.values[column];
      if (value && values.indexOf(value) === -1) values.push(value);
    });
    values.sort();
    return values;
  }

  function fillFilter(select, column, allLabel) {
    var previous = select.value;
    var options = distinctValues(column);
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
    show(select.parentNode.classList.contains("filters-field") ? select.parentNode : select,
      headers.indexOf(column) !== -1);
  }

  function toggleTagFilter(tag) {
    if (hasTag(selectedTags, tag)) {
      selectedTags = selectedTags.filter(function (existing) {
        return existing.toLowerCase() !== tag.toLowerCase();
      });
    } else {
      selectedTags.push(tag);
    }
    fillTagFilter();
    render();
  }

  // Tag filter: one toggle button per known tag; selecting several narrows the
  // list to items carrying all of them.
  function fillTagFilter() {
    var tags = allTags();
    selectedTags = selectedTags.filter(function (tag) {
      return hasTag(tags, tag);
    });
    el.tagFilterList.textContent = "";
    tags.forEach(function (tag) {
      var button = document.createElement("button");
      button.type = "button";
      button.className = "tag tag-toggle";
      button.textContent = tag;
      button.setAttribute("aria-pressed", hasTag(selectedTags, tag) ? "true" : "false");
      button.addEventListener("click", function () { toggleTagFilter(tag); });
      el.tagFilterList.appendChild(button);
    });
    show(el.clearTagsBtn, selectedTags.length > 0);
    show(el.tagFilter, hasTagsColumn() && tags.length > 0);
    syncFilterControls();
  }

  // Number of filters (category, room, tags) currently narrowing the list.
  function activeFilterCount() {
    var count = selectedTags.length;
    if (el.filterCategory.value) count += 1;
    if (el.filterRoom.value) count += 1;
    return count;
  }

  function syncFilterControls() {
    var count = activeFilterCount();
    el.filtersCount.textContent = String(count);
    show(el.filtersCount, count > 0);
    el.clearFiltersBtn.disabled = count === 0;
  }

  function showFilters(open) {
    show(el.filtersPanel, open);
    el.filtersBtn.setAttribute("aria-expanded", open ? "true" : "false");
  }

  function clearFilters() {
    if (activeFilterCount() === 0) return;
    el.filterCategory.value = "";
    el.filterRoom.value = "";
    selectedTags = [];
    fillTagFilter();
    render();
  }

  function clearTagFilter() {
    if (selectedTags.length === 0) return;
    selectedTags = [];
    fillTagFilter();
    render();
  }

  function buildCard(row) {
    var record = row.values;
    var card = document.createElement("article");
    card.className = "card";
    if (record.ID) card.setAttribute("data-item-id", record.ID);

    // The media box keeps its 4:3 aspect ratio whether or not a photo loads, so
    // every card in a row stays the same height.
    var media = document.createElement("div");
    media.className = "card-media";
    if (record.PhotoName) {
      var img = document.createElement("img");
      img.className = "thumb";
      img.alt = record.Name || "Item photo";
      img.addEventListener("click", function () {
        openPhoto(record.PhotoName, img.alt);
      });
      media.appendChild(img);
      photoUrl(record.PhotoName).then(function (url) {
        img.src = url;
      }, function () {
        img.remove();
        media.className = "card-media card-media-empty";
      });
    } else {
      media.className = "card-media card-media-empty";
    }
    card.appendChild(media);

    var body = document.createElement("div");
    body.className = "card-body";

    var title = document.createElement("h3");
    title.textContent = record.Name || "(no name)";
    // The heading clamps to two lines, so keep the full name in the tooltip.
    title.title = title.textContent;
    body.appendChild(title);

    var list = document.createElement("dl");
    headers.forEach(function (header) {
      if (header === "Name" || CARD_HIDDEN_COLUMNS.indexOf(header) !== -1) return;
      var value = record[header];
      if (!value) return;
      var dt = document.createElement("dt");
      dt.textContent = header;
      var dd = document.createElement("dd");
      dd.textContent = formatValue(value);
      list.appendChild(dt);
      list.appendChild(dd);
    });
    body.appendChild(list);

    var tags = parseTags(record[TAGS_COLUMN]);
    if (tags.length > 0) {
      var tagList = document.createElement("div");
      tagList.className = "tags";
      tags.forEach(function (tag) {
        tagList.appendChild(tagChip(tag));
      });
      body.appendChild(tagList);
    }

    card.appendChild(body);

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

    var sorted = visible.slice().sort(compareRows);

    if (!groupField) {
      var flat = document.createElement("div");
      flat.className = "cards";
      sorted.forEach(function (row) {
        flat.appendChild(buildCard(row));
      });
      el.items.appendChild(flat);
      return;
    }

    // Totals ignore the current filter — but not deletions — so the header can
    // show how much the filter is hiding.
    var totals = {};
    rows.forEach(function (row) {
      if (row.values.Status === "deleted") return;
      var name = groupKey(row.values, groupField);
      totals[name] = (totals[name] || 0) + 1;
    });
    var grouped = groupRows(sorted, groupField);
    var groups = grouped.groups;
    var expanded = expandedForGroup();

    grouped.order.forEach(function (category) {
      var panel = document.createElement("details");
      panel.className = "category-panel";
      panel.open = expanded[category] === true;
      panel.addEventListener("toggle", function () {
        expanded[category] = panel.open;
      });

      var heading = document.createElement("summary");
      heading.className = "category-title";
      var name = document.createElement("span");
      name.className = "category-name";
      name.textContent = category;
      heading.appendChild(name);

      var shown = groups[category].length;
      var total = totals[category] || shown;
      var count = document.createElement("span");
      count.className = "category-count";
      // The bracketed figures read as noise to a screen reader; the prose
      // sibling below carries the same information.
      count.setAttribute("aria-hidden", "true");
      count.textContent = shown < total
        ? "(" + shown + " / " + total + ")"
        : "(" + total + ")";
      heading.appendChild(count);

      var countLabel = document.createElement("span");
      countLabel.className = "visually-hidden";
      countLabel.textContent = shown < total
        ? shown + " of " + total + " items shown"
        : total + (total === 1 ? " item" : " items");
      heading.appendChild(countLabel);
      panel.appendChild(heading);

      var grid = document.createElement("div");
      grid.className = "cards";
      groups[category].forEach(function (row) {
        grid.appendChild(buildCard(row));
      });
      panel.appendChild(grid);
      el.items.appendChild(panel);
    });
  }

  function setAllPanels(open) {
    var panels = el.items.querySelectorAll(".category-panel");
    Array.prototype.forEach.call(panels, function (panel) {
      panel.open = open;
    });
  }

  // After a save the item may sit inside a collapsed panel: open only that
  // panel, leave the others as they are, and flag the card so the user sees
  // where the saved item landed.
  function revealItem(id) {
    var row = findRowById(id);
    if (!row) return;
    if (groupField) {
      var expanded = expandedForGroup();
      var key = groupKey(row.values, groupField);
      if (expanded[key] !== true) {
        expanded[key] = true;
        render();
      }
    }
    var card = null;
    var cards = el.items.querySelectorAll(".card");
    Array.prototype.forEach.call(cards, function (candidate) {
      if (candidate.getAttribute("data-item-id") === id) card = candidate;
    });
    if (!card) return;
    if (highlightTimer) clearTimeout(highlightTimer);
    card.classList.add("card-highlight");
    card.scrollIntoView({ behavior: "smooth", block: "center" });
    highlightTimer = setTimeout(function () {
      card.classList.remove("card-highlight");
      highlightTimer = null;
    }, HIGHLIGHT_TIMEOUT);
  }

  /* ------------------------------------------------------------------- form */

  function fieldId(header) {
    return "field-" + header.replace(/[^A-Za-z0-9_-]/g, "_");
  }

  function buildDatalist(header) {
    var list = document.createElement("datalist");
    list.id = fieldId(header) + "-options";
    distinctValues(header).forEach(function (value) {
      var option = document.createElement("option");
      option.value = value;
      list.appendChild(option);
    });
    return list;
  }

  // Renders the chips of the tags currently attached to the edited item; each
  // chip carries a button that removes that single tag.
  function renderFormTags() {
    var list = formTagChips;
    if (!list) return;
    list.textContent = "";
    formTags.forEach(function (tag) {
      var chip = document.createElement("span");
      chip.className = "tag tag-editable";
      var text = document.createElement("span");
      text.textContent = tag;
      var remove = document.createElement("button");
      remove.type = "button";
      remove.className = "tag-remove";
      remove.textContent = "×";
      remove.setAttribute("aria-label", "Remove tag " + tag);
      remove.addEventListener("click", function () {
        formTags = formTags.filter(function (existing) {
          return existing !== tag;
        });
        renderFormTags();
      });
      chip.appendChild(text);
      chip.appendChild(remove);
      list.appendChild(chip);
    });
  }

  // Accepts both a picked suggestion and freshly typed text; several tags can
  // be pasted at once because the separator is honoured here too.
  function addFormTags(text) {
    parseTags(text).forEach(function (tag) {
      if (!hasTag(formTags, tag)) formTags.push(tag);
    });
    renderFormTags();
  }

  function buildTagField(wrapper, record) {
    formTags = parseTags(record ? record[TAGS_COLUMN] : "");

    var chips = document.createElement("div");
    chips.className = "tags";
    formTagChips = chips;

    var input = document.createElement("input");
    input.type = "text";
    input.id = fieldId(TAGS_COLUMN);
    input.name = TAGS_COLUMN;
    input.autocomplete = "off";
    input.placeholder = "Type a tag and press Enter";

    var datalist = document.createElement("datalist");
    datalist.id = fieldId(TAGS_COLUMN) + "-options";
    allTags().forEach(function (tag) {
      var option = document.createElement("option");
      option.value = tag;
      datalist.appendChild(option);
    });
    input.setAttribute("list", datalist.id);

    input.addEventListener("keydown", function (event) {
      if (event.key === "Enter" || event.key === TAG_SEPARATOR || event.key === ",") {
        // Enter would otherwise submit the dialog form.
        event.preventDefault();
        addFormTags(input.value);
        input.value = "";
      } else if (event.key === "Backspace" && input.value === "" && formTags.length > 0) {
        formTags.pop();
        renderFormTags();
      }
    });
    // Picking a suggestion replaces the typed text in one go; turn it into a
    // chip straight away. Text still being typed is only turned into a chip on
    // Enter, so the dialog does not resize under the pointer, and readForm()
    // keeps whatever is left in the box when the item is saved.
    input.addEventListener("input", function (event) {
      var inputType = event.inputType;
      if (inputType && inputType !== "insertReplacementText") return;
      if (!hasTag(allTags(), input.value.trim())) return;
      addFormTags(input.value);
      input.value = "";
    });

    var hint = document.createElement("p");
    hint.className = "hint";
    hint.textContent = "Pick an existing tag or type a new one, then press Enter.";

    wrapper.appendChild(chips);
    wrapper.appendChild(input);
    wrapper.appendChild(datalist);
    wrapper.appendChild(hint);
    renderFormTags();
  }

  function buildForm(record) {
    // Clearing the container also drops the datalists built for the previous
    // dialog, so their IDs are never duplicated.
    el.formFields.textContent = "";
    formTags = [];
    formTagChips = null;
    activeHeaders().forEach(function (header) {
      var wrapper = document.createElement("div");
      wrapper.className = "form-row";

      var label = document.createElement("label");
      label.textContent = header;
      label.htmlFor = fieldId(header);

      if (header === TAGS_COLUMN) {
        wrapper.appendChild(label);
        buildTagField(wrapper, record);
        el.formFields.appendChild(wrapper);
        return;
      }

      var input;
      var datalist = null;
      if (header === "Notes") {
        input = document.createElement("textarea");
        input.rows = 3;
      } else {
        input = document.createElement("input");
        input.type = header === "Quantity" ? "number" : "text";
        if (header === "Quantity") input.min = "0";
        if (SUGGEST_COLUMNS.indexOf(header) !== -1) {
          datalist = buildDatalist(header);
          input.setAttribute("list", datalist.id);
          input.autocomplete = "off";
        }
      }
      input.id = fieldId(header);
      input.name = header;
      input.value = record ? (record[header] || "") : (header === "Quantity" ? "1" : "");
      if (header === "Name") input.required = true;

      wrapper.appendChild(label);
      wrapper.appendChild(input);
      if (datalist) wrapper.appendChild(datalist);
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
    removePhotoFlag = false;
    var row = id ? findRowById(id) : null;
    el.dialogTitle.textContent = row ? "Edit Item" : "Add Item";
    buildForm(row ? row.values : null);
    el.photoInput.value = "";
    var current = row && row.values.PhotoName;
    el.photoCurrent.textContent = current ? "Current photo: " + current + " (choose a file to replace it)" : "";
    show(el.photoCurrent, !!current);
    show(el.photoPreview, false);
    el.photoThumb.src = "";
    if (current) {
      photoUrl(current).then(function (url) {
        el.photoThumb.src = url;
        show(el.photoPreview, true);
        show(el.photoCurrent, false);
      }, function () {
        // Failed to load — fall back to the text hint already shown
      });
    }
    clearFormError();
    setSaving(false);
    el.dialog.showModal();
  }

  function readForm() {
    var record = {};
    activeHeaders().forEach(function (header) {
      var input = document.getElementById(fieldId(header));
      if (header === TAGS_COLUMN) {
        // Whatever is still typed in the tag box counts as a tag as well.
        if (input) addFormTags(input.value);
        if (input) input.value = "";
        record[header] = formatTags(formTags);
        return;
      }
      record[header] = input ? input.value.trim() : "";
    });
    return record;
  }

  function saveItem(event) {
    if (event) event.preventDefault();
    if (saving) return;

    var input = readForm();
    if (headers.indexOf("Name") !== -1 && !input.Name) {
      showFormError("Name is required.", false);
      return;
    }

    var existing = editingId ? findRowById(editingId) : null;
    var file = el.photoInput.files && el.photoInput.files[0];

    setSaving(true);
    clearFormError();

    var uploaded = file ? uploadPhoto(file) : Promise.resolve(null);
    var savedId = null;

    uploaded.then(function (photoName) {
      var record = Object.assign({}, existing ? existing.values : {}, input);
      if (photoName) {
        record.PhotoName = photoName;
      } else if (removePhotoFlag) {
        record.PhotoName = "";
      }
      if (!existing) {
        record.ID = newId();
        record.CreatedAt = new Date().toISOString();
      }
      record.UpdatedAt = new Date().toISOString();
      savedId = record.ID;
      return existing ? updateRow(existing.index, record) : addRow(record);
    }).then(function () {
      el.dialog.close();
      showToast("Saved");
      return refresh();
    }).then(function () {
      // The sort selection is deliberately left untouched here; the user's
      // persisted preference decides where the saved item shows up.
      if (savedId) revealItem(savedId);
    }).catch(function (error) {
      showFormError("Save failed: " + errorMessage(error), true);
    }).then(function () {
      setSaving(false);
    });
  }

  function confirmDelete(record) {
    if (!window.confirm('Delete "' + (record.Name || "this item") + '"? It will be hidden from the app but kept in the spreadsheet.')) return;
    var row = findRowById(record.ID);
    if (!row) return;
    setStatus("Deleting…");
    var updated = Object.assign({}, row.values, {
      Status: "deleted",
      UpdatedAt: new Date().toISOString()
    });
    updateRow(row.index, updated).then(function () {
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
      fillTagFilter();
      render();
    }, function (error) {
      setStatus("Could not load the inventory: " + errorMessage(error), true);
    });
  }

  function onSignedIn() {
    var name = account.name || account.username || "";
    el.accountName.textContent = name;
    el.accountEmail.textContent = account.username || "";
    el.accountInitials.textContent = initials(name);
    el.accountPanelInitials.textContent = el.accountInitials.textContent;
    el.accountBtn.setAttribute("aria-label", "Account: " + name);
    el.accountBtn.title = name;
    show(el.signIn, false);
    show(el.accountBtn, true);
    loadAccountPhoto();
    show(el.toolbar, true);
    show(el.listHeader, true);
    show(el.items, true);
    refresh();
  }

  function showAccountMenu(open) {
    show(el.accountPanel, open);
    el.accountBtn.setAttribute("aria-expanded", open ? "true" : "false");
  }

  function syncSortControls() {
    el.sortField.value = sortField;
    var descending = sortDirection === "desc";
    el.sortDirectionIcon.textContent = descending ? "↓" : "↑";
    el.sortDirectionBtn.setAttribute(
      "aria-label",
      "Sort direction: " + (descending ? "descending" : "ascending")
    );
    el.sortDirectionBtn.title = el.sortDirectionBtn.getAttribute("aria-label");
  }

  // Without panels there is nothing to expand or collapse, so both buttons go
  // away in the flat mode.
  function syncGroupControls() {
    el.groupField.value = groupField;
    show(el.expandAllBtn, !!groupField);
    show(el.collapseAllBtn, !!groupField);
  }

  function wireEvents() {
    el.signIn.addEventListener("click", signIn);
    el.signOut.addEventListener("click", signOut);
    el.addBtn.addEventListener("click", function () { openForm(null); });
    el.refreshBtn.addEventListener("click", function () { refresh(); });
    el.expandAllBtn.addEventListener("click", function () { setAllPanels(true); });
    el.collapseAllBtn.addEventListener("click", function () { setAllPanels(false); });
    el.search.addEventListener("input", render);
    el.filterCategory.addEventListener("change", function () {
      syncFilterControls();
      render();
    });
    el.filterRoom.addEventListener("change", function () {
      syncFilterControls();
      render();
    });
    el.filtersBtn.addEventListener("click", function () {
      showFilters(el.filtersPanel.hidden);
    });
    el.accountBtn.addEventListener("click", function () {
      showAccountMenu(el.accountPanel.hidden);
    });
    el.clearFiltersBtn.addEventListener("click", clearFilters);
    // Close the filters popover on an outside click or Escape, like a menu.
    document.addEventListener("click", function (event) {
      if (el.filtersPanel.hidden) return;
      // Tag toggles rebuild themselves, so a click target may already be gone
      // from the document by the time this runs; that is not an outside click.
      if (!document.contains(event.target)) return;
      if (el.filtersBtn.contains(event.target) || el.filtersPanel.contains(event.target)) return;
      showFilters(false);
    });
    document.addEventListener("keydown", function (event) {
      if (event.key !== "Escape" || el.filtersPanel.hidden) return;
      showFilters(false);
      el.filtersBtn.focus();
    });
    // The account popover behaves like the filters one: outside click, Escape.
    document.addEventListener("click", function (event) {
      if (el.accountPanel.hidden) return;
      if (!document.contains(event.target)) return;
      if (el.accountBtn.contains(event.target) || el.accountPanel.contains(event.target)) return;
      showAccountMenu(false);
    });
    document.addEventListener("keydown", function (event) {
      if (event.key !== "Escape" || el.accountPanel.hidden) return;
      showAccountMenu(false);
      el.accountBtn.focus();
    });
    el.sortField.addEventListener("change", function () {
      sortField = SORT_FIELDS.indexOf(el.sortField.value) !== -1
        ? el.sortField.value
        : DEFAULT_SORT_FIELD;
      saveSort();
      syncSortControls();
      render();
    });
    el.sortDirectionBtn.addEventListener("click", function () {
      sortDirection = sortDirection === "desc" ? "asc" : "desc";
      saveSort();
      syncSortControls();
      render();
    });
    el.groupField.addEventListener("change", function () {
      groupField = GROUP_FIELDS.indexOf(el.groupField.value) !== -1
        ? el.groupField.value
        : DEFAULT_GROUP_FIELD;
      saveGroup();
      syncGroupControls();
      render();
    });
    el.clearTagsBtn.addEventListener("click", clearTagFilter);
    el.form.addEventListener("submit", saveItem);
    el.cancelBtn.addEventListener("click", function () { el.dialog.close(); });
    el.retryBtn.addEventListener("click", function () { saveItem(); });
    el.dismissErrorBtn.addEventListener("click", clearFormError);
    el.removePhotoBtn.addEventListener("click", function () {
      removePhotoFlag = true;
      show(el.photoPreview, false);
      el.photoThumb.src = "";
      show(el.photoCurrent, false);
    });
    el.photoCloseBtn.addEventListener("click", function () { el.photoDialog.close(); });
    el.photoDialog.addEventListener("click", function (event) {
      if (event.target === el.photoDialog) { el.photoDialog.close(); }
    });
  }

  window.addEventListener("error", function (e) {
    var msg = (e.error && e.error.message) ? e.error.message : (e.message || "An unexpected error occurred.");
    console.error("Unhandled error:", e.error || e);
    setStatus("An unexpected error occurred: " + msg, true);
  });

  window.addEventListener("unhandledrejection", function (e) {
    var msg = (e.reason && e.reason.message) ? e.reason.message : (e.reason ? String(e.reason) : "An unhandled error occurred.");
    console.error("Unhandled rejection:", e.reason);
    setStatus("An unexpected error occurred: " + msg, true);
  });

  function start() {
    try {
      loadTheme();
      if (typeof msal === "undefined") {
        setStatus("The Microsoft sign-in library (MSAL) could not be loaded. Check your network or ad blocker.", true);
        return;
      }
      if (typeof CONFIG === "undefined") {
        setStatus("Configuration file (config.js) could not be loaded. Check your deployment and network connection.", true);
        return;
      }
      if (!isConfigured()) {
        setStatus("Not configured yet — open config.js and fill in the Client ID, Excel file ID and folder ID. See README for instructions.", true);
        return;
      }
      wireEvents();
      fillSortOptions();
      loadSort();
      syncSortControls();
      fillGroupOptions();
      loadGroup();
      syncGroupControls();
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
        console.error("Authentication error:", error);
        setStatus("Authentication could not start: " + errorMessage(error), true);
      });
    } catch (error) {
      console.error("Application startup error:", error);
      setStatus("Application failed to start: " + errorMessage(error), true);
    }
  }

  start();
})();
