/* Home Inventory — shared, state-free helpers. */

(function (global) {
  "use strict";

  var TAG_SEPARATOR = " ";
  var ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
  var ISO_DATE_TIME = /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?$/;

  function fileExtension(name) {
    var match = /\.([a-z0-9]{1,5})$/i.exec(name || "");
    return match ? match[1].toLowerCase() : "jpg";
  }

  function errorMessage(error) {
    if (!error) return "Unknown error.";
    return error.message || String(error);
  }

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

  // Timestamp a row sorts by, or null when it has no usable value. An empty
  // UpdatedAt falls back to CreatedAt for rows written before that column existed.
  function rowTime(record, field) {
    var date = parseDate(record[field]);
    if (!date && field === "UpdatedAt") date = parseDate(record.CreatedAt);
    return date ? date.getTime() : null;
  }

  function compareText(a, b, directed) {
    return directed(a.toLowerCase().localeCompare(b.toLowerCase()));
  }

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

  global.HomeInventoryLib = {
    ISO_DATE: ISO_DATE,
    ISO_DATE_TIME: ISO_DATE_TIME,
    TAG_SEPARATOR: TAG_SEPARATOR,
    compareText: compareText,
    errorMessage: errorMessage,
    fileExtension: fileExtension,
    formatTags: formatTags,
    formatValue: formatValue,
    hasTag: hasTag,
    pad2: pad2,
    parseDate: parseDate,
    parseTags: parseTags,
    rowTime: rowTime
  };
})(window);
