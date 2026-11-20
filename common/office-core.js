/* Office Suite — shared code used by Letter Generator, Notice and Proceedings.

   UPDATING: the pages load the files in common/ with a version stamp
   (e.g. common/office-core.js?v=20261119). After changing any file in
   common/, raise that number in index.html, Letter Generator.html,
   Notice.html and Proceedings.html, so browsers fetch the new copy
   instead of an old cached one.
   Loaded with a plain <script src> (not a module) so the apps still work when
   opened straight from disk (file://). Everything here is global, exactly as it
   was when each app held its own copy. Variables such as lastFocusedField,
   showToast and renderMergeTable are still defined by each app; they are only
   looked up when a function runs, so load order is not a problem.
*/

/* ══════════════════════════════════════
   Messages (toast at the bottom right)
   ══════════════════════════════════════
   showToast(message)           green, 3 seconds
   showToast(message, "error")  red, 6 seconds
   Without a type, a message worded as a problem ("Error...", "Could not...",
   "Cannot...", "...failed", "...required", "Please...") is shown as an error.
   Click the message to close it early. Used by all three apps instead of
   alert boxes; confirm boxes are still used where a yes/no answer is needed. */
var toastTimer = null;

function looksLikeError(message) {
  return /^\s*(❌\s*)?(error|could not|cannot|invalid|failed|copy failed|please)\b|\b(failed|required|must have|not found|is blank)\b|\bfirst,/i.test(
    String(message),
  );
}

function showToast(message, type) {
  type = type || (looksLikeError(message) ? "error" : "success");
  var toast = document.getElementById("toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "toast";
    toast.innerHTML = '<span id="toastMessage"></span>';
    document.body.appendChild(toast);
  }
  toast.onclick = function () {
    toast.classList.remove("show");
  };
  toast.className = "toast " + type;
  void toast.offsetWidth; // restart the slide-in for a new message
  document.getElementById("toastMessage").textContent = message;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(
    function () {
      toast.classList.remove("show");
    },
    type === "error" ? 6000 : 3000,
  );
}

/* ══════════════════════════════════════
   Dates
   ══════════════════════════════════════ */

function todayFormatted() {
  var d = new Date();
  return (
    String(d.getDate()).padStart(2, "0") +
    "/" +
    String(d.getMonth() + 1).padStart(2, "0") +
    "/" +
    d.getFullYear()
  );
}


/* ══════════════════════════════════════
   Office profile (read from URL parameter or browser storage)
   ══════════════════════════════════════ */

function getOfficeProfile() {
  // 1. Try URL parameter (passed by index.html — works reliably with file:// protocol)
  try {
    var params = new URLSearchParams(window.location.search);
    var encoded = params.get("profile");
    if (encoded) {
      return JSON.parse(decodeURIComponent(escape(atob(encoded))));
    }
  } catch (e) {}
  // 2. Fall back to localStorage
  try {
    var raw = localStorage.getItem("officeProfile");
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}


/* ══════════════════════════════════════
   Preview zoom and fullscreen
   ══════════════════════════════════════ */

var previewZoom = 0.75;

function applyZoom() {
  document.querySelectorAll(".page-container").forEach(function (el) {
    el.style.zoom = previewZoom;
  });
  var el = document.getElementById("zoomLevel");
  if (el) el.textContent = Math.round(previewZoom * 100) + "%";
}

function zoomIn() {
  if (previewZoom < 2.0) {
    previewZoom = Math.round((previewZoom + 0.1) * 10) / 10;
    applyZoom();
  }
}

function zoomOut() {
  if (previewZoom > 0.3) {
    previewZoom = Math.round((previewZoom - 0.1) * 10) / 10;
    applyZoom();
  }
}

function zoomReset() {
  previewZoom = 0.75;
  applyZoom();
}

function toggleFullscreen() {
  var panel = document.getElementById("previewPanel");
  if (!document.fullscreenElement) {
    panel.requestFullscreen && panel.requestFullscreen();
  } else {
    document.exitFullscreen && document.exitFullscreen();
  }
}

function toggleFormFullscreen() {
  var panel = document.getElementById("formPanel");
  if (!document.fullscreenElement) {
    panel.requestFullscreen && panel.requestFullscreen();
  } else {
    document.exitFullscreen && document.exitFullscreen();
  }
}


/* ══════════════════════════════════════
   Form helpers
   ══════════════════════════════════════ */

function autoExpand(el) {
  el.style.height = "auto";
  el.style.height = el.scrollHeight + "px";
}

function closeModal() {
  document.getElementById("modal").style.display = "none";
}

function insertDataSheetValue(text) {
  insertValueAtCursor(lastFocusedField, text);
}


/* ══════════════════════════════════════
   Mail merge (shared state and row navigation)
   ══════════════════════════════════════ */

var mergeState = { headers: [], rows: [], currentIndex: -1 };

function mergeSubstitute(text) {
  if (mergeState.currentIndex < 0 || !mergeState.rows.length) return text;
  var row = mergeState.rows[mergeState.currentIndex];
  mergeState.headers.forEach(function (h) {
    var re = new RegExp(
      "\\{" + h.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\}",
      "gi",
    );
    text = text.replace(re, row[h] || "");
  });
  return text;
}

function escMerge(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function updateMergeIndicator() {
  var el = document.getElementById("mergeRowIndicator");
  if (!el) return;
  el.textContent =
    mergeState.currentIndex < 0
      ? "Showing template"
      : "Row " +
        (mergeState.currentIndex + 1) +
        " of " +
        mergeState.rows.length;
}

function nextMergeRow() {
  if (!mergeState.rows.length) return;
  var next = mergeState.currentIndex + 1;
  if (next >= mergeState.rows.length) next = 0;
  applyMergeRow(next);
}

function prevMergeRow() {
  if (!mergeState.rows.length) return;
  var prev = mergeState.currentIndex - 1;
  if (prev < 0) prev = mergeState.rows.length - 1;
  applyMergeRow(prev);
}


/* ══════════════════════════════════════
   Saved-address picker
   ══════════════════════════════════════ */

/* ── Saved-address picker (tick boxes) ──
   Same code in Letter, Notice and Proceedings. Lists the saved
   addresses (shared browser storage: ml_addresses for To,
   ml_copy_addresses for Copy To) with tick boxes. Addresses already
   in the form start ticked. On Apply, newly ticked ones fill empty
   rows or new rows, and unticked ones are removed (the first row is
   only emptied, since it holds the "+" button). Each app passes its
   own container and add-row function; see pickAddresses(). */
var AddressPicker = (function () {
  var root = null;
  var state = null;

  function readList(key) {
    try {
      var list = JSON.parse(localStorage.getItem(key));
      return Array.isArray(list) ? list : [];
    } catch (e) {
      return [];
    }
  }

  function norm(s) {
    return String(s == null ? "" : s)
      .replace(/\r\n?/g, "\n")
      .trim();
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function rowsOf(container) {
    return Array.prototype.filter.call(container.children, function (r) {
      return !!r.querySelector("textarea");
    });
  }

  function textOf(row) {
    return norm(row.querySelector("textarea").value);
  }

  function build() {
    root = document.createElement("div");
    root.className = "ap-overlay";
    root.innerHTML =
      '<div class="ap-dialog" role="dialog" aria-modal="true" aria-labelledby="apTitle">' +
      '<div class="ap-head"><h3 id="apTitle"></h3>' +
      '<button type="button" class="ap-x" title="Close" aria-label="Close">&times;</button></div>' +
      '<input type="search" class="ap-search" placeholder="Search saved addresses..." aria-label="Search saved addresses" />' +
      '<div class="ap-list" role="group" aria-labelledby="apTitle"></div>' +
      '<div class="ap-foot"><span class="ap-count" role="status"></span>' +
      '<button type="button" class="ap-btn ap-cancel">Cancel</button>' +
      '<button type="button" class="ap-btn ap-apply">Apply</button></div>' +
      "</div>";
    document.body.appendChild(root);
    root.addEventListener("click", function (e) {
      if (e.target === root) close();
    });
    root.querySelector(".ap-x").onclick = close;
    root.querySelector(".ap-cancel").onclick = close;
    root.querySelector(".ap-apply").onclick = apply;
    root.querySelector(".ap-search").oninput = filter;
    root.querySelector(".ap-list").onchange = updateCount;
    root.addEventListener("keydown", function (e) {
      if (e.key === "Escape") close();
    });
  }

  function boxes() {
    return Array.prototype.slice.call(
      root.querySelectorAll('.ap-list input[type="checkbox"]'),
    );
  }

  function updateCount() {
    var n = boxes().filter(function (b) {
      return b.checked;
    }).length;
    root.querySelector(".ap-count").textContent = state.items.length
      ? n + " selected"
      : "";
  }

  function filter() {
    var q = norm(root.querySelector(".ap-search").value).toLowerCase();
    root.querySelectorAll(".ap-item").forEach(function (el) {
      el.hidden = !!q && el.textContent.toLowerCase().indexOf(q) === -1;
    });
  }

  function open(cfg) {
    if (!root) build();
    state = cfg;
    state.items = readList(cfg.storageKey);
    var present = rowsOf(cfg.container).map(textOf);
    state.initial = {};
    var html = "";
    state.items.forEach(function (item, i) {
      var on = present.indexOf(norm(item.address)) !== -1;
      if (on) state.initial[i] = true;
      html +=
        '<label class="ap-item"><input type="checkbox" value="' +
        i +
        '"' +
        (on ? " checked" : "") +
        " /><span><strong>" +
        esc(item.title) +
        '</strong><span class="ap-addr">' +
        esc(item.address) +
        "</span></span></label>";
    });
    if (!html) {
      html =
        '<p class="ap-empty">No saved addresses yet. Use the ⚙️ button next to the heading to save some.</p>';
    }
    root.querySelector("#apTitle").textContent = cfg.title;
    root.querySelector(".ap-list").innerHTML = html;
    var search = root.querySelector(".ap-search");
    search.value = "";
    search.hidden = state.items.length < 6;
    root.querySelector(".ap-apply").disabled = !state.items.length;
    updateCount();
    root.classList.add("open");
    (search.hidden
      ? root.querySelector(".ap-list input") || root.querySelector(".ap-x")
      : search
    ).focus();
  }

  function close() {
    if (root) root.classList.remove("open");
  }

  function clearOrRemove(row, container) {
    if (rowsOf(container)[0] === row) {
      row.querySelector("textarea").value = "";
      var sel = row.querySelector("select");
      if (sel) sel.value = "";
      if (typeof autoExpand === "function")
        autoExpand(row.querySelector("textarea"));
    } else {
      row.remove();
    }
  }

  function fill(row, item) {
    var text = row.querySelector("textarea");
    text.value = item.address;
    var sel = row.querySelector("select");
    if (sel) {
      var match = Array.prototype.find.call(sel.options, function (o) {
        return o.value !== "" && o.textContent === item.title;
      });
      sel.value = match ? match.value : "";
    }
    if (typeof autoExpand === "function") autoExpand(text);
  }

  function apply() {
    var cfg = state;
    var c = cfg.container;
    var checked = {};
    boxes().forEach(function (b) {
      if (b.checked) checked[b.value] = true;
    });
    var added = 0;
    var removed = 0;

    // Unticked: take out the rows that hold exactly that address.
    Object.keys(cfg.initial).forEach(function (i) {
      if (checked[i]) return;
      var addr = norm(cfg.items[i].address);
      rowsOf(c).forEach(function (row) {
        if (textOf(row) === addr) {
          clearOrRemove(row, c);
          removed++;
        }
      });
    });

    // Ticked: fill an empty row, or add one.
    Object.keys(checked).forEach(function (i) {
      var item = cfg.items[i];
      var addr = norm(item.address);
      var rows = rowsOf(c);
      if (rows.some(function (r) { return textOf(r) === addr; })) return;
      var target = rows.filter(function (r) { return !textOf(r); })[0];
      if (!target) {
        cfg.addRow();
        rows = rowsOf(c);
        target = rows[rows.length - 1];
      }
      if (!target) return;
      fill(target, item);
      added++;
    });

    close();
    cfg.refresh();
    if (cfg.toast && (added || removed)) {
      var parts = [];
      if (added) parts.push(added + " added");
      if (removed) parts.push(removed + " removed");
      cfg.toast("Addresses: " + parts.join(", "));
    }
  }

  return { open: open, close: close };
})();

/* ════════════════════════════════════════════════════════════════
   Functions below were the same in substance in all three apps but
   differed in small app-specific details. Those details now come from
   the APP settings object that each app defines at the top of its own
   script (preview element, how to refresh the preview, how to show a
   message, where data sheets are kept).
   ════════════════════════════════════════════════════════════════ */

/* ══════════════════════════════════════
   Text formatting and numbering
   ══════════════════════════════════════ */

function applyInlineFormatting(text) {
  // Combined bold+underline, either nesting order, then standalone
  text = text.replace(/\*\*__(.+?)__\*\*/g, "<b><u>$1</u></b>");
  text = text.replace(/__\*\*(.+?)\*\*__/g, "<b><u>$1</u></b>");
  text = text.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
  text = text.replace(/__(.+?)__/g, "<u>$1</u>");
  return text;
}

function toLowerLetter(num) {
  var n = num;
  var result = "";
  while (n > 0) {
    var rem = (n - 1) % 26;
    result = String.fromCharCode(97 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result || "a";
}

function toRomanLower(num) {
  var map = [
    [10, "x"],
    [9, "ix"],
    [5, "v"],
    [4, "iv"],
    [1, "i"],
  ];
  var n = num;
  var result = "";
  for (var m = 0; m < map.length; m++) {
    while (n >= map[m][0]) {
      result += map[m][1];
      n -= map[m][0];
    }
  }
  return result || "i";
}

function insertValueAtCursor(field, text) {
  if (
    !field ||
    !document.body.contains(field) ||
    (field.tagName !== "TEXTAREA" && field.tagName !== "INPUT")
  ) {
    showToast(
      "Click into a field first, then click a value to insert it.",
    );
    return;
  }

  var start = field.selectionStart != null ? field.selectionStart : field.value.length;
  var end = field.selectionEnd != null ? field.selectionEnd : field.value.length;
  var before = field.value.substring(0, start);
  var after = field.value.substring(end);

  field.value = before + text + after;
  var newPos = start + text.length;
  field.focus();
  field.setSelectionRange(newPos, newPos);

  var evt = new Event("input", { bubbles: true });
  field.dispatchEvent(evt);
}

/* ══════════════════════════════════════
   Collapsible sections and copy options
   ══════════════════════════════════════ */

function toggleSimpleSection(key) {
  var body = document.getElementById("body-" + key);
  var btn = document.getElementById("togglebtn-" + key);
  var companion = document.getElementById(key + "Label");
  if (!body) return;
  var isOpen = body.style.display !== "none";
  body.style.display = isOpen ? "none" : "block";
  if (companion) companion.style.display = isOpen ? "none" : "flex";
  if (btn) btn.textContent = isOpen ? "▼" : "▲";
}

function toggleAllCopyOptions(checked) {
  for (var i = 0; i < COPY_OPTION_CHECKBOX_IDS.length; i++) {
    var el = document.getElementById(COPY_OPTION_CHECKBOX_IDS[i]);
    if (el) el.checked = checked;
  }
}

function updateCopySelectAllState() {
  var total = 0;
  var checkedCount = 0;
  for (var i = 0; i < COPY_OPTION_CHECKBOX_IDS.length; i++) {
    var el = document.getElementById(COPY_OPTION_CHECKBOX_IDS[i]);
    if (!el) continue;
    total++;
    if (el.checked) checkedCount++;
  }
  var selectAllEl = document.getElementById("copySelectAll");
  if (!selectAllEl) return;
  selectAllEl.checked = checkedCount === total;
  selectAllEl.indeterminate = checkedCount > 0 && checkedCount < total;
}

/* ══════════════════════════════════════
   Popup dialog
   ══════════════════════════════════════ */

function showModal(title, content) {
  document.getElementById("modalTitle").textContent = title;
  document.getElementById("modalBody").innerHTML = content;
  document.getElementById("modal").style.display = "block";
}

/* ══════════════════════════════════════
   Header with logo
   ══════════════════════════════════════ */

function buildHeaderWithLogo(textContent, mb, textAlign) {
  var mode = localStorage.getItem("headerMode_" + APP.key);
  if (mode === "none") {
    return (
      '<div style="text-align:' +
      textAlign +
      ";margin-bottom:" +
      mb +
      ';">' +
      textContent +
      "</div>"
    );
  }
  var logo = localStorage.getItem("officeLogo");
  var pos = localStorage.getItem("officeLogoPos") || "center";
  var size = localStorage.getItem("officeLogoSize") || "115";
  if (!logo) {
    return (
      '<div style="text-align:' +
      textAlign +
      ";margin-bottom:" +
      mb +
      ';">' +
      textContent +
      "</div>"
    );
  }
  var imgTag =
    '<img src="' +
    logo +
    '" style="height:' +
    size +
    'px;max-width:400px;object-fit:contain;">';
  if (pos === "center") {
    return (
      '<div style="text-align:' +
      textAlign +
      ";margin-bottom:" +
      mb +
      ';">' +
      '<div style="margin-bottom:8px;">' +
      imgTag +
      "</div>" +
      textContent +
      "</div>"
    );
  }
  var hrM = textContent.match(/<hr[^>]*>/i);
  var bodyContent = hrM ? textContent.replace(hrM[0], "") : textContent;
  var hrHtml = hrM ? hrM[0] : "";
  return (
    '<table style="width:100%;border-collapse:collapse;margin-bottom:0;padding:0;">' +
    "<tr>" +
    '<td style="border:none;padding:0 15px 0 0;vertical-align:middle;width:1%;white-space:nowrap;">' +
    imgTag +
    "</td>" +
    '<td style="border:none;padding:0;vertical-align:middle;text-align:' +
    textAlign +
    ';">' +
    bodyContent +
    "</td>" +
    "</tr></table>" +
    hrHtml +
    '<div style="margin-bottom:' +
    mb +
    ';"></div>'
  );
}

/* ══════════════════════════════════════
   Preview mode (document / data sheet) and printing
   ══════════════════════════════════════ */

function setPreviewMode(mode) {
  var letterBtn = document.getElementById("previewModeLetterBtn");
  var dataBtn = document.getElementById("previewModeDataBtn");
  var letterEl = document.querySelector(APP.previewSelector);
  var dataEl = document.getElementById("dataSheetPanel");

  var activeStyle = "background: var(--primary, #2c3e50); color: white;";
  var inactiveStyle = "background: white; color: #333;";

  if (mode === "data") {
    letterEl.style.display = "none";
    dataEl.style.display = "block";
    letterBtn.style.cssText = letterBtn.style.cssText.replace(
      /background:[^;]*;\s*color:[^;]*;/,
      inactiveStyle,
    );
    dataBtn.style.cssText = dataBtn.style.cssText.replace(
      /background:[^;]*;\s*color:[^;]*;/,
      activeStyle,
    );
  } else {
    letterEl.style.display = "";
    dataEl.style.display = "none";
    letterBtn.style.cssText = letterBtn.style.cssText.replace(
      /background:[^;]*;\s*color:[^;]*;/,
      activeStyle,
    );
    dataBtn.style.cssText = dataBtn.style.cssText.replace(
      /background:[^;]*;\s*color:[^;]*;/,
      inactiveStyle,
    );
  }
}

function printPreview() {
  var content;
  if (APP.getPrintHTML) {
    content = APP.getPrintHTML();
  } else {
    var preview = document.getElementById(APP.previewId);
    content = preview ? preview.innerHTML : "";
  }
  if (!content.trim()) {
    showToast("Nothing to print.", "error");
    return;
  }
  // Clone and strip zoom
  var tmp = document.createElement("div");
  tmp.innerHTML = content;
  tmp.querySelectorAll(".page-container").forEach(function (el) {
    el.style.zoom = "";
  });
  var font = labels[currentLang].font;
  var win = window.open("", "_blank");
  win.document.write(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Print</title><style>' +
      "* { margin: 0; padding: 0; box-sizing: border-box; }" +
      "body { font-family: " +
      font +
      "; background: white; }" +
      ".page-container { width: 100%; box-shadow: none; margin: 0; padding: 0; }" +
      "@page { size: A4; margin: 15mm 20mm; }" +
      "table { border-collapse: collapse; }" +
      (APP.printExtraCss ? APP.printExtraCss(font) : "") +
      ".draft-separator { border: none !important; background: none !important; color: transparent !important; font-size: 0 !important; padding: 0 !important; margin: 0 !important; height: 0 !important; overflow: hidden !important; }" +
      "</style></head><body>" +
      tmp.innerHTML +
      "</body></html>",
  );
  win.document.close();
  win.onload = function () {
    win.focus();
    win.print();
    win.close();
  };
}

/* ══════════════════════════════════════
   Mail merge (upload, parse, table, drafts)
   ══════════════════════════════════════ */

function toggleMailMergeControls() {
  var controls = document.getElementById("mailMergeControls");
  if (!controls) return;
  controls.style.display =
    controls.style.display === "none" ? "flex" : "none";
}

function handleMergeUpload(event) {
  var file = event.target.files[0];
  if (!file) return;
  var ext = file.name.split(".").pop().toLowerCase();
  if (ext === "csv") {
    var reader = new FileReader();
    reader.onload = function (e) {
      parseMergeCSV(e.target.result);
    };
    reader.readAsText(file);
  } else {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), {
          type: "array",
        });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (raw.length < 2) {
          showToast(
            "Table needs at least a header row and one data row.",
            "error",
          );
          return;
        }
        processMergeData(raw);
      } catch (err) {
        showToast("Could not read file.", "error");
      }
    };
    reader.readAsArrayBuffer(file);
  }
  event.target.value = "";
}

function parseMergeCSV(text) {
  // Full character-by-character parser that preserves newlines inside quoted fields
  var raw = [],
    row = [],
    field = "",
    inQuote = false;
  for (var i = 0; i < text.length; i++) {
    var c = text[i],
      next = i + 1 < text.length ? text[i + 1] : "";
    if (inQuote) {
      if (c === '"' && next === '"') {
        field += '"';
        i++;
      } // escaped quote
      else if (c === '"') {
        inQuote = false;
      } // closing quote
      else {
        field += c;
      } // preserve \n inside quotes
    } else {
      if (c === '"') {
        inQuote = true;
      } else if (c === ",") {
        row.push(field);
        field = "";
      } else if (c === "\r" && next === "\n") {
        // CRLF
        row.push(field);
        if (
          row.some(function (v) {
            return v.trim();
          })
        )
          raw.push(row);
        row = [];
        field = "";
        i++;
      } else if (c === "\n" || c === "\r") {
        // LF or CR
        row.push(field);
        if (
          row.some(function (v) {
            return v.trim();
          })
        )
          raw.push(row);
        row = [];
        field = "";
      } else {
        field += c;
      }
    }
  }
  // flush last field/row
  row.push(field);
  if (
    row.some(function (v) {
      return v.trim();
    })
  )
    raw.push(row);
  if (raw.length < 2) {
    showToast("Table needs at least a header row and one data row.", "error");
    return;
  }
  processMergeData(raw);
}

function processMergeData(raw) {
  var headers = raw[0]
    .map(function (h) {
      return String(h).trim();
    })
    .filter(Boolean);
  var rows = [];
  for (var i = 1; i < raw.length; i++) {
    var anyVal = false;
    for (var j = 0; j < headers.length; j++) {
      if (String(raw[i][j] || "").trim()) {
        anyVal = true;
        break;
      }
    }
    if (!anyVal) continue;
    var rowObj = {};
    headers.forEach(function (h, idx) {
      var _v = String(
        raw[i][idx] !== undefined ? raw[i][idx] : "",
      ).replace(/\\n/g, "\n");
      if (
        _v.charAt(0) === '"' &&
        _v.charAt(_v.length - 1) === '"' &&
        _v.length >= 2
      )
        _v = _v.slice(1, -1);
      rowObj[h] = _v;
    });
    rows.push(rowObj);
  }
  if (!rows.length) {
    showToast("No data rows found.", "error");
    return;
  }
  mergeState.headers = headers;
  mergeState.rows = rows;
  mergeState.currentIndex = -1;
  showMergePanel();
  showToast(
    "Loaded " +
      rows.length +
      " row(s) with columns: " +
      headers.join(", "),
  );
  // placeholder check and "Save every row as a draft" (templates.js)
  if (window.OfficeTemplates) OfficeTemplates.afterMergeLoad();
}

function showMergePanel() {
  document.getElementById("mergePanel").style.display = "block";
  document.getElementById("mergeClearBtn").style.display = "inline-block";
  document.getElementById("allDraftsBtn").style.display = "inline-block";
  updateMergeIndicator();
  renderMergeTable();
}

function renderMergeTable() {
  var container = document.getElementById("mergeTableContainer");
  if (!container || !mergeState.headers.length) return;
  var html = "<table><thead>";
  // Placeholder row
  html += '<tr class="merge-placeholder-row">';
  mergeState.headers.forEach(function (h) {
    html +=
      '<td title="Copy this placeholder into your ' +
APP.mergeFieldsLabel +
'">' +
      escMerge("{" + h + "}") +
      "</td>";
  });
  html += "</tr>";
  // Header row
  html += '<tr class="merge-header-row">';
  mergeState.headers.forEach(function (h) {
    html += "<td>" + escMerge(h) + "</td>";
  });
  html += "</tr></thead><tbody>";
  // Data rows
  mergeState.rows.forEach(function (row, index) {
    var active = index === mergeState.currentIndex;
    html +=
      '<tr class="merge-data-row' +
      (active ? " active" : "") +
      '" onclick="applyMergeRow(' +
      index +
      ')">';
    mergeState.headers.forEach(function (h) {
      var val = row[h] || "";
      html += "<td>" + escMerge(val) + "</td>";
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  container.innerHTML = html;
  if (mergeState.currentIndex >= 0) {
    var activeRow = container.querySelector("tr.active");
    if (activeRow) activeRow.scrollIntoView({ block: "nearest" });
  }
}

function applyMergeRow(index) {
  mergeState.currentIndex = index;
  updateMergeIndicator();
  renderMergeTable();
  APP.refreshPreview();
}

function clearMergeData() {
  mergeState.headers = [];
  mergeState.rows = [];
  mergeState.currentIndex = -1;
  document.getElementById("mergePanel").style.display = "none";
  document.getElementById("mergeClearBtn").style.display = "none";
  document.getElementById("allDraftsBtn").style.display = "none";
  if (window.OfficeTemplates) OfficeTemplates.afterMergeClear();
  APP.refreshPreview();
}

function generateAllDrafts() {
  if (!mergeState.rows.length) return;
  var preview = document.getElementById(APP.previewId);
  var savedIndex = mergeState.currentIndex;
  var parts = [];
  for (var i = 0; i < mergeState.rows.length; i++) {
    mergeState.currentIndex = i;
    APP.refreshPreview();
    var separator =
      i > 0
        ? '<div class="draft-separator" style="page-break-before:always;margin:30px 0 20px;border-top:2px dashed #bbb;text-align:center;padding-top:10px;color:#999;font-size:11px;">— Draft ' +
          (i + 1) +
          " of " +
          mergeState.rows.length +
          " —</div>"
        : "";
    parts.push(separator + preview.innerHTML);
  }
  mergeState.currentIndex = savedIndex;
  updateMergeIndicator();
  preview.innerHTML = parts.join("");
  applyZoom();
}

/* ══════════════════════════════════════
   Data sheet (Excel/CSV shown beside the preview)
   ══════════════════════════════════════ */

function handleDataSheetUpload(files) {
  if (!files || !files.length) return;
  var file = files[0];
  var isCsv = /\.csv$/i.test(file.name);
  var reader = new FileReader();

  reader.onload = function (e) {
    try {
      var workbook;
      if (isCsv) {
        workbook = XLSX.read(e.target.result, { type: "string" });
      } else {
        var arr = new Uint8Array(e.target.result);
        workbook = XLSX.read(arr, { type: "array" });
      }

      var sheets = {};
      for (var s = 0; s < workbook.SheetNames.length; s++) {
        var sheetName = workbook.SheetNames[s];
        var ws = workbook.Sheets[sheetName];
        var rows = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          defval: "",
          raw: false,
        });
        sheets[sheetName] = rows;
      }

      APP.setDataSheets(sheets);
      activeDataSheetName = workbook.SheetNames[0] || null;
      renderDataSheetTabs();
      setPreviewMode("data");
      showToast("Data sheet uploaded: " + file.name);
    } catch (err) {
      showToast("Error reading data sheet file", "error");
    }
  };

  if (isCsv) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function renderDataSheetTabs() {
  var tabsEl = document.getElementById("dataSheetTabs");
  var sheetNames = Object.keys(APP.getDataSheets() || {});

  if (sheetNames.length === 0) {
    tabsEl.innerHTML = "";
    document.getElementById("dataSheetGrid").innerHTML =
      '<div style="color:#888; font-size:12px;">No data sheet uploaded yet.</div>';
    return;
  }

  if (!activeDataSheetName || sheetNames.indexOf(activeDataSheetName) === -1) {
    activeDataSheetName = sheetNames[0];
  }

  tabsEl.innerHTML = "";
  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    var btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = name;
    btn.style.cssText =
      "padding:4px 10px;font-size:12px;border:1px solid #ccc;border-radius:4px;cursor:pointer;" +
      (name === activeDataSheetName
        ? "background: var(--primary, #2c3e50); color:white;"
        : "background:#f0f0f0; color:#333;");
    (function (sheetName) {
      btn.addEventListener("click", function () {
        activeDataSheetName = sheetName;
        var searchEl = document.getElementById("dataSheetSearch");
        if (searchEl) searchEl.value = "";
        renderDataSheetTabs();
      });
    })(name);
    tabsEl.appendChild(btn);
  }

  renderDataSheetGrid();
}

function renderDataSheetGrid(filterText) {
  var gridEl = document.getElementById("dataSheetGrid");
  var allRows = (APP.getDataSheets() && APP.getDataSheets()[activeDataSheetName]) || [];

  if (allRows.length === 0) {
    gridEl.innerHTML =
      '<div style="color:#888; font-size:12px;">This sheet has no data.</div>';
    return;
  }

  var term = (filterText || "").trim().toLowerCase();
  var rows = allRows;
  if (term) {
    rows = [allRows[0]].concat(
      allRows.slice(1).filter(function (row) {
        return row.some(function (cell) {
          return String(cell == null ? "" : cell)
            .toLowerCase()
            .indexOf(term) !== -1;
        });
      }),
    );
  }

  if (rows.length <= 1 && term) {
    gridEl.innerHTML =
      '<div style="color:#888; font-size:12px; padding-top:6px;">No matching rows.</div>';
    return;
  }

  var table = document.createElement("table");
  table.style.cssText =
    "border-collapse: collapse; width: max-content; font-size: 12px;";

  for (var r = 0; r < rows.length; r++) {
    var tr = document.createElement("tr");
    var row = rows[r];
    for (var c = 0; c < row.length; c++) {
      var td = document.createElement("td");
      var val = row[c] === null || row[c] === undefined ? "" : String(row[c]);
      td.textContent = val;
      td.style.cssText =
        "border: 1px solid #ddd; padding: 4px 8px; cursor: pointer; white-space: nowrap;" +
        (r === 0 ? " font-weight: bold; background: #f0f0f0;" : "");
      td.addEventListener("mouseenter", function () {
        this.style.background = "#dff0ff";
      });
      td.addEventListener("mouseleave", function () {
        this.style.background =
          this.parentElement.rowIndex === 0 ? "#f0f0f0" : "";
      });
      td.addEventListener("click", function () {
        insertDataSheetValue(this.textContent);
      });
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  gridEl.innerHTML = "";
  gridEl.appendChild(table);
}

/* ══════════════════════════════════════
   Saved address book (To, Copy To, and Letter's Appellate Authority)
   ══════════════════════════════════════
   The lists live in browser storage and are shared by all three apps.
   Every change re-reads the saved list first, changes it, then saves it,
   and finds the entry by its title and text (not by position). So an
   address added in one app or tab is never wiped out by another.
   Each app says through APP where it keeps its copy of the lists
   (APP.addressList) and how to reload them and its dropdowns
   (APP.reloadAddresses). */
var ADDRESS_KEYS = {
  address: "ml_addresses",
  to: "ml_addresses",
  copy: "ml_copy_addresses",
  appellate: "ml_appellate_addresses",
};
var ADDRESS_LABELS = {
  address: ["Manage Addresses", "Address"],
  to: ["Manage Addresses", "Address"],
  copy: ["Manage Copy Addresses", "Copy Address"],
  appellate: ["Manage Appellate Authority Addresses", "Appellate Authority Address"],
};

function escapeHtml(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}


// type is "address" / "to" / "copy" / "appellate" (or a storage key)
function readAddressList(type) {
  try {
    var list = JSON.parse(localStorage.getItem(ADDRESS_KEYS[type] || type));
    return Array.isArray(list) ? list : [];
  } catch (e) {
    return [];
  }
}

function managerTitle(type) {
  return (ADDRESS_LABELS[type] || ADDRESS_LABELS.address)[0];
}

// The list as this app shows it (same order as its dropdowns)
function managerItems(type) {
  return APP.addressList(type) || [];
}

function updateAddressList(type, change) {
  const items = readAddressList(type);
  change(items);
  try {
    localStorage.setItem(ADDRESS_KEYS[type], JSON.stringify(items));
  } catch (e) {
    showToast("Could not save. Browser storage may be blocked.", "error");
    return false;
  }
  APP.reloadAddresses();
  return true;
}

function openManager(type) {
  currentManager = type;
  APP.reloadAddresses();
  const items = managerItems(type);
  let html = items.length
    ? ""
    : '<p class="addr-empty">No saved addresses yet.</p>';
  items.forEach((item, i) => {
    html += `
      <div class="addr-card">
        <h4>${escapeHtml(item.title)}</h4>
        <pre>${escapeHtml(item.address)}</pre>
        <button type="button" class="btn btn-primary btn-small" onclick="editItem(${i})">Edit</button>
        <button type="button" class="btn btn-danger btn-small" onclick="deleteItem(${i})">Delete</button>
      </div>`;
  });
  html += `<button type="button" class="btn btn-small" onclick="addNewItem()">+ Add New ${(ADDRESS_LABELS[type] || ADDRESS_LABELS.address)[1]}</button>`;
  showModal(managerTitle(type), html);
}

function addNewItem() {
  showAddressForm(-1);
}

function showAddressForm(index) {
  const item = index >= 0 ? managerItems(currentManager)[index] : null;
  const html = `
    <div class="form-group">
      <label for="addrFormTitle">Title</label>
      <input type="text" id="addrFormTitle" value="${escapeHtml(item ? item.title : "")}" placeholder="Short name shown in the dropdown">
    </div>
    <div class="form-group">
      <label for="addrFormAddress">Address (press Enter for a new line)</label>
      <textarea id="addrFormAddress" rows="5">${escapeHtml(item ? item.address : "")}</textarea>
    </div>
    <div style="display: flex; gap: 8px; flex-wrap: wrap;">
      <button type="button" class="btn btn-small" onclick="saveAddressForm(${index})">Save</button>
      <button type="button" class="btn btn-primary btn-small" onclick="openManager(currentManager)">Cancel</button>
    </div>`;
  showModal(
    managerTitle(currentManager) + (item ? " — Edit" : " — Add New"),
    html,
  );
  document.getElementById("addrFormTitle").focus();
}

function saveAddressForm(index) {
  const title = document.getElementById("addrFormTitle").value.trim();
  const address = document.getElementById("addrFormAddress").value.trim();
  if (!title || !address) {
    showToast("Both title and address are required");
    return;
  }
  const old = index >= 0 ? managerItems(currentManager)[index] : null;
  if (
    updateAddressList(currentManager, (items) => {
      const at = old
        ? items.findIndex(
            (x) => x.title === old.title && x.address === old.address,
          )
        : -1;
      if (at !== -1) items[at] = { title, address };
      else items.push({ title, address });
    })
  ) {
    openManager(currentManager);
    showToast(old ? "Address updated" : "Address added");
  }
}

function editItem(index) {
  showAddressForm(index);
}

function deleteItem(index) {
  const item = managerItems(currentManager)[index];
  if (!item || !confirm(`Delete "${item.title}"?`)) return;
  if (
    updateAddressList(currentManager, (items) => {
      const at = items.findIndex(
        (x) => x.title === item.title && x.address === item.address,
      );
      if (at !== -1) items.splice(at, 1);
    })
  ) {
    openManager(currentManager);
    showToast("Address deleted");
  }
}

// Another app or tab changed the address book: refresh the lists and dropdowns here
window.addEventListener("storage", function (e) {
  if (
    e.key === null ||
    e.key === ADDRESS_KEYS.address ||
    e.key === ADDRESS_KEYS.copy ||
    e.key === ADDRESS_KEYS.appellate
  ) {
    APP.reloadAddresses();
  }
});

/* ══════════════════════════════════════
   Body text: paragraphs, numbered points (# ## ###), bullets (*)
   ══════════════════════════════════════ */
// A body line that is already HTML (an inserted table) is passed through.
// Letter/Notice table lines contain "<table"; Proceedings puts any
// substituted HTML at the start of a line (APP.bodyHtmlAtLineStart).
function isEmbeddedHtmlLine(line) {
  return APP.bodyHtmlAtLineStart
    ? line.indexOf("<") === 0
    : line.indexOf("<table") !== -1;
}

function parseBodyIntoBlocks(processedBody) {
  var lines = processedBody.split("\n");
  var blocks = [];
  var current = null;
  var level1 = 0;
  var level2 = 0;
  var level3 = 0;

  function closeCurrent() {
    if (current) blocks.push(current);
    current = null;
  }

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    if (line === "") {
      closeCurrent();
      continue;
    }

    if (isEmbeddedHtmlLine(line)) {
      closeCurrent();
      blocks.push({ type: "table", html: line });
      continue;
    }

    var markerMatch = line.match(/^(###|##|#)\s*(.*)$/);
    if (!markerMatch) {
      markerMatch = line.match(/^(\*)\s+(.*)$/);
    }
    if (markerMatch) {
      closeCurrent();
      var marker = markerMatch[1];
      var text = markerMatch[2];
      var type, label;
      if (marker === "#") {
        level1++;
        level2 = 0;
        level3 = 0;
        type = "num1";
        label = "(" + level1 + ")";
      } else if (marker === "##") {
        level2++;
        level3 = 0;
        type = "num2";
        label = "(" + toLowerLetter(level2) + ")";
      } else if (marker === "###") {
        level3++;
        type = "num3";
        label = "(" + toRomanLower(level3) + ")";
      } else {
        type = "bullet";
        label = "\u2022";
      }
      current = { type: type, label: label, lines: [text] };
      continue;
    }

    if (current) {
      current.lines.push(line);
      continue;
    }

    var noIndent = false;
    var noSpacer = false;
    var plainText = line;
    if (plainText.indexOf("~~~") === 0) {
      noIndent = true;
      noSpacer = true;
      plainText = plainText.substring(3).replace(/^\s+/, "");
    } else if (plainText.indexOf("~~") === 0) {
      noSpacer = true;
      plainText = plainText.substring(2).replace(/^\s+/, "");
    } else if (plainText.indexOf("~") === 0) {
      noIndent = true;
      plainText = plainText.substring(1).replace(/^\s+/, "");
    }

    blocks.push({
      type: "plain",
      lines: [plainText],
      noIndent: noIndent,
      noSpacer: noSpacer,
    });
  }
  closeCurrent();
  return blocks;
}

function renderBodyBlocksHTML(blocks, numbered, pStyleBase) {
  var html = "";
  // Proceedings uses <div> blocks, Letter and Notice use <p>
  var tag = APP.bodyBlockTag || "p";
  var plainCounter = 0;
  var NBSP_PER_LEVEL = 6;
  var isFirstTopLevelBlock = true;

  function repeatNbsp(count) {
    return new Array(count + 1).join("&nbsp;");
  }

  function formatLines(lines) {
    return lines
      .map(function (l) {
        return applyInlineFormatting(l);
      })
      .join("<br>");
  }

  function labelForBlock(block) {
    var levelIndex =
      block.type === "num1"
        ? 0
        : block.type === "num2"
          ? 1
          : block.type === "num3"
            ? 2
            : 0;
    var indentUnits = 1 + levelIndex;
    var nbspPrefix = repeatNbsp(indentUnits * NBSP_PER_LEVEL);
    return nbspPrefix + block.label + " " + formatLines(block.lines);
  }

  function addSpacerIfNeeded() {
    // A literal blank paragraph survives e-Office's paste-cleaning,
    // unlike margin-bottom on <p> tags, which can get stripped.
    // Sized to ~1em (16px) to match the gap e-Office falls back to
    // when it strips a margin declaration entirely.
    if (!isFirstTopLevelBlock) {
      html +=
        "<" +
        tag +
        ' style="margin: 0; padding: 0; font-size: 16px; line-height: 1;">&nbsp;</' +
        tag +
        ">";
    }
    isFirstTopLevelBlock = false;
  }

  var i = 0;
  while (i < blocks.length) {
    var block = blocks[i];

    if (block.type === "table") {
      html += block.html;
      isFirstTopLevelBlock = false;
      i++;
      continue;
    }

    if (block.type === "plain") {
      if (!block.noSpacer) {
        addSpacerIfNeeded();
      } else {
        isFirstTopLevelBlock = false;
      }
      var prefix = "";
      if (numbered) {
        plainCounter++;
        prefix = plainCounter + ". ";
      }
      var extraStyle = block.noIndent ? " text-indent: 0;" : "";
      html +=
        "<" +
        tag +
        ' style="' +
        pStyleBase +
        extraStyle +
        '">' +
        prefix +
        formatLines(block.lines) +
        "</" +
      tag +
      ">";
      i++;
      continue;
    }

    // Marker block (num1/num2/num3/bullet): group every consecutive
    // marker block into ONE <p>, joined with <br>, so spacing between
    // points is governed by line-height instead of per-paragraph margin.
    addSpacerIfNeeded();
    var groupLines = [];
    while (
      i < blocks.length &&
      blocks[i].type !== "plain" &&
      blocks[i].type !== "table"
    ) {
      groupLines.push(labelForBlock(blocks[i]));
      i++;
    }
    html +=
      "<" +
      tag +
      ' style="' +
      pStyleBase +
      ' text-indent: 0;">' +
      groupLines.join("<br>") +
      "</" +
      tag +
      ">";
  }

  return html;
}

/* ══════════════════════════════════════
   Tables from Excel (all three apps build the HTML here)
   ══════════════════════════════════════ */
// Table HTML for the preview and the e-Office copy. The outline stops
// e-Office from dropping the outer border. splitAfterRows starts a new
// table (with the header row repeated) after those row numbers.
function buildTableHTMLFromData(cellData, splitAfterRows) {
  if (!cellData || cellData.length === 0) return "";

  var tableOpenTag =
    '<table style="width: 588px; max-width: 588px; border-collapse: collapse; margin: 0 auto; border: 1px solid #000; outline: 1px solid #000; outline-offset: -1px; font-family: mandaram, \'Noto Sans Malayalam\', \'Malayalam Sangam MN\', serif; margin-left: auto; margin-right: auto; display: table;">';
  var tableOpenTagSplit =
    '<table style="width: 588px; max-width: 588px; border-collapse: collapse; margin: 24px auto 0 auto; border: 1px solid #000; outline: 1px solid #000; outline-offset: -1px; font-family: mandaram, \'Noto Sans Malayalam\', \'Malayalam Sangam MN\', serif; margin-left: auto; margin-right: auto; display: table; page-break-before: always;">';
  var splitDividerHTML =
    '<p style="text-align: center; font-weight: bold; margin: 16px 0; font-family: mandaram, \'Noto Sans Malayalam\', \'Malayalam Sangam MN\', serif;">(Table continued)</p>';
  // Proceedings centres the table in a wrapper block
  var cleanedHTML =
    (APP.tableWrapper
      ? '<div style="text-align: center; margin: 20px auto; width: 100%;">'
      : "") + tableOpenTag;

  var splitRows = [];
  if (Array.isArray(splitAfterRows)) {
    for (var s = 0; s < splitAfterRows.length; s++) {
      var n = parseInt(splitAfterRows[s], 10);
      if (n > 0 && splitRows.indexOf(n) === -1) splitRows.push(n);
    }
  } else {
    var singleN = parseInt(splitAfterRows, 10);
    if (singleN > 0) splitRows.push(singleN);
  }
  splitRows.sort(function (a, b) {
    return a - b;
  });
  var headerRowHTML = "";

  var numCols = 0;
  for (var r0 = 0; r0 < cellData.length; r0++) {
    if (cellData[r0] && cellData[r0].length > 0) {
      numCols = cellData[r0].length;
      break;
    }
  }
  var TOTAL_TABLE_WIDTH = 588;
  var colWidths = [];
  if (numCols > 0) {
    var baseWidth = Math.floor(TOTAL_TABLE_WIDTH / numCols);
    var remainder = TOTAL_TABLE_WIDTH - baseWidth * numCols;
    for (var c1 = 0; c1 < numCols; c1++) {
      colWidths.push(baseWidth + (c1 < remainder ? 1 : 0));
    }
  }

  for (var i = 0; i < cellData.length; i++) {
    var cells = cellData[i];
    if (!cells || cells.length === 0) continue;

    if (splitRows.indexOf(i) !== -1) {
      cleanedHTML +=
        "</table>" +
        splitDividerHTML +
        tableOpenTagSplit +
        headerRowHTML;
    }

    cleanedHTML += "<tr>";
    var rowHTML = "<tr>";

    for (var j = 0; j < cells.length; j++) {
      // keep a 0 in a cell (it used to show as blank)
      var cleanText =
        cells[j] === null || cells[j] === undefined ? "" : String(cells[j]).trim();

      var colWidthPx =
        colWidths.length > j
          ? colWidths[j]
          : Math.floor(588 / cells.length);
      var cellStyle =
        "border: 1px solid #000; padding: 2px; text-align: center; width: " +
        colWidthPx +
        "px;";
      var cellAttrs = ' style="' + cellStyle;

      if (i === 0) {
        cellAttrs += " background-color: #f0f0f0;";
      }

      cellAttrs += '"';

      var spanStyle =
        "font-family: mandaram, 'Noto Sans Malayalam', 'Malayalam Sangam MN', serif; font-size: 13pt;" +
        (i === 0 ? " font-weight: bold;" : "");
      var cellText =
        '<span style="' + spanStyle + '">' + cleanText + "</span>";

      var cellHTML = "<td" + cellAttrs + ">" + cellText + "</td>";
      cleanedHTML += cellHTML;
      rowHTML += cellHTML;
    }

    cleanedHTML += "</tr>";
    rowHTML += "</tr>";

    if (i === 0) {
      headerRowHTML = rowHTML;
    }
  }

  cleanedHTML += "</table>" + (APP.tableWrapper ? "</div>" : "");
  return cleanedHTML;
}

/* Letter and Notice table editor ({{TABLE_n}} placeholders, data in
   APP.getTables()). Proceedings has its own table system and defines its
   own insertTablePlaceholder, renderEditableTableGrid and applyTableSplit,
   which replace these when its script loads. */
function getTableCellData(tbl) {
  if (!tbl) return null;
  if (tbl.cellData) return tbl.cellData;
  if (tbl.rawHTML) {
    tbl.cellData = extractCellDataFromRawHTML(tbl.rawHTML);
    return tbl.cellData;
  }
  return null;
}

function extractCellDataFromRawHTML(rawHTML) {
  var tempDiv = document.createElement("div");
  tempDiv.innerHTML = rawHTML;

  var table = tempDiv.querySelector("table");
  if (!table) return [];

  var rows = table.querySelectorAll("tr");
  var cellData = [];

  for (var i = 0; i < rows.length; i++) {
    var cells = rows[i].querySelectorAll("td, th");
    if (cells.length === 0) continue;

    var rowData = [];
    for (var j = 0; j < cells.length; j++) {
      var cleanText = cells[j].textContent || cells[j].innerText || "";
      cleanText = cleanText.replace(/style="[^"]*"/g, "");
      cleanText = cleanText.replace(/border:\s*[^;]*;/g, "");
      cleanText = cleanText.replace(/padding:\s*[^;]*;/g, "");
      cleanText = cleanText.replace(/background-color:\s*[^;]*;/g, "");
      cleanText = cleanText.replace(/font-weight:\s*[^;]*;/g, "");
      cleanText = cleanText.replace(/text-align:\s*[^;]*;/g, "");
      cleanText = cleanText.replace(/white-space:\s*[^;]*;/g, "");
      rowData.push(cleanText.trim());
    }
    cellData.push(rowData);
  }

  return cellData;
}

function regenerateTableHTML(tableId) {
  var tbl = APP.getTables()[tableId];
  var cellData = getTableCellData(tbl);
  if (!tbl || !cellData) return;
  tbl.html = buildTableHTMLFromData(cellData, tbl.splitAfterRows || []);
  APP.refreshPreview();
}

function applyTableSplit(tableId, rowNums) {
  var tbl = APP.getTables()[tableId];
  var cellData = getTableCellData(tbl);
  if (!tbl || !cellData) {
    showToast("Cannot split this table \u2014 original data not found.");
    return;
  }
  tbl.splitAfterRows = rowNums;
  tbl.html = buildTableHTMLFromData(cellData, rowNums);
  APP.refreshPreview();
  showToast(
    rowNums.length > 0
      ? "Table split after row(s) " + rowNums.join(", ") + "."
      : "Table splits removed.",
  );
}

function insertTablePlaceholder(tableId) {
  var letterBody = document.getElementById("letterBody");
  var cursorPos = letterBody.selectionStart || letterBody.value.length;
  var placeholder = "{{" + tableId + "}}";
  var before = letterBody.value.substring(0, cursorPos);
  var after = letterBody.value.substring(cursorPos);
  letterBody.value = before + placeholder + after;
  letterBody.focus();
  letterBody.setSelectionRange(
    cursorPos + placeholder.length,
    cursorPos + placeholder.length,
  );
  APP.refreshPreview();
}

function renderEditableTableGrid(tableId) {
  var gridEl = document.getElementById("tableEditGrid_" + tableId);
  if (!gridEl) return;
  var tbl = APP.getTables()[tableId];
  var cellData = getTableCellData(tbl);
  if (!cellData) {
    gridEl.innerHTML =
      '<div style="color:#888; font-size:12px;">No editable data found for this table.</div>';
    return;
  }

  var table = document.createElement("table");
  table.style.cssText =
    "border-collapse: collapse; width: max-content; font-size: 12px;";

  for (var r = 0; r < cellData.length; r++) {
    var tr = document.createElement("tr");
    for (var c = 0; c < cellData[r].length; c++) {
      var td = document.createElement("td");
      td.style.cssText = "border: 1px solid #ddd; padding: 2px;";
      var input = document.createElement("input");
      input.type = "text";
      input.value = cellData[r][c];
      input.style.cssText =
        "width: 110px; padding: 3px 4px; font-size: 12px; border: 1px solid #ccc; border-radius: 3px; box-sizing: border-box;" +
        (r === 0 ? " font-weight: bold; background: #f7f7f7;" : "");
      (function (row, col) {
        input.addEventListener("input", function () {
          APP.getTables()[tableId].cellData[row][col] = this.value;
          regenerateTableHTML(tableId);
        });
      })(r, c);
      td.appendChild(input);
      tr.appendChild(td);
    }

    var delTd = document.createElement("td");
    delTd.style.cssText = "border: none; padding: 2px;";
    var delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.textContent = "🗑";
    delBtn.title = "Delete row";
    delBtn.style.cssText =
      "cursor: pointer; background: #e74c3c; color: white; border: none; border-radius: 3px; padding: 2px 6px; font-size: 11px;";
    (function (row) {
      delBtn.addEventListener("click", function () {
        if (APP.getTables()[tableId].cellData.length <= 1) {
          showToast("Table must have at least one row.");
          return;
        }
        APP.getTables()[tableId].cellData.splice(row, 1);
        regenerateTableHTML(tableId);
        renderEditableTableGrid(tableId);
      });
    })(r);
    delTd.appendChild(delBtn);
    tr.appendChild(delTd);

    table.appendChild(tr);
  }

  gridEl.innerHTML = "";
  gridEl.appendChild(table);

  var addRowBtn = document.createElement("button");
  addRowBtn.type = "button";
  addRowBtn.textContent = "+ Add Row";
  addRowBtn.style.cssText =
    "margin-top: 6px; cursor: pointer; background: #2e7d32; color: white; border: none; border-radius: 3px; padding: 4px 10px; font-size: 11px;";
  addRowBtn.addEventListener("click", function () {
    var numCols =
      APP.getTables()[tableId].cellData[0]
        ? APP.getTables()[tableId].cellData[0].length
        : 1;
    var newRow = [];
    for (var k = 0; k < numCols; k++) newRow.push("");
    APP.getTables()[tableId].cellData.push(newRow);
    regenerateTableHTML(tableId);
    renderEditableTableGrid(tableId);
  });
  gridEl.appendChild(addRowBtn);
}

/* ══════════════════════════════════════
   Header & logo settings changed elsewhere
   ══════════════════════════════════════
   The logo, header image and header choice are set on the home page. If
   they change while an app is open (another tab), or the app page comes
   back from the browser's Back/Forward cache, redraw the preview so it
   shows the current setting. */
(function () {
  var KEYS = ["officeLogo", "officeLogoPos", "officeLogoSize", "officeHeader"];
  function redraw() {
    if (typeof APP === "undefined" || !APP.refreshPreview) return;
    try {
      APP.refreshPreview();
    } catch (e) {}
  }
  window.addEventListener("storage", function (e) {
    if (
      e.key === null ||
      KEYS.indexOf(e.key) !== -1 ||
      (typeof APP !== "undefined" && e.key === "headerMode_" + APP.key)
    ) {
      redraw();
    }
  });
  window.addEventListener("pageshow", function (e) {
    if (e.persisted) redraw();
  });
})();
