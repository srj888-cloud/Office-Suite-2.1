/* e-Office Suite - templates and mail-merge tools.

   A template is a reusable starting point: the whole form (as a draft
   holds it) saved under a name, usually written with {placeholders} that
   match the column headers of a data sheet. Templates are stored with the
   drafts (common/drafts.js, store "templates").

   In each app this adds, in place of the old Template dropdown:
     Start from …        opens a NEW draft filled from the template
     Save as template    saves the current form (asks for a name)
     Templates           use / edit / update / delete, download / upload
     Upload file         the app's own JSON file upload, as before
   and in the Mail Merge panel:
     a check of the form's {placeholders} against the sheet's columns,
     Save every row as a draft,
     Open data sheet (the Google Sheet saved with the template; download it
     as CSV there and upload it with Mail Merge).

   Placeholders match column headers ignoring capital letters
   ({name} = "Name"); spelling and spaces must match. Near misses are
   pointed out by the check.
*/
var OfficeTemplates = (function () {
  var D = window.OfficeDrafts;
  var IGNORE = { collegename: 1 }; // Proceedings fills {collegeName} itself
  var SKIP_KEYS = { dataSheets: 1, copyOptions: 1 };

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /* ── placeholders ── */
  // {Name} in any text of the form (not {{TABLE_1}}), in order, no repeats
  function placeholders(data) {
    var seen = {};
    var out = [];
    (function walk(v, key) {
      if (key && SKIP_KEYS[key]) return;
      if (typeof v === "string") {
        var re = /(^|[^{])\{([^{}\n]{1,60})\}(?!\})/g;
        var m;
        while ((m = re.exec(v))) {
          var name = m[2].trim();
          var k = name.toLowerCase();
          if (!name || IGNORE[k] || seen[k]) continue;
          seen[k] = 1;
          out.push(name);
        }
      } else if (Array.isArray(v)) {
        v.forEach(function (x) {
          walk(x);
        });
      } else if (v && typeof v === "object") {
        Object.keys(v).forEach(function (k) {
          walk(v[k], k);
        });
      }
    })(data);
    return out;
  }

  // the form's data with every {placeholder} filled from the current merge row
  function fillDeep(v, key) {
    if (key && SKIP_KEYS[key]) return v;
    if (typeof v === "string") return mergeSubstitute(v);
    if (Array.isArray(v)) return v.map(function (x) {
      return fillDeep(x);
    });
    if (v && typeof v === "object") {
      var o = {};
      Object.keys(v).forEach(function (k) {
        o[k] = fillDeep(v[k], k);
      });
      return o;
    }
    return v;
  }

  function loose(s) {
    return String(s).toLowerCase().replace(/[^a-z0-9ഀ-ൿ]/g, "");
  }
  function distance(a, b) {
    a = a.toLowerCase();
    b = b.toLowerCase();
    var prev = [];
    for (var j = 0; j <= b.length; j++) prev[j] = j;
    for (var i = 1; i <= a.length; i++) {
      var cur = [i];
      for (var k = 1; k <= b.length; k++) {
        cur[k] = Math.min(prev[k] + 1, cur[k - 1] + 1, prev[k - 1] + (a[i - 1] === b[k - 1] ? 0 : 1));
      }
      prev = cur;
    }
    return prev[b.length];
  }
  function closest(name, headers) {
    var l = loose(name);
    for (var i = 0; i < headers.length; i++) if (loose(headers[i]) === l) return headers[i];
    var best = null;
    var bestD = 3;
    headers.forEach(function (h) {
      var d = distance(name, h);
      if (d < bestD) {
        bestD = d;
        best = h;
      }
    });
    return best;
  }

  /* ── styles ── */
  var CSS =
    ".tpl-bar{margin-bottom:4px}" +
    ".tpl-bar>label{display:block;font-weight:600;margin-bottom:6px}" +
    ".tpl-row{display:flex;flex-wrap:wrap;gap:6px;align-items:center}" +
    ".tpl-row select{flex:1;min-width:180px}" +
    ".tpl-row button{white-space:nowrap}" +
    ".tpl-info{display:block;margin-top:6px;font-size:12px;color:#5d6d7e;line-height:1.5}" +
    ".tpl-chip{display:inline-block;background:#eef4fb;color:#2c3e50;border-radius:10px;padding:0 7px;margin:1px 2px;font-size:11.5px;font-family:inherit}" +
    ".tpl-chip.bad{background:#fdecea;color:#c0392b}" +
    ".tpl-list{list-style:none;margin:10px 0 0;padding:0;max-height:52vh;overflow-y:auto}" +
    ".tpl-item{border:1px solid #e3e9ed;border-radius:10px;padding:10px 12px;margin-bottom:8px;background:#fbfcfd}" +
    ".tpl-item strong{color:#2c3e50}" +
    ".tpl-item small{display:block;color:#7f8c8d;font-size:12px;margin-top:2px;word-break:break-all}" +
    ".tpl-btns{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}" +
    ".tpl-form label{display:block;font-size:12.5px;font-weight:600;color:#5d6d7e;margin:10px 0 4px}" +
    ".tpl-form input{width:100%;padding:8px 10px;border:1px solid #cfd8dc;border-radius:7px;font-size:14px;font-family:inherit}" +
    ".tpl-form p{font-size:12.5px;color:#7f8c8d;margin-top:6px;line-height:1.5}" +
    ".tpl-foot{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px;padding-top:12px;border-top:1px solid #e3e9ed}" +
    ".merge-check{margin:8px 0;padding:9px 11px;border-radius:8px;background:#f4f8fb;font-size:12.5px;line-height:1.55;color:#2c3e50}" +
    ".merge-check.warn{background:#fff7e6}" +
    ".merge-check .ok{color:#1e8449;font-weight:600}" +
    ".merge-check .bad{color:#c0392b;font-weight:600}" +
    ".merge-check .muted{color:#7f8c8d}" +
    ".merge-tools{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px}" +
    ".sheet-row{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-top:8px}" +
    ".sheet-row small{flex-basis:100%;color:#7f8c8d;font-size:11.5px}";
  function addStyles() {
    if (document.getElementById("tplStyles")) return;
    var st = document.createElement("style");
    st.id = "tplStyles";
    st.textContent = CSS;
    document.head.appendChild(st);
  }

  /* ── the template bar in the form ── */
  var oldGroup = null;

  function renderBar() {
    var bar = document.getElementById("tplBar");
    if (!bar) return Promise.resolve();
    return D.tplList(APP.draftApp).then(function (list) {
      var sel = document.getElementById("tplStart");
      sel.innerHTML =
        '<option value="">' + (list.length ? "Start a new draft from a template…" : "No templates saved yet") + "</option>" +
        list
          .map(function (t) {
            return '<option value="' + esc(t.id) + '">' + esc(t.name) + "</option>";
          })
          .join("");
      sel.disabled = !list.length;
      updateInfo();
    }, function () {});
  }

  function updateInfo() {
    var info = document.getElementById("tplInfo");
    if (!info) return;
    var o = D.getOrigin();
    var ph = [];
    try {
      ph = placeholders(APP.getDraftData());
    } catch (e) {}
    var parts = [];
    if (o && o.templateId) parts.push("From template <b>" + esc(o.templateName) + "</b>.");
    if (ph.length) {
      parts.push(
        "Placeholders: " +
          ph.map(function (p) {
            return '<span class="tpl-chip">{' + esc(p) + "}</span>";
          }).join("") +
          " — fill them with 📊 Mail Merge.",
      );
    }
    info.innerHTML = parts.join(" ");
    info.style.display = parts.length ? "" : "none";
  }

  function addBar() {
    var anchor = document.getElementById("templateSelect") || document.getElementById("documentType");
    oldGroup = anchor && anchor.closest(".form-group");
    if (!oldGroup || document.getElementById("tplBar")) return;
    Array.prototype.forEach.call(oldGroup.children, function (c) {
      c.style.display = "none";
    });
    var bar = document.createElement("div");
    bar.className = "tpl-bar";
    bar.id = "tplBar";
    bar.innerHTML =
      "<label for=\"tplStart\">Template</label>" +
      '<div class="tpl-row">' +
      '<select id="tplStart"></select>' +
      '<button type="button" class="btn btn-small" data-tpl="save" title="Save the form as a reusable template">💾 Save as template</button>' +
      '<button type="button" class="btn btn-small" data-tpl="manage" title="Use, rename, update or delete templates">📑 Templates</button>' +
      '<button type="button" class="btn btn-import btn-small" data-tpl="upload" title="Load a draft or template file (.json) saved earlier">📂 Upload file</button>' +
      "</div>" +
      '<small class="tpl-info" id="tplInfo" style="display:none"></small>';
    oldGroup.insertBefore(bar, oldGroup.firstChild);
    bar.style.display = "";
    bar.addEventListener("click", function (e) {
      var b = e.target.closest("[data-tpl]");
      if (!b) return;
      var act = b.getAttribute("data-tpl");
      if (act === "save") saveDialog();
      if (act === "manage") showWindow();
      if (act === "upload") {
        var input = oldGroup.querySelector('input[type="file"]');
        if (input) input.click();
      }
    });
    document.getElementById("tplStart").addEventListener("change", function () {
      if (this.value) D.openTemplate(this.value);
    });
    // placeholders listed under the bar follow the form
    var t = null;
    document.addEventListener("input", function (e) {
      if (!e.target.closest || !e.target.closest("#formPanel")) return;
      clearTimeout(t);
      t = setTimeout(updateInfo, 600);
    }, true);
    renderBar();
  }

  /* ── Save as template ── */
  function saveDialog() {
    var data = APP.getDraftData();
    var o = D.getOrigin() || {};
    var ph = placeholders(data);
    var defName = o.templateName || (APP.draftTitle(data) || "").trim();
    showModal(
      "Save as template",
      '<div class="tpl-form">' +
        '<label for="tplName">Template name</label>' +
        '<input id="tplName" value="' + esc(defName) + '" placeholder="e.g. Guest lecturer salary">' +
        '<label for="tplSheet">Google Sheet link (optional)</label>' +
        '<input id="tplSheet" value="' + esc(o.sheetUrl || "") + '" placeholder="Link to the sheet tab with the data for this template">' +
        "<p>" +
        (ph.length
          ? "Placeholders in this form: " + ph.map(function (p) {
              return '<span class="tpl-chip">{' + esc(p) + "}</span>";
            }).join("") + ". The sheet needs a column with each of these names."
          : "Tip: write <b>{Column name}</b> where a value from the data sheet should go, e.g. {Name} or {Amount}.") +
        "</p>" +
        '<div class="tpl-btns"><button type="button" class="btn btn-success btn-small" id="tplSaveBtn">Save template</button>' +
        '<button type="button" class="btn btn-small" onclick="closeModal()">Cancel</button></div></div>',
    );
    var nameEl = document.getElementById("tplName");
    nameEl.focus();
    nameEl.select();
    document.getElementById("tplSaveBtn").onclick = function () {
      var name = nameEl.value.trim();
      var sheet = document.getElementById("tplSheet").value.trim();
      if (!name) {
        showToast("Enter a name for the template.", "error");
        nameEl.focus();
        return;
      }
      D.tplList(APP.draftApp).then(function (list) {
        var same = list.filter(function (t) {
          return t.name.toLowerCase() === name.toLowerCase();
        })[0];
        if (same && !confirm('A template called "' + same.name + '" already exists. Replace it with this form?')) return;
        var now = Date.now();
        var tpl = same
          ? Object.assign({}, same, { name: name, data: data, sheetUrl: sheet, updated: now })
          : { id: D.newId(), app: APP.draftApp, name: name, data: data, sheetUrl: sheet, created: now, updated: now };
        return D.tplPut(tpl).then(function () {
          closeModal();
          showToast((same ? "Template updated: " : "Template saved: ") + name);
          renderBar();
          refreshSheetRow();
        });
      }).catch(function (err) {
        showToast("Could not save the template. " + (err && err.message ? err.message : ""), "error");
      });
    };
  }

  /* ── Templates window ── */
  var editingId = null;

  function showWindow() {
    D.tplList(APP.draftApp).then(function (list) {
      var html =
        '<p class="draft-hint" style="margin:0">Choosing <b>Use</b> starts a new draft from the template; your current draft stays saved.</p>';
      if (!list.length) {
        html += '<p class="draft-empty">No templates yet. Fill in the form (with {placeholders} where data goes) and click 💾 Save as template.</p>';
      } else {
        html += '<ul class="tpl-list">';
        list.forEach(function (t) {
          var ph = placeholders(t.data);
          if (editingId === t.id) {
            html +=
              '<li class="tpl-item tpl-form" data-id="' + esc(t.id) + '">' +
              '<label>Template name</label><input data-f="name" value="' + esc(t.name) + '">' +
              '<label>Google Sheet link (optional)</label><input data-f="sheet" value="' + esc(t.sheetUrl || "") + '">' +
              '<div class="tpl-btns"><button type="button" class="btn btn-success btn-small" data-act="saveedit">Save</button>' +
              '<button type="button" class="btn btn-small" data-act="canceledit">Cancel</button></div></li>';
            return;
          }
          html +=
            '<li class="tpl-item" data-id="' + esc(t.id) + '"><strong>' + esc(t.name) + "</strong>" +
            "<small>" +
            (ph.length ? ph.map(function (p) {
              return '<span class="tpl-chip">{' + esc(p) + "}</span>";
            }).join("") : "No placeholders") +
            "</small>" +
            (t.sheetUrl ? "<small>📊 " + esc(t.sheetUrl) + "</small>" : "") +
            "<small>Saved " + esc(D.when(t.updated)) + "</small>" +
            '<div class="tpl-btns">' +
            '<button type="button" class="btn btn-primary btn-small" data-act="use">Use</button>' +
            '<button type="button" class="btn btn-small" data-act="edit">Rename / link</button>' +
            '<button type="button" class="btn btn-small" data-act="update" title="Replace the saved content with what is in the form now">Update with current form</button>' +
            '<button type="button" class="btn btn-danger btn-small" data-act="delete">Delete</button>' +
            "</div></li>";
        });
        html += "</ul>";
      }
      html +=
        '<div class="tpl-foot">' +
        '<button type="button" class="btn btn-small" data-act="download"' + (list.length ? "" : " disabled") + ">Download templates</button>" +
        '<button type="button" class="btn btn-small" data-act="upload">Upload templates</button>' +
        '<input type="file" id="tplUploadInput" accept=".json,application/json" hidden>' +
        '<span class="draft-hint">A file with this app\'s templates, to share or move to another computer. Templates are also in “Back up drafts &amp; templates”.</span>' +
        "</div>";
      showModal("Templates · " + (D.APP_NAMES[APP.draftApp] || ""), '<div id="tplWindow">' + html + "</div>");
      var root = document.getElementById("tplWindow");
      root.onclick = onWindowClick;
      document.getElementById("tplUploadInput").onchange = uploadTemplates;
    });
  }

  function onWindowClick(e) {
    var b = e.target.closest("[data-act]");
    if (!b) return;
    var act = b.getAttribute("data-act");
    var li = b.closest("[data-id]");
    var id = li && li.getAttribute("data-id");
    if (act === "use") {
      closeModal();
      D.openTemplate(id);
    } else if (act === "edit") {
      editingId = id;
      showWindow();
    } else if (act === "canceledit") {
      editingId = null;
      showWindow();
    } else if (act === "saveedit") {
      var name = li.querySelector('[data-f="name"]').value.trim();
      var sheet = li.querySelector('[data-f="sheet"]').value.trim();
      if (!name) {
        showToast("Enter a name for the template.", "error");
        return;
      }
      D.tplGet(id).then(function (t) {
        if (!t) return;
        t.name = name;
        t.sheetUrl = sheet;
        t.updated = Date.now();
        return D.tplPut(t);
      }).then(function () {
        editingId = null;
        showWindow();
        renderBar();
        refreshSheetRow();
      });
    } else if (act === "update") {
      D.tplGet(id).then(function (t) {
        if (!t || !confirm('Replace the saved content of "' + t.name + '" with what is in the form now?')) return;
        t.data = APP.getDraftData();
        t.updated = Date.now();
        return D.tplPut(t).then(function () {
          showToast("Template updated: " + t.name);
          showWindow();
        });
      });
    } else if (act === "delete") {
      D.tplGet(id).then(function (t) {
        if (!t || !confirm('Delete the template "' + t.name + '"? Drafts made from it are kept. This cannot be undone.')) return;
        return D.tplRemove(id).then(function () {
          showToast("Template deleted");
          showWindow();
          renderBar();
        });
      });
    } else if (act === "download") {
      downloadTemplates();
    } else if (act === "upload") {
      document.getElementById("tplUploadInput").click();
    }
  }

  function downloadTemplates() {
    D.tplList(APP.draftApp).then(function (list) {
      var out = { type: "office-suite-templates", version: 1, app: APP.draftApp, exported: new Date().toISOString(), templates: list };
      var blob = new Blob([JSON.stringify(out, null, 2)], { type: "application/json" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = "office_suite_templates_" + APP.draftApp + ".json";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast("Downloaded " + list.length + " template(s).");
    });
  }

  function uploadTemplates() {
    var input = this;
    var file = input.files && input.files[0];
    input.value = "";
    if (!file) return;
    var r = new FileReader();
    r.onload = function () {
      var json;
      try {
        json = JSON.parse(r.result);
      } catch (e) {
        showToast("Could not read the file. Choose a templates file.", "error");
        return;
      }
      if (!json || json.type !== "office-suite-templates" || !Array.isArray(json.templates)) {
        showToast("This is not a templates file. (A single draft file goes in 📂 Upload file.)", "error");
        return;
      }
      var added = 0;
      var skipped = 0;
      Promise.all(
        json.templates
          .filter(function (t) {
            return t && t.id && t.name && t.data && D.APP_FILES[t.app];
          })
          .map(function (t) {
            return D.tplGet(t.id).then(function (have) {
              if (have && (have.updated || 0) >= (t.updated || 0)) {
                skipped++;
                return;
              }
              added++;
              return D.tplPut(t);
            });
          }),
      ).then(function () {
        showToast("Templates added or updated: " + added + (skipped ? " (" + skipped + " already up to date)" : "") + ".");
        showWindow();
        renderBar();
      });
    };
    r.readAsText(file);
  }

  /* ── mail merge: check, save rows, Google Sheet ── */
  function mergePanelBox() {
    var panel = document.getElementById("mergePanel");
    if (!panel) return null;
    var box = document.getElementById("mergeCheck");
    if (!box) {
      box = document.createElement("div");
      box.id = "mergeCheck";
      box.className = "merge-check";
      panel.insertBefore(box, panel.firstChild);
    }
    return box;
  }

  function afterMergeLoad() {
    var box = mergePanelBox();
    if (!box) return;
    var headers = mergeState.headers || [];
    var ph = [];
    try {
      ph = placeholders(APP.getDraftData());
    } catch (e) {}
    var lower = headers.map(function (h) {
      return h.toLowerCase().trim();
    });
    var missing = ph.filter(function (p) {
      return lower.indexOf(p.toLowerCase().trim()) === -1;
    });
    var used = {};
    ph.forEach(function (p) {
      used[p.toLowerCase().trim()] = 1;
    });
    var unused = headers.filter(function (h) {
      return !used[h.toLowerCase().trim()];
    });
    var html = "";
    if (!ph.length) {
      html +=
        '<span class="muted">This draft has no {placeholders} yet. Write a column name in braces where its value should go: </span>' +
        headers.map(function (h) {
          return '<span class="tpl-chip">{' + esc(h) + "}</span>";
        }).join("");
    } else if (!missing.length) {
      html += '<span class="ok">✓ All ' + ph.length + " placeholder(s) match columns in the sheet.</span>";
    } else {
      html += '<span class="bad">⚠ ' + missing.length + " placeholder(s) have no matching column and will stay as typed:</span><br>";
      html += missing.map(function (p) {
        var c = closest(p, headers);
        return '<span class="tpl-chip bad">{' + esc(p) + "}</span>" + (c ? ' <span class="muted">— did you mean “' + esc(c) + "”?</span>" : "");
      }).join("<br>");
    }
    if (ph.length && unused.length) {
      html += '<br><span class="muted">Columns not used: ' + unused.map(esc).join(", ") + "</span>";
    }
    html +=
      '<div class="merge-tools"><button type="button" class="btn btn-small" id="mergeSaveRows">💾 Save every row as a draft (' +
      mergeState.rows.length + ")</button></div>";
    box.innerHTML = html;
    box.className = "merge-check" + (missing.length ? " warn" : "");
    box.style.display = "";
    document.getElementById("mergeSaveRows").onclick = saveRows;
  }

  function afterMergeClear() {
    var box = document.getElementById("mergeCheck");
    if (box) box.style.display = "none";
  }

  function saveRows() {
    var rows = mergeState.rows || [];
    if (!rows.length) return;
    if (!confirm("Save " + rows.length + " separate drafts, one for each row of the sheet, with the placeholders filled in?")) return;
    var raw = APP.getDraftData();
    var keep = mergeState.currentIndex;
    var jobs = [];
    var firstCol = mergeState.headers[0];
    for (var i = 0; i < rows.length; i++) {
      mergeState.currentIndex = i;
      var filled = fillDeep(raw);
      var label = String(rows[i][firstCol] || "Row " + (i + 1)).trim();
      var title = (APP.draftTitle(filled) || "Untitled").trim();
      jobs.push(D.saveNew(filled, title + " – " + label));
    }
    mergeState.currentIndex = keep;
    Promise.all(jobs).then(
      function () {
        showToast("Saved " + rows.length + " drafts. Open 📂 Drafts to see them.");
      },
      function (err) {
        showToast("Could not save the drafts. " + (err && err.message ? err.message : ""), "error");
      },
    );
  }

  function sheetUrlNow() {
    var o = D.getOrigin();
    return o && o.sheetUrl ? o.sheetUrl : "";
  }

  function openSheet() {
    var url = sheetUrlNow();
    if (url) window.open(url, "_blank", "noopener");
  }

  function refreshSheetRow() {
    var row = document.getElementById("sheetRow");
    if (!row) return;
    var url = sheetUrlNow();
    document.getElementById("sheetOpenBtn").style.display = url ? "" : "none";
    document.getElementById("sheetHint").textContent = url
      ? "Data sheet: " + url
      : "Tip: save a template with its Google Sheet link to open the sheet from here.";
  }

  function addSheetRow() {
    var section = document.querySelector(".merge-section");
    if (!section || document.getElementById("sheetRow")) return;
    var row = document.createElement("div");
    row.className = "sheet-row";
    row.id = "sheetRow";
    row.innerHTML =
      '<button type="button" class="btn btn-small" id="sheetOpenBtn" title="Open the data sheet in a new tab, then File → Download → CSV">📊 Open data sheet</button>' +
      '<small id="sheetHint"></small>';
    var panel = document.getElementById("mergePanel");
    section.insertBefore(row, panel || null);
    document.getElementById("sheetOpenBtn").onclick = openSheet;
    refreshSheetRow();
  }

  function initApp() {
    if (typeof APP === "undefined" || !APP.draftApp || !D) return;
    addStyles();
    addBar();
    addSheetRow();
  }
  document.addEventListener("officedrafts:ready", initApp);
  document.addEventListener("officetemplates:changed", function () {
    renderBar();
  });

  return {
    placeholders: placeholders,
    fillDeep: fillDeep,
    closest: closest,
    afterMergeLoad: afterMergeLoad,
    afterMergeClear: afterMergeClear,
    renderBar: renderBar,
    showWindow: showWindow,
    saveDialog: saveDialog,
  };
})();
