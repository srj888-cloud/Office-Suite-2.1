/* Office Suite - drafts saved in this browser.

   Drafts are kept in the browser's IndexedDB database "officeSuite21"
   (store "drafts"), so no export/import is needed day to day. A draft is
   the same data the app's own JSON export holds, plus a title and times:
     { id, app: "letter" | "notice" | "proceedings", title,
       created, updated, data }

   Storage belongs to this browser on this computer (and to the address the
   app is opened from). Clearing browsing data removes it, so the Drafts
   window offers "Back up all drafts" (one JSON file) and "Restore backup".

   Templates (reusable starting points, saved with Save as template) live
   in a second store, "templates": { id, app, name, data, sheetUrl,
   created, updated }. A draft started from a template remembers
   templateId / templateName / sheetUrl (used by mail merge).

   Used by the three apps (auto-save + Drafts window) and by index.html
   (Recent drafts). An app takes part by defining, in its APP settings:
     draftApp             "letter" / "notice" / "proceedings"
     getDraftData()       -> object (what the JSON export holds)
     applyDraftData(data)    fills the form from that object
     draftTitle(data)     -> short title for the list
*/
var OfficeDrafts = (function () {
  // Version 2.1 keeps its own database. The live 2.0 site (same web
  // address, so the same browser storage) uses "officeSuite" at version 1
  // and cannot open it once it is upgraded; see migrateFromOld().
  var DB_NAME = "officeSuite21";
  var DB_VERSION = 1; // drafts + templates
  var OLD_DB = "officeSuite"; // version 2.0's drafts
  var STORE = "drafts";
  var TSTORE = "templates";
  var APP_FILES = {
    letter: "Letter Generator.html",
    notice: "Notice.html",
    proceedings: "Proceedings.html",
  };
  var APP_NAMES = {
    letter: "Letter",
    notice: "Circular / Notice",
    proceedings: "Proceedings",
  };
  var APP_ICONS = { letter: "✉️", notice: "📢", proceedings: "⚖️" };

  /* ── storage ── */
  var dbPromise = null;
  function openDb() {
    if (!dbPromise) {
      dbPromise = new Promise(function (resolve, reject) {
        if (!window.indexedDB) {
          reject(new Error("This browser cannot save drafts (no IndexedDB)."));
          return;
        }
        var created = false;
        var req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = function (e) {
          created = e.oldVersion === 0;
          var db = req.result;
          // existing drafts are kept; only missing stores are added
          [STORE, TSTORE].forEach(function (name) {
            if (!db.objectStoreNames.contains(name)) {
              var store = db.createObjectStore(name, { keyPath: "id" });
              store.createIndex("updated", "updated");
              store.createIndex("app", "app");
            }
          });
        };
        req.onblocked = function () {
          reject(new Error("Close the other Office Suite tabs and reload this page to update saved drafts."));
        };
        req.onsuccess = function () {
          var db = req.result;
          if (!created) {
            resolve(db);
            // every start: if the 2.0 database was upgraded again (for
            // example by an old cached copy of 2.1), repair it
            repairOldIfBroken(db);
            return;
          }
          // first start of 2.1 in this browser: bring over the 2.0 drafts
          migrateFromOld(db).then(
            function () {
              resolve(db);
            },
            function () {
              resolve(db);
            },
          );
        };
        req.onerror = function () {
          reject(req.error);
        };
      });
      dbPromise.catch(function () {
        dbPromise = null; // allow a retry later
      });
    }
    return dbPromise;
  }

  /* ── version 2.0's drafts ──
     On the first start of 2.1 in a browser, its drafts (and any templates
     an earlier 2.1 test put there) are copied into this database. If an
     earlier 2.1 had upgraded the 2.0 database to version 2 (which 2.0 can
     no longer open), it is rebuilt at version 1 with its drafts, so the
     2.0 site works again. The 2.0 data is otherwise left untouched. */
  function rawOpen(name, version, upgrade) {
    return new Promise(function (resolve, reject) {
      var q = version ? indexedDB.open(name, version) : indexedDB.open(name);
      q.onupgradeneeded = function () {
        if (upgrade) upgrade(q.result);
      };
      q.onsuccess = function () {
        resolve(q.result);
      };
      q.onerror = function () {
        reject(q.error);
      };
    });
  }
  function readAll(db, store) {
    return new Promise(function (resolve) {
      if (!db.objectStoreNames.contains(store)) {
        resolve([]);
        return;
      }
      var r = db.transaction(store, "readonly").objectStore(store).getAll();
      r.onsuccess = function () {
        resolve(r.result || []);
      };
      r.onerror = function () {
        resolve([]);
      };
    });
  }
  function putAll(db, store, items) {
    return new Promise(function (resolve) {
      if (!items.length) {
        resolve();
        return;
      }
      var tx = db.transaction(store, "readwrite");
      var s = tx.objectStore(store);
      items.forEach(function (i) {
        s.put(i);
      });
      tx.oncomplete = tx.onerror = tx.onabort = function () {
        resolve();
      };
    });
  }
  function oldDbVersion() {
    // only look at the 2.0 database if it exists (opening it would create it)
    if (!indexedDB.databases) return Promise.resolve(0);
    return indexedDB.databases().then(
      function (list) {
        var d = (list || []).filter(function (x) {
          return x.name === OLD_DB;
        })[0];
        return d ? d.version || 0 : 0;
      },
      function () {
        return 0;
      },
    );
  }
  function migrateFromOld(newDb) {
    return oldDbVersion().then(function (v) {
      if (!v) return;
      return rawOpen(OLD_DB).then(function (old) {
        return Promise.all([readAll(old, STORE), readAll(old, TSTORE)]).then(function (both) {
          old.close();
          var drafts = both[0];
          return putAll(newDb, STORE, drafts)
            .then(function () {
              return putAll(newDb, TSTORE, both[1]);
            })
            .then(function () {
              // repair in the background; a 2.0 tab left open only delays it
              if (v >= 2) restoreOldDb(drafts);
            });
        });
      });
    });
  }
  // Later starts: only when the 2.0 database is at version 2 (which 2.0
  // cannot open). Anything in it that 2.1 does not have yet (or has an
  // older copy of) is kept here, then the 2.0 database is rebuilt.
  function repairOldIfBroken(newDb) {
    oldDbVersion().then(function (v) {
      if (v < 2) return;
      return rawOpen(OLD_DB).then(function (old) {
        return Promise.all([readAll(old, STORE), readAll(old, TSTORE)]).then(function (both) {
          old.close();
          return Promise.all([readAll(newDb, STORE), readAll(newDb, TSTORE)]).then(function (mine) {
            function newer(items, have) {
              var byId = {};
              have.forEach(function (h) {
                byId[h.id] = h;
              });
              return items.filter(function (i) {
                return !byId[i.id] || (i.updated || 0) > (byId[i.id].updated || 0);
              });
            }
            return putAll(newDb, STORE, newer(both[0], mine[0]))
              .then(function () {
                return putAll(newDb, TSTORE, newer(both[1], mine[1]));
              })
              .then(function () {
                restoreOldDb(both[0]);
              });
          });
        });
      });
    }).catch(function () {});
  }

  function restoreOldDb(drafts) {
    var del = indexedDB.deleteDatabase(OLD_DB);
    del.onsuccess = function () {
      rawOpen(OLD_DB, 1, function (db) {
        var s = db.createObjectStore(STORE, { keyPath: "id" });
        s.createIndex("updated", "updated");
        s.createIndex("app", "app");
      }).then(function (db) {
        return putAll(db, STORE, drafts).then(function () {
          db.close();
        });
      }, function () {});
    };
  }

  function run(mode, work, storeName) {
    storeName = storeName || STORE;
    return openDb().then(function (db) {
      // another tab upgrading the database closes this connection
      db.onversionchange = function () {
        db.close();
        dbPromise = null;
      };
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(storeName, mode);
        var req = work(tx.objectStore(storeName));
        tx.oncomplete = function () {
          resolve(req ? req.result : undefined);
        };
        tx.onerror = tx.onabort = function () {
          reject(tx.error || new Error("Draft storage error"));
        };
      });
    });
  }

  function get(id) {
    return run("readonly", function (s) {
      return s.get(id);
    });
  }
  function put(draft) {
    return run("readwrite", function (s) {
      return s.put(draft);
    });
  }
  function remove(id) {
    return run("readwrite", function (s) {
      return s.delete(id);
    });
  }
  // newest first; app is optional
  function list(app) {
    return run("readonly", function (s) {
      return s.getAll();
    }).then(function (all) {
      return (all || [])
        .filter(function (d) {
          return !app || d.app === app;
        })
        .sort(function (a, b) {
          return (b.updated || 0) - (a.updated || 0);
        });
    });
  }

  /* templates: same shape of calls as drafts */
  function tplGet(id) {
    return run("readonly", function (s) {
      return s.get(id);
    }, TSTORE);
  }
  function tplPut(tpl) {
    return run("readwrite", function (s) {
      return s.put(tpl);
    }, TSTORE);
  }
  function tplRemove(id) {
    return run("readwrite", function (s) {
      return s.delete(id);
    }, TSTORE);
  }
  // by name (A-Z); app is optional
  function tplList(app) {
    return run("readonly", function (s) {
      return s.getAll();
    }, TSTORE).then(function (all) {
      return (all || [])
        .filter(function (d) {
          return !app || d.app === app;
        })
        .sort(function (a, b) {
          return String(a.name).localeCompare(String(b.name));
        });
    });
  }

  function newId() {
    return (
      Date.now().toString(36) + Math.random().toString(36).slice(2, 8)
    );
  }

  // Ask the browser not to clear our storage when space runs low
  var persistAsked = false;
  function askPersist() {
    if (persistAsked) return;
    persistAsked = true;
    try {
      if (navigator.storage && navigator.storage.persist) {
        navigator.storage.persist();
      }
    } catch (e) {}
  }

  /* ── backup ── */
  function exportAll() {
    // resolves to { drafts, templates, phrases } counts (zeros when nothing was saved)
    return Promise.all([list(), tplList()]).then(function (both) {
      var drafts = both[0];
      var templates = both[1];
      var hasPhrases = window.OfficePhrases && OfficePhrases.isCustom();
      if (!drafts.length && !templates.length && !hasPhrases) {
        return { drafts: 0, templates: 0, phrases: 0 };
      }
      var out = {
        type: "office-suite-drafts",
        version: 2,
        exported: new Date().toISOString(),
        drafts: drafts,
        templates: templates,
      };
      // the phrase list travels with the drafts (only when it was changed)
      if (window.OfficePhrases && OfficePhrases.isCustom()) {
        out.phrases = OfficePhrases.getList();
      }
      var blob = new Blob([JSON.stringify(out, null, 2)], {
        type: "application/json",
      });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download =
        "office_suite_drafts_" + new Date().toISOString().slice(0, 10) + ".json";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      return {
        drafts: drafts.length,
        templates: templates.length,
        phrases: out.phrases ? out.phrases.length : 0,
      };
    });
  }

  // Adds drafts from a backup file. A draft already here is replaced only
  // when the backup copy is newer. Resolves to {added, updated, skipped}.
  function importFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        var json;
        try {
          json = JSON.parse(reader.result);
        } catch (e) {
          reject(new Error("Could not read the file. Choose a drafts backup file."));
          return;
        }
        if (!json || json.type !== "office-suite-drafts" || !Array.isArray(json.drafts)) {
          reject(new Error("This is not an Office Suite drafts backup file."));
          return;
        }
        var counts = { added: 0, updated: 0, skipped: 0, phrases: 0, templates: 0 };
        if (Array.isArray(json.phrases) && window.OfficePhrases) {
          try {
            counts.phrases = OfficePhrases.mergeList(json.phrases);
          } catch (e) {}
        }
        var valid = json.drafts.filter(function (d) {
          return d && d.id && APP_FILES[d.app] && d.data;
        });
        var validTemplates = (Array.isArray(json.templates) ? json.templates : []).filter(function (d) {
          return d && d.id && APP_FILES[d.app] && d.data && d.name;
        });
        Promise.all(
          valid.map(function (d) {
            return get(d.id).then(function (have) {
              if (have && (have.updated || 0) >= (d.updated || 0)) {
                counts.skipped++;
                return;
              }
              counts[have ? "updated" : "added"]++;
              return put(d);
            });
          }).concat(
            validTemplates.map(function (d) {
              return tplGet(d.id).then(function (have) {
                if (have && (have.updated || 0) >= (d.updated || 0)) return;
                counts.templates++;
                return tplPut(d);
              });
            }),
          ),
        ).then(function () {
          askPersist();
          resolve(counts);
        }, reject);
      };
      reader.onerror = function () {
        reject(new Error("Could not read the file."));
      };
      reader.readAsText(file);
    });
  }

  /* ── helpers for lists ── */
  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
  function when(t) {
    if (!t) return "";
    var d = new Date(t);
    var now = new Date();
    var time = d.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
    if (d.toDateString() === now.toDateString()) return "Today " + time;
    var y = new Date(now);
    y.setDate(now.getDate() - 1);
    if (d.toDateString() === y.toDateString()) return "Yesterday " + time;
    return (
      d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }) +
      " " +
      time
    );
  }

  /* ── inside an app: auto-save and the Drafts window ── */
  var current = null; // the draft being edited
  var origin = null; // { templateId, templateName, sheetUrl } when started from a template
  var lastJson = null; // what was last saved (or loaded)
  var timer = null;
  var saving = null;

  function setStatus(text, kind) {
    var el = document.getElementById("draftStatus");
    if (!el) return;
    el.textContent = text;
    el.className = "draft-status" + (kind ? " " + kind : "");
  }

  function urlWithDraft(id) {
    var params = new URLSearchParams(location.search);
    if (id) params.set("draft", id);
    else params.delete("draft");
    params.delete("template");
    var q = params.toString();
    return location.pathname + (q ? "?" + q : "");
  }

  // Save now if anything changed. Resolves when stored.
  function saveNow() {
    clearTimeout(timer);
    timer = null;
    if (saving) return saving.then(saveNow);
    var data;
    try {
      data = APP.getDraftData();
    } catch (e) {
      return Promise.resolve();
    }
    var json = JSON.stringify(data);
    if (json === lastJson) return Promise.resolve();
    var now = Date.now();
    var draft = current
      ? Object.assign({}, current)
      : { id: newId(), app: APP.draftApp, created: now };
    if (!current && origin) {
      if (origin.templateId) {
        draft.templateId = origin.templateId;
        draft.templateName = origin.templateName;
      }
      if (origin.sheetUrl) draft.sheetUrl = origin.sheetUrl;
    }
    draft.title = (APP.draftTitle(data) || "").trim() || "Untitled";
    draft.updated = now;
    draft.data = data;
    setStatus("Saving…");
    saving = put(draft).then(
      function () {
        var isNew = !current;
        current = draft;
        lastJson = json;
        saving = null;
        setStatus("Saved " + new Date(now).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" }), "ok");
        if (isNew) {
          askPersist();
          // a reload now reopens this draft
          try {
            history.replaceState(history.state, "", urlWithDraft(draft.id));
          } catch (e) {}
        }
      },
      function (err) {
        saving = null;
        setStatus("Not saved", "error");
        showToast("Could not save the draft in this browser. " + (err && err.message ? err.message : ""), "error");
      },
    );
    return saving;
  }

  // Any edit or click (form fields, add/remove buttons, address picker,
  // table tools) schedules a check; saveNow() only writes when the form
  // data actually changed. The Drafts bar and window are ignored.
  function changed(e) {
    var t = e.target;
    if (!t || !t.closest) return;
    if (t.closest(".draft-bar, .draft-actions, .draft-list, .draft-backup")) return;
    if (e.type === "input" && t.closest("#formPanel")) setStatus("Editing…");
    clearTimeout(timer);
    timer = setTimeout(saveNow, 1500);
  }

  // After "Clear form": keep the saved draft as it was; further typing
  // starts a new draft instead of overwriting it with an empty form.
  function detach() {
    if (typeof APP === "undefined" || !APP.draftApp) return;
    clearTimeout(timer);
    current = null;
    try {
      lastJson = JSON.stringify(APP.getDraftData());
    } catch (e) {}
    setStatus("New draft");
    try {
      history.replaceState(history.state, "", urlWithDraft(null));
    } catch (e) {}
  }

  function goTo(url) {
    saveNow().then(function () {
      location.href = url;
    });
  }

  function openDraft(id) {
    goTo(urlWithDraft(id));
  }

  function newDraft() {
    goTo(urlWithDraft(null));
  }

  // Start a new draft from a template (the current draft is saved first)
  function openTemplate(id) {
    var url = urlWithDraft(null);
    goTo(url + (url.indexOf("?") === -1 ? "?" : "&") + "template=" + encodeURIComponent(id));
  }

  // Which template (and Google Sheet link) this draft came from, if any
  function getOrigin() {
    var src = current || origin;
    if (!src || (!src.templateId && !src.sheetUrl)) return null;
    return { templateId: src.templateId || "", templateName: src.templateName || "", sheetUrl: src.sheetUrl || "" };
  }

  // Remember a Google Sheet link for this draft (and the template it came from)
  function setSheetUrl(url) {
    if (current) {
      current.sheetUrl = url;
      put(current);
    } else {
      origin = Object.assign({}, origin || {}, { sheetUrl: url });
    }
    var o = getOrigin();
    if (o && o.templateId) {
      tplGet(o.templateId).then(function (tpl) {
        if (tpl && tpl.sheetUrl !== url) {
          tpl.sheetUrl = url;
          tplPut(tpl);
        }
      });
    }
  }

  // Save a separate, finished draft (used by "Save every row as a draft")
  function saveNew(data, title) {
    var now = Date.now();
    var o = getOrigin() || {};
    var draft = {
      id: newId(),
      app: APP.draftApp,
      created: now,
      updated: now,
      title: (title || "").trim() || "Untitled",
      data: data,
    };
    if (o.templateId) {
      draft.templateId = o.templateId;
      draft.templateName = o.templateName;
      draft.sheetUrl = o.sheetUrl;
    }
    return put(draft).then(function () {
      askPersist();
      return draft;
    });
  }

  function deleteDraft(id) {
    get(id).then(function (d) {
      if (!d) return;
      if (!confirm('Delete the draft "' + d.title + '"? This cannot be undone.')) return;
      remove(id).then(function () {
        if (current && current.id === id) {
          current = null;
          lastJson = JSON.stringify(APP.getDraftData()); // don't re-save it straight away
          setStatus("Draft deleted");
          try {
            history.replaceState(history.state, "", urlWithDraft(null));
          } catch (e) {}
        }
        showToast("Draft deleted");
        showWindow();
      });
    });
  }

  function showWindow() {
    saveNow().then(function () {
      return list(APP.draftApp);
    }).then(
      function (drafts) {
        var html =
          '<div class="draft-actions">' +
          '<button type="button" class="btn btn-small" onclick="OfficeDrafts.newDraft()">+ New draft</button>' +
          '<span class="draft-hint">Drafts are saved automatically in this browser as you type.</span>' +
          "</div>";
        if (!drafts.length) {
          html += '<p class="draft-empty">No saved drafts yet. Start typing in the form and it is saved here.</p>';
        } else {
          html += '<ul class="draft-list">';
          drafts.forEach(function (d) {
            var isOpen = current && current.id === d.id;
            html +=
              '<li class="draft-item' + (isOpen ? " current" : "") + '">' +
              '<div class="draft-info"><strong>' + esc(d.title) + "</strong>" +
              "<small>" + esc(when(d.updated)) + (isOpen ? " · open now" : "") + "</small></div>" +
              (isOpen
                ? ""
                : '<button type="button" class="btn btn-primary btn-small" onclick="OfficeDrafts.openDraft(\'' + esc(d.id) + "')\">Open</button>") +
              '<button type="button" class="btn btn-danger btn-small" onclick="OfficeDrafts.deleteDraft(\'' + esc(d.id) + "')\">Delete</button>" +
              "</li>";
          });
          html += "</ul>";
        }
        html +=
          '<div class="draft-backup">' +
          '<button type="button" class="btn btn-small" onclick="OfficeDrafts.backup()">Back up drafts &amp; templates</button>' +
          '<button type="button" class="btn btn-small" onclick="document.getElementById(\'draftRestoreInput\').click()">Restore backup</button>' +
          '<input type="file" id="draftRestoreInput" accept=".json,application/json" hidden onchange="OfficeDrafts.restore(this)">' +
          '<span class="draft-hint">One file with the drafts, templates and phrases of all three apps. Keep a copy before clearing browser data or changing computers.</span>' +
          "</div>";
        showModal("My Drafts · " + (APP_NAMES[APP.draftApp] || ""), html);
      },
      function (err) {
        showToast("Could not open saved drafts. " + (err && err.message ? err.message : ""), "error");
      },
    );
  }

  function backup() {
    saveNow()
      .then(exportAll)
      .then(
        function (c) {
          if (!c.drafts && !c.templates && !c.phrases) {
            showToast("There is nothing to back up yet.", "error");
            return;
          }
          showToast(
            "Backed up " + c.drafts + " draft(s)" +
              (c.templates ? ", " + c.templates + " template(s)" : "") +
              (c.phrases ? " and " + c.phrases + " phrase(s)" : "") + ".",
          );
        },
        function (err) {
          showToast("Could not back up drafts. " + err.message, "error");
        },
      );
  }

  function restore(input) {
    var file = input.files && input.files[0];
    input.value = "";
    if (!file) return;
    importFile(file).then(
      function (c) {
        showToast(
          "Restored: " + c.added + " new, " + c.updated + " updated" +
            (c.skipped ? ", " + c.skipped + " already up to date" : "") +
            (c.templates ? ", " + c.templates + " template(s)" : "") +
            (c.phrases ? ", " + c.phrases + " phrase(s) added" : "") + ".",
        );
        if (document.getElementById("draftRestoreInput")) showWindow();
      },
      function (err) {
        showToast(err.message, "error");
      },
    );
  }

  function addBar() {
    var header = document.querySelector("#formPanel .panel-header");
    if (!header || document.getElementById("draftStatus")) return;
    var bar = document.createElement("span");
    bar.className = "draft-bar";
    bar.innerHTML =
      '<span id="draftStatus" class="draft-status"></span>' +
      '<button type="button" class="draft-btn" onclick="OfficeDrafts.showWindow()" title="Saved drafts in this browser">📂 Drafts</button>';
    var fsBtn = document.getElementById("formFsBtn");
    header.insertBefore(bar, fsBtn || null);
  }

  function initApp() {
    if (typeof APP === "undefined" || !APP.draftApp) return;
    addBar();
    var params = new URLSearchParams(location.search);
    var id = params.get("draft");
    var templateId = !id && params.get("template");
    var ready = templateId
      ? tplGet(templateId).then(
          function (tpl) {
            if (!tpl || tpl.app !== APP.draftApp) {
              showToast("That template was not found in this browser.", "error");
              return;
            }
            APP.applyDraftData(tpl.data);
            origin = { templateId: tpl.id, templateName: tpl.name, sheetUrl: tpl.sheetUrl || "" };
            setStatus("New draft from “" + tpl.name + "”", "ok");
          },
          function (err) {
            showToast("Saved templates are not available in this browser. " + (err && err.message ? err.message : ""), "error");
          },
        )
      : id
      ? get(id).then(
          function (d) {
            if (!d || d.app !== APP.draftApp) {
              showToast("That draft was not found in this browser.", "error");
              try {
                history.replaceState(history.state, "", urlWithDraft(null));
              } catch (e) {}
              return;
            }
            APP.applyDraftData(d.data);
            current = d;
            setStatus("Opened " + when(d.updated), "ok");
          },
          function (err) {
            setStatus("Drafts unavailable", "error");
            showToast("Saved drafts are not available in this browser. " + (err && err.message ? err.message : ""), "error");
          },
        )
      : Promise.resolve();
    ready.then(function () {
      try {
        lastJson = JSON.stringify(APP.getDraftData()); // starting point: no save until something changes
      } catch (e) {}
      ["input", "change", "click"].forEach(function (type) {
        document.addEventListener(type, changed, true);
      });
      document.addEventListener("visibilitychange", function () {
        if (document.visibilityState === "hidden") saveNow();
      });
      window.addEventListener("pagehide", saveNow);
      // the template bar and mail merge tools (templates.js) start now
      document.dispatchEvent(new Event("officedrafts:ready"));
    });
  }

  // Templates saved by older versions in the shared "ml_templates" slot
  // (Letter and Notice) are moved to the templates store once.
  function moveOldTemplates() {
    var old;
    try {
      if (localStorage.getItem("ml_templates_moved")) return;
      old = JSON.parse(localStorage.getItem("ml_templates") || "[]");
    } catch (e) {
      return;
    }
    if (!Array.isArray(old) || !old.length) {
      try {
        localStorage.setItem("ml_templates_moved", "1");
      } catch (e) {}
      return;
    }
    tplList().then(function (have) {
      var jobs = [];
      old.forEach(function (o, i) {
        if (!o || !o.data || !o.name) return;
        var d = o.data;
        var app = d.noticeTitle !== undefined || Array.isArray(d.to) ? "notice" : d.toAddresses !== undefined ? "letter" : null;
        if (!app) return;
        var dup = have.some(function (h) {
          return h.app === app && h.name === o.name;
        });
        if (dup) return;
        var now = Date.now() + i;
        jobs.push(tplPut({ id: newId(), app: app, name: String(o.name), data: d, sheetUrl: "", created: now, updated: now }));
      });
      return Promise.all(jobs).then(function () {
        return jobs.length;
      });
    }).then(function (moved) {
      try {
        localStorage.setItem("ml_templates_moved", "1");
      } catch (e) {}
      // let open lists (home page cards, template bar) show them
      if (moved) document.dispatchEvent(new Event("officetemplates:changed"));
    }, function () {});
  }

  // run after each app's own start-up code (window.onload / DOMContentLoaded)
  if (typeof window !== "undefined") {
    window.addEventListener("load", function () {
      setTimeout(initApp, 0);
      setTimeout(moveOldTemplates, 50);
    });
  }

  return {
    APP_FILES: APP_FILES,
    APP_NAMES: APP_NAMES,
    APP_ICONS: APP_ICONS,
    list: list,
    get: get,
    remove: remove,
    exportAll: exportAll,
    importFile: importFile,
    esc: esc,
    when: when,
    saveNow: saveNow,
    detach: detach,
    showWindow: showWindow,
    openDraft: openDraft,
    newDraft: newDraft,
    deleteDraft: deleteDraft,
    backup: backup,
    restore: restore,
    // templates
    tplGet: tplGet,
    tplPut: tplPut,
    tplRemove: tplRemove,
    tplList: tplList,
    newId: newId,
    openTemplate: openTemplate,
    getOrigin: getOrigin,
    setSheetUrl: setSheetUrl,
    saveNew: saveNew,
  };
})();
