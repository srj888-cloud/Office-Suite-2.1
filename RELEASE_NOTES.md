# e-Office Suite 2.1 — release notes (20 November 2026)

## New: templates with placeholders + mail merge

- **💾 Save as template** in every app saves the form (with `{placeholders}` such as `{Name}`, `{Amount}`) under a name, optionally with the link to its Google Sheet tab.
- **Start from…** opens a new draft from a template — in the app, or from the **Start from template…** list on the home page cards. The draft you were working on stays saved.
- **📑 Templates** window: Use, Rename / link, Update with current form, Delete, Download templates / Upload templates (to share with colleagues).
- **Mail merge check**: after uploading the CSV, the app lists any placeholder that has no matching column (with a "did you mean…" hint) and the columns that are not used. Capital letters are ignored (`{name}` = `Name`).
- **Save every row as a draft**: one finished draft per row of the sheet, named after the row.
- **📊 Open data sheet** opens the template's Google Sheet; **Get from Google Sheet (try)** reads a sheet shared as "Anyone with the link" directly (experimental).
- Placeholders are now filled everywhere: also in Excel tables, and in Letter's Copy To and Enclosures in the preview.
- **Back up drafts & templates**: the backup file now includes templates.
- Templates saved by older versions are moved over automatically.

---

# e-Office Suite 2.0 — release notes (19 November 2026)

## New

- **⚙ Settings panel** on the home page with tabs for Office Profile, RTI Profile, Header & Logo and Phrases. The home page now shows a short profile summary and the app cards.
- **Drafts saved in the browser** — automatic saving while typing, **📂 Drafts** window in every app (open, delete, new draft), **Recent drafts** on the home page, and **Back up all drafts / Restore backup** (one file with drafts and phrases). A reload reopens the same draft.
- **Phrase suggestions** — type the first letters of an office phrase (Malayalam InScript or English) and press Tab. 58 starter phrases; **★ Save as phrase** for your own; Ctrl+Space shows all; manage them in Settings → Phrases with **Download phrases / Upload phrases**.
- **Profile files** — buttons renamed to **Upload profile** and **Download filled profile** (now always includes the RTI section), plus a new **Download sample profile**.

## Changed

- Name: **e-Office Suite**.
- Profile fields reordered and renamed (Name of Office Head, Designation of the Approver, Name of S.P.I.O.); Malayalam text under the labels removed.
- Header options renamed **Text only / Logo with text / Header image**; the Header & Logo tab was redesigned and shows what each app really uses.
- One message style in all apps (green for success, red for errors); alert boxes replaced.
- Letter tables get the same border outline as Notice and Proceedings (keeps the outer border in e-Office).

## Fixed

- Saved addresses could be lost when two apps were open in different tabs.
- Header choice did not show in an app that was already open (or reopened with Back/Forward).
- Notice: Clear Form and loading a template removed the To and Copy To sections (Clear Form then failed).
- Proceedings: a failed copy showed "Success!".
- A 0 in Excel data showed as an empty table cell (Proceedings).
- Address titles with special characters showed wrongly in the Letter and Notice dropdowns.
- Letter: importing a template with no To/Reference/Copy To entries left the old first entry.

## Under the hood

- Code used by all three apps moved into `common/` (900–1,100 fewer lines in each app); each app's own details are in its `APP` settings.
- Shared files are linked with a version stamp so browsers load new copies after an update.

## Before updating

Drafts, phrases and the profile live in each user's browser and are kept across the update. On the online version, users who had the old site open should reload the page once.
