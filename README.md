# e-Office Suite

A free tool for drafting e-Office documents — **Letters**, **Proceedings** and **Circulars/Notices** — built for Kerala government offices. Drafts are prepared in Malayalam or English and copied into e-Office with the formatting intact.

No installation, no server, no login. Everything runs in your browser and stays on your computer.

## Features

- **Three apps** — Letter Generator, Proceedings, and Circular/Notice, with a live preview as you type.
- **Settings in one place** — click **⚙ Settings** on the home page:
  - **Office Profile** — fill in your office details once; all apps use them. Use **Download filled profile** / **Upload profile** to move them to another browser or computer, or **Download sample profile** for a file with every field to fill in.
  - **RTI Profile** — a separate Name of S.P.I.O., From Address, Designation and Appellate Authority for RTI replies (choose **RTI** on the Letter card).
  - **Header & Logo** — upload a logo and/or a full header image, and choose per app: *Text only*, *Logo with text* or *Header image*.
  - **Phrases** — manage the phrase list, **Download phrases** / **Upload phrases** as an Excel CSV file.
- **Drafts saved automatically** — every draft is saved in the browser as you type. Reopen it from **📂 Drafts** in the app or from **Recent drafts** on the home page. **Back up drafts & templates** keeps a copy of all drafts, templates and phrases in one file.
- **Templates with placeholders** — write a draft with `{Column name}` placeholders and **💾 Save as template**. **Start from…** (in the app or on the home page card) opens a new draft from it; **📑 Templates** to rename, update, delete, download or upload them.
- **Phrase suggestions** — type the first few letters of a common office phrase (Malayalam with InScript or any keyboard, or English) and press **Tab** to insert it. **Ctrl+Space** shows all phrases. Select text in a form and click **★ Save as phrase** to add your own.
- **Tables from Excel** — insert tables from an Excel file, edit cells, split long tables.
- **Mail merge** — upload a CSV/Excel file (for example downloaded from Google Sheets) whose column headers match the placeholders. The app checks that every placeholder has a column (and points out near misses), fills them in the preview and the e-Office copy, and **Save every row as a draft** makes one finished draft per row. **📊 Open data sheet** opens the Google Sheet saved with the template; **Get from Google Sheet (try)** reads it directly when the sheet is shared as "Anyone with the link" (that makes the data readable by anyone with the link — for personal data, download the CSV instead).
- **Print ready** — clean A4 printout.
- **Private** — no data leaves your computer.

## How to use

**Online:** open the GitHub Pages link and bookmark it.

**Offline (ZIP):**
1. Download the ZIP and extract it. Keep the **`common`** folder next to the HTML files — the apps need it.
2. Open `index.html` in Chrome or Edge.
3. Click **⚙ Settings**, fill in the Office Profile (Malayalam and English) and click **Save Profile**.
4. Open any app and start drafting.

Note: uploading Excel files (tables, mail merge, data sheets) uses the SheetJS library, which is loaded from the internet.

## Your data

Profile, drafts, templates and phrases are kept in this browser on this computer. Clearing the browser's data removes them, so use **Back up drafts & templates** and **Download filled profile** now and then, and before moving to another computer.

## For maintainers

- Shared code is in `common/` (`office-core.js`, `drafts.js`, `templates.js`, `phrases.js`, `office.css`); each app sets its own details in an `APP` object at the top of its script.
- The pages link the shared files with a version stamp (`common/office-core.js?v=…`). After changing anything in `common/`, raise that number in all four HTML pages so browsers load the new files.
- `phrases/starter-phrases.csv` is the starter phrase list (also built into `phrases.js`).

## Developed by

Sreejith M, Head Accountant
Govt. Arts and Science College Uduma, Kasaragod
📧 officesuite888@gmail.com
