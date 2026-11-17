# Paperline roadmap

Paperline is built in phases. Each phase ends with something people can use, and every phase keeps one codebase for web, desktop and mobile.

## Phase 1: Web app with the three core editors (this PR)

- Start screen with a "what would you like to create?" picker; the ribbon adapts to the file type.
- Document, spreadsheet and presentation editors.
- Open and save `.docx`, `.xlsx`, `.pptx`, plus `.csv`, `.txt`, `.md` and `.html` in.
- Native `.pline` format (versioned zip with a JSON model).
- PDF export and direct printing.
- English and Arabic interface with full right-to-left layout.
- Formula engine with unit tests; CI builds and tests every push.

- A PDF app, device fonts, templates, and Word, Excel and PowerPoint tools (see the README for the full list).

**Known limits** that later phases fix:
- Word import keeps structure (headings, lists, tables, images, bold, italic) but drops colors, fonts, alignment, headers, footers and footnotes.
- PDF export of documents, sheets and slides draws pages as images, so their text cannot be selected or searched.
- Spreadsheet charts are kept in `.pline` files but not yet written to `.xlsx`.
- Presentation text boxes use one style per box, and there are no per-object animations yet.

## Phase 2: Desktop apps for Windows and macOS

- Wrap the web app with **Tauri 2** (small installers, about 10 MB, using the system web view).
- Native open and save dialogs, recent files, "Open with Paperline" and double-click to open `.docx`, `.xlsx`, `.pptx` and `.pline`.
- Autosave and crash recovery to a local folder.
- Native print dialog and printer selection.
- Signed installers (`.msi` for Windows, `.dmg` for macOS) and an auto-updater, built by GitHub Actions.
- Works fully offline.

## Phase 3: Mobile apps for iPhone and Android

- Package with **Capacitor**, sharing the same React code.
- Touch-first ribbon: a compact bottom toolbar with sheets for each tool group.
- Pinch to zoom, touch cell selection and handles, an on-screen keyboard that does not cover the cell or text being edited.
- Open from the Files app or Google Files and from email attachments, and share back out.
- App Store and Play Store builds.

## Phase 4: Closer to full Office fidelity

**Documents**
- Own `.docx` reader that keeps fonts, colors, alignment, spacing, page size and margins.
- Headers, footers, page numbers, footnotes, table of contents, columns, section breaks.
- Track changes and comments that round-trip with Word.
- Find and replace, spell check (English and Arabic), real page view with live pagination.

**Spreadsheets**
- Charts written into `.xlsx` files.
- Freeze panes of any size, pivot tables.
- 300+ functions, named ranges, array formulas.
- Large files: virtual columns as well as rows, background calculation in a web worker.

**Presentations**
- Rich text inside boxes (mixed styles per word).
- Per-object animations, grouping, and a second-screen presenter view.
- Master slides and layouts that follow the opened `.pptx` theme.

**Everywhere**
- Vector PDF export with selectable text and embedded fonts (Arabic shaping included).
- Legacy formats `.doc`, `.xls` and `.ppt`, plus OpenDocument `.odt`, `.ods` and `.odp`.

## Phase 5: The rest of the Office family

Each one gets its own card in the "New file" picker and its own ribbon:

- **Notebook** (like OneNote): sections, pages, free-placed notes, checklists; import `.one` where possible.
- **Diagram** (like Visio): shapes, connectors, flowchart and org-chart templates; open and save `.vsdx`.
- **Publication** (like Publisher): page layout with frames, flyers, cards; import `.pub` as read-only.
- **Database** (like Access): tables, forms, simple queries and reports; import `.accdb`/`.mdb` tables.
- **Project plan** (like Project): tasks, Gantt timeline, dependencies; open `.mpp` read-only and `.xml` read-write.
- **Email drafts** (like Outlook `.msg`/`.eml` viewer and composer), as a later stretch.

## Phase 6: Optional sync and collaboration (later)

Files stay local by default. If wanted later: optional sign-in, sync between your own devices, sharing a link, and real-time co-editing (CRDT based, for example Yjs) with comments and version history.

## Ongoing

- Accessibility: full keyboard use, screen reader labels, high-contrast theme.
- Dark mode for the app chrome.
- More interface languages after English and Arabic.
- Performance budgets and automated round-trip tests against a library of real Office files.
