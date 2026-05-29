# Office Suite Viewer

Preview Word, Excel, PDF, Markdown, and more directly in VS Code. Compare Excel/CSV with VCS revisions, cell-level blame, 3-way merge editor, TortoiseSVN/Git integration, self-contained HTML export, and Mermaid v11 diagram support.

## Office Viewer (Fork)

> Forked from [cweijan/vscode-office](https://github.com/cweijan/vscode-office), originally maintained by [RJ.Wang](https://github.com/rjwang1982), now maintained by [haroldcc](https://github.com/haroldcc).

This extension is a maintained fork with major enhancements in VCS integration, diff/merge tooling, and robustness, on top of the prior improvements by RJ.Wang.

## What's new in this fork

### VCS Deep Integration (NEW)
- **Cell-level Blame** — See who last modified each cell in xlsx/csv files, with history replay and LRU cache for performance.
- **Office SCM Panel** — Independent SCM panel showing xlsx/csv file changes in the workspace.
- **Revision Diff** — Compare Excel/CSV files with any Git/SVN revision via QuickPick revision picker.
- **Configurable Blame Depth** — `vscode-office.blame.depth` controls how far back blame scans.

### Diff Enhancements (NEW)
- **Only-Changes Filter** — Toggle to show only rows with differences instead of the full sheet.
- **HTML Diff Export** — Export diff results as self-contained HTML for sharing.
- **Sticky Header** — Column headers stay visible while scrolling through large diffs.
- **Cell-Level Diff Highlighting** — Individual changed cells are highlighted instead of full-row coloring.
- **Reference Switcher Dropdown** — Quick ref switching directly in the diff panel toolbar.
- **Cell Blame Popover** — Hover/click a cell in the diff to see its last commit info.

### 3-Way Merge Editor (NEW)
- Resolve xlsx/csv merge conflicts with a visual 3-way merge editor (theirs → base ← ours).
- Auto-detect conflicts on file open with `vscode-office.merge.autoPromptOnConflict`.

### TortoiseSVN/TortoiseGit Integration (NEW)
- Context menu commands: Show Log, Diff, Blame for xlsx/csv files (Windows only).
- Configure Office Suite Viewer as external diff tool for Tortoise.

### Robustness Improvements (NEW)
- **CSV Delimiter Auto-Detect** — Automatically detects comma, semicolon, tab, or pipe delimiters.
- **CSV Conflict Detection** — Proactive `saveError` warnings when CSV data conflicts with detected format.
- **Encoding Auto-Detection** — Status bar encoding selector with real-time re-decode without reloading.
- **Lazy Sheet Loading** — Large (>10MB) XLSX files only load sheets on demand for instant open.

### Editor Experience (inherited and enhanced)
- **Self-Contained HTML Export** — Local images embedded as Base64, fully portable single-file HTML.
- **Mermaid v11 Support** — Upgraded from v8.8.0 to v11.14.0 with local loading for offline reliability.
- **Configurable Editor Mode** — WYSIWYG, Instant Rendering, or Split View (`vscode-office.editorMode`).
- **Cleaner Rendering** — Content fills available editor width, left-aligned layout.
- **Smaller Package** — Removed bundled Icon Theme and Java Decompiler (~4.4 MB saved), optimized to ~5.25 MB.
- **One Dark Modern Themes** — Bundled editor themes for comfortable editing.

## Introduction

English | [简体中文](README-CN.md)

This extension supports previewing the following file types in VS Code:

- Excel: `.xls`, `.xlsx`, `.csv`
- Word: `.docx`
- SVG: `.svg`
- PDF: `.pdf`
- Fonts: `.ttf`, `.otf`, `.woff`, `.woff2`
- Markdown: `.md`
- HTTP requests: `.http`
- Windows Registry files: `.reg`
- Archive files: `.zip`, `.jar`, `.vsix`, `.rar`

## Markdown

This extension replaces the default Markdown editor with Vditor. Please note that Vditor is no longer actively maintained.

If you want to use the original VS Code editor, add the following to your `settings.json`:

```json
{
    "workbench.editorAssociations": {
        "*.md": "default",
        "*.markdown": "default"
    }
}
```

Right-click in the editor to export Markdown to PDF, DOCX, or HTML. PDF export requires Chromium, which can be configured via `vscode-office.chromiumPath`.

When exporting to HTML, local images are automatically embedded as Base64, so the exported file is fully self-contained and can be shared directly without losing any images.

![Markdown Editor Screenshot](images/screenshot.png)

Keyboard shortcuts are based on [Vditor shortcuts](shortcut.md), with additional commands:

- Move list up: `Ctrl+Alt+I` / `⌘ ^ I`
- Move list down: `Ctrl+Alt+J` / `⌘ ^ J`
- Edit in VS Code: `Ctrl+Alt+E` / `⌘ ^ E`

Tips:

- Resize the editor with Ctrl/Cmd + mouse scroll
- Open hyperlinks with Ctrl/Meta + click or double-click

## HTML

The HTML editor supports live preview. Press `Ctrl+Shift+V` to open the live view.

## Acknowledgements

This project would not exist without the work of the following authors:

- **[cweijan](https://github.com/cweijan)** — Author of the original [vscode-office](https://github.com/cweijan/vscode-office) extension, which this fork is based on. Also maintains a [customized Vditor build](https://github.com/vscode-ext-studio/vditor) tailored for the extension.
- **[RJ.Wang](https://github.com/rjwang1982)** — Maintainer of [vscode-office-enhanced](https://github.com/rjwang1982/vscode-office), which added Mermaid v11 support, self-contained HTML export, package size optimization, and many editor UX improvements that this fork inherits.
- **[Vanessa219 (Liyuan Li)](https://github.com/Vanessa219)** — Author of [Vditor](https://github.com/Vanessa219/vditor), the browser-based Markdown WYSIWYG editor at the heart of this extension's Markdown editing experience. Developed under the [B3log](https://b3log.org) open-source community, licensed under MIT.

## Credits

- PDF rendering: [mozilla/pdf.js/](https://github.com/mozilla/pdf.js/)
- DOCX rendering: [VolodymyrBaydalka/docxjs](https://github.com/VolodymyrBaydalka/docxjs)
- XLSX rendering:
  - [SheetJS/sheetjs](https://github.com/SheetJS/sheetjs): XLSX parsing
  - [myliang/x-spreadsheet](https://github.com/myliang/x-spreadsheet): XLSX rendering
- HTTP: [Rest Client](https://github.com/Huachao/vscode-restclient)
- Markdown: [Vanessa219/vditor](https://github.com/Vanessa219/vditor)
- Mermaid diagrams: [mermaid-js/mermaid](https://github.com/mermaid-js/mermaid)