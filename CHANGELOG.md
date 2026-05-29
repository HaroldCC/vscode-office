# Change log

# 1.0.0 2026-5-29 (Published by haroldcc)

**First release as Office Suite Viewer** — a fork of [cweijan/vscode-office](https://github.com/cweijan/vscode-office) with major enhancements focused on VCS integration, diff/merge tooling, and robustness.

Based on the upstream work by cweijan and RJ.Wang, with the following new features added:

## VCS Deep Integration (NEW)
- **Cell-level Blame** — See who last modified each cell in xlsx/csv files, with history replay and LRU cache for performance.
- **Office SCM Panel** — Independent SCM panel showing xlsx/csv file changes in the workspace.
- **Revision Diff** — Compare Excel/CSV files with any Git/SVN revision via QuickPick revision picker.
- **GitApiHelper** — Singleton wrapping the `vscode.git` Extension API for deep Git integration.
- **CLI Fallback** — Graceful fallback to `git`/`svn` CLI when the Git extension API is unavailable.
- **Configurable Blame Depth** — `vscode-office.blame.depth` controls how far back blame scans.

## Diff Enhancements (NEW)
- **Only-Changes Filter** — Toggle to show only rows with differences instead of the full sheet.
- **HTML Diff Export** — Export diff results as self-contained HTML for sharing.
- **Sticky Header** — Column headers stay visible while scrolling through large diffs.
- **Cell-Level Diff Highlighting** — Individual changed cells are highlighted instead of full-row coloring.
- **Reference Switcher Dropdown** — Quick ref switching directly in the diff panel toolbar.
- **Cell Blame Popover** — Hover/click a cell in the diff to see its last commit info.
- **Editable Current Side** — Diff current side supports edits with auto-encoding detection.

## 3-Way Merge Editor (NEW)
- Resolve xlsx/csv merge conflicts with a visual 3-way merge editor (theirs → base ← ours).
- Auto-detect conflicts on file open with `vscode-office.merge.autoPromptOnConflict`.
- Manual trigger via `office.merge.openConflictEditor` command.

## TortoiseSVN/TortoiseGit Integration (NEW)
- Context menu commands: Show Log, Diff, Blame for xlsx/csv files (Windows only).
- Configure Office Suite Viewer as external diff tool for Tortoise.
- Auto-detect TortoiseProc.exe path.

## Robustness Improvements (NEW)
- **CSV Delimiter Auto-Detect** — Automatically detects comma, semicolon, tab, or pipe delimiters.
- **CSV Conflict Detection** — Proactive `saveError` warnings when CSV data conflicts with detected format.
- **Encoding Auto-Detection** — Status bar encoding selector with real-time re-decode without reloading.
- **Lazy Sheet Loading** — Large (>10MB) XLSX files only load sheets on demand for instant open.

## Credits
Forked from [rjwang1982/vscode-office](https://github.com/rjwang1982/vscode-office) (Office Viewer Enhanced), originally by [cweijan](https://github.com/cweijan). Thanks to RJ.Wang for the prior maintenance work including Mermaid v11, HTML embed, and package optimizations.

---

*(For full upstream history, see [cweijan/vscode-office](https://github.com/cweijan/vscode-office))*
