# Excel Enhancements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add four enhancements to the Excel/CSV viewer: auto encoding detection, cell-level diff highlighting, arbitrary revision diff, and lazy sheet loading for large files.

**Architecture:** Each feature is independently implementable. Features 1 (encoding) and 4 (lazy load) modify `excel_reader.ts` and `Excel.tsx`. Feature 2 (cell diff) modifies `excel_diff.ts`. Feature 3 (revision diff) modifies `excelDiffProvider.ts`, `extension.ts`, and `package.json`. All features share no conflicting code paths.

**Tech Stack:** TypeScript, React 18, XLSX (SheetJS), x-data-spreadsheet, jschardet (new), VSCode Extension API

---

## File Structure

| File | Action | Responsibility |
|------|--------|---------------|
| `src/react/view/excel/excel_reader.ts` | Modify | Add `detectEncoding()`, `loadSheetNames()`, `loadSingleSheet()` |
| `src/react/view/excel/Excel.tsx` | Modify | Integrate encoding detection + lazy loading |
| `src/react/view/excel/excel_diff.ts` | Modify | Cell-level style assignment in `diffSheet()` |
| `src/provider/excelDiffProvider.ts` | Modify | Add `diffWithRevision()`, `getGitLog()`, `getSvnLog()` |
| `src/extension.ts` | Modify | Register `office.excel.diffWithRevision` command |
| `package.json` | Modify | Add `jschardet` dep, new command, menu contribution |

---

### Task 1: Add jschardet dependency

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Install jschardet**

Run:
```bash
cd F:/Dev/vscode-office && npm install jschardet
```

- [ ] **Step 2: Verify installation**

Run:
```bash
cd F:/Dev/vscode-office && node -e "const j = require('jschardet'); console.log(j.detect(Buffer.from('hello')))"
```
Expected: `{ encoding: 'ascii', confidence: 1 }`

- [ ] **Step 3: Commit**

```bash
git add package.json package-lock.json
git commit -m "chore: add jschardet for CSV encoding detection"
```

---

### Task 2: Implement CSV auto encoding detection

**Files:**
- Modify: `src/react/view/excel/excel_reader.ts`
- Modify: `src/react/view/excel/Excel.tsx`

- [ ] **Step 1: Add encoding detection function to excel_reader.ts**

Add at the top of `src/react/view/excel/excel_reader.ts`, after the XLSX import:

```typescript
import jschardet from 'jschardet';

const ENCODING_NORMALIZE_MAP: { [key: string]: string } = {
    'ascii': 'utf-8',
    'utf8': 'utf-8',
    'utf-8': 'utf-8',
    'gbk': 'gbk',
    'gb2312': 'gb2312',
    'gb18030': 'gb18030',
    'big5': 'big5',
    'euc-tw': 'big5',
    'shift_jis': 'shift_jis',
    'shiftjis': 'shift_jis',
    'euc-jp': 'euc-jp',
    'eucjp': 'euc-jp',
    'euc-kr': 'euc-kr',
    'euckr': 'euc-kr',
    'iso-8859-1': 'iso-8859-1',
    'latin1': 'iso-8859-1',
    'windows-1252': 'windows-1252',
    'cp1252': 'windows-1252',
    'utf-16le': 'utf-16le',
    'utf-16be': 'utf-16be',
    'iso-8859-2': 'iso-8859-2',
    'windows-1250': 'windows-1250',
    'windows-1251': 'windows-1251',
    'koi8-r': 'koi8-r',
    'iso-8859-5': 'iso-8859-5',
    'tis-620': 'tis-620',
    'iso-8859-9': 'iso-8859-9',
};

function normalizeEncoding(raw: string): string {
    const lower = raw.toLowerCase().replace(/[^a-z0-9-]/g, '');
    return ENCODING_NORMALIZE_MAP[lower] || ENCODING_NORMALIZE_MAP[raw.toLowerCase()] || 'utf-8';
}

export function detectEncoding(buffer: ArrayBuffer): { encoding: string; confidence: number } {
    const sample = new Uint8Array(buffer.slice(0, 4096));
    const result = jschardet.detect(Buffer.from(sample));
    if (result && result.confidence >= 0.7) {
        return {
            encoding: normalizeEncoding(result.encoding),
            confidence: result.confidence,
        };
    }
    return { encoding: 'utf-8', confidence: result?.confidence || 0 };
}
```

- [ ] **Step 2: Integrate detection in Excel.tsx**

In `src/react/view/excel/Excel.tsx`, add import:

```typescript
import { loadSheets, detectEncoding } from "./excel_reader.ts";
```

Then modify the `renderExcel` function inside the `useEffect`. Replace the `fetch(path).then(...)` chain (lines 145-198 in the current file) with logic that detects encoding for CSV files when no encoding is explicitly passed:

```typescript
const renderExcel = (path: string, ext: string, encoding: string = '', isEncodingExplicit: boolean = false) => {
    const startTime = Date.now();
    console.log('Loading Excel file...');
    setLoading(true);
    fetch(path).then(response => response.arrayBuffer()).then(res => {
        let effectiveEncoding = encoding || 'utf-8';

        // Auto-detect encoding for CSV when no explicit encoding was set
        if (!isEncodingExplicit && ext?.match(/csv/i)) {
            const detected = detectEncoding(res);
            if (detected.confidence >= 0.7) {
                effectiveEncoding = detected.encoding;
                // Notify extension to update status bar
                handler.emit('detectedEncoding', effectiveEncoding);
            }
        }

        const excelData = loadSheets(res, ext, effectiveEncoding);
        // ... rest of the existing code unchanged ...
```

Update the handler bindings to pass `isEncodingExplicit`:

```typescript
handler.on("open", ({ path, ext, encoding }) => {
    lastPathRef.current = path;
    lastExtRef.current = ext;
    // If encoding is provided by the extension (from globalState), treat as explicit
    renderExcel(path, ext, encoding, !!encoding && encoding !== 'utf-8');
}).on("changeEncoding", (encoding: string) => {
    if (lastPathRef.current) {
        renderExcel(lastPathRef.current, lastExtRef.current, encoding, true);
    }
})
```

- [ ] **Step 3: Handle detectedEncoding event in Extension**

In `src/provider/officeViewerProvider.ts` (or `commonHandler.ts` depending on where the handler bindings are), add a listener for the `detectedEncoding` event that updates the encoding status bar without triggering a re-render. This is informational only — the webview has already used the detected encoding.

Find where `handler.on('init', ...)` is set up and add:

```typescript
handler.on('detectedEncoding', (encoding: string) => {
    encodingStatusBar.setEncoding(uri, encoding);
});
```

Also add a `setEncoding` method to `EncodingStatusBar` in `src/common/encodingStatusBar.ts`:

```typescript
setEncoding(uri: string, encoding: string) {
    this.encodingMap.set(uri, encoding);
    this.state.update(`encoding_${uri}`, encoding);
    if (this.currentUri === uri) {
        this.statusBarItem.text = `$(file-code) ${this.getDisplayName(encoding)}`;
    }
}
```

- [ ] **Step 4: Verify manually**

1. Create a GBK-encoded CSV file (e.g., with Chinese characters)
2. Open it in the extension
3. Confirm it renders correctly without manual encoding selection
4. Confirm the status bar shows "GBK"
5. Manually switch to UTF-8 — confirm it takes priority and persists

- [ ] **Step 5: Commit**

```bash
git add src/react/view/excel/excel_reader.ts src/react/view/excel/Excel.tsx src/common/encodingStatusBar.ts src/provider/officeViewerProvider.ts
git commit -m "feat: auto-detect CSV encoding using jschardet"
```

---

### Task 3: Implement cell-level diff highlighting

**Files:**
- Modify: `src/react/view/excel/excel_diff.ts`

- [ ] **Step 1: Add new style constants**

In `src/react/view/excel/excel_diff.ts`, replace the existing styles section (lines 34-37) with:

```typescript
const STYLE_ADDED = { bgcolor: '#d4edda' };
const STYLE_DELETED = { bgcolor: '#f8d7da', color: '#721c24' };
const STYLE_MODIFIED_CELL = { bgcolor: '#fff3cd' };
const STYLE_EMPTY_PLACEHOLDER = { bgcolor: '#f5f5f5' };
const STYLE_UNCHANGED_IN_MODIFIED_ROW = { bgcolor: '#ffffff' };
```

- [ ] **Step 2: Update style arrays in diffSheet()**

In the `diffSheet()` function (line 228), update the style arrays to include the new styles:

```typescript
const leftStyles = [STYLE_DELETED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
const rightStyles = [STYLE_ADDED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
const DELETED_IDX = 0, MODIFIED_CELL_IDX = 1, EMPTY_IDX = 2, UNCHANGED_CELL_IDX = 3;
const ADDED_IDX = 0;
```

- [ ] **Step 3: Apply per-cell styles in match rows**

Replace the match-row handling code inside the `for (const row of aligned)` loop (lines 264-292). The key change: instead of assigning `style: MODIFIED_IDX` to ALL cells when the row is modified, assign it only to cells that actually differ. Unchanged cells in a modified row get `UNCHANGED_CELL_IDX`:

```typescript
if (row.type === 'match') {
    const baseRow = baseRows[row.baseRi];
    const currRow = currRows[row.currRi];
    let rowModified = false;

    const leftCells: any = {};
    const rightCells: any = {};

    for (let ci = 0; ci < maxCols; ci++) {
        const baseCell = baseRow?.cells?.[ci];
        const currCell = currRow?.cells?.[ci];
        const baseText = baseCell?.text != null ? String(baseCell.text) : '';
        const currText = currCell?.text != null ? String(currCell.text) : '';

        if (baseText !== currText) {
            leftCells[ci] = { text: baseText, style: MODIFIED_CELL_IDX };
            rightCells[ci] = { text: currText, style: MODIFIED_CELL_IDX };
            rowModified = true;
        } else if (baseText) {
            leftCells[ci] = { text: baseText };
            rightCells[ci] = { text: currText };
        }
    }

    if (rowModified) {
        // Go back and assign UNCHANGED_CELL_IDX to cells that didn't change in this row
        for (let ci = 0; ci < maxCols; ci++) {
            if (leftCells[ci] && leftCells[ci].style === undefined) {
                leftCells[ci].style = UNCHANGED_CELL_IDX;
            }
            if (rightCells[ci] && rightCells[ci].style === undefined) {
                rightCells[ci].style = UNCHANGED_CELL_IDX;
            }
        }
        changeRows.push(outRi);
        stats.modified++;
    } else {
        stats.unchanged++;
    }
    leftRows[outRi] = { cells: leftCells };
    rightRows[outRi] = { cells: rightCells };
    outRi++;
}
```

Note: The stats counting changes — previously it counted per-cell, now it counts per-row for `modified` and `unchanged`. This is a behavioral improvement (stats now represent row-level changes which matches the UI navigation).

- [ ] **Step 4: Update stats counting in added/deleted blocks**

For `added` and `deleted` rows, count rows instead of cells for consistency:

```typescript
} else if (row.type === 'added') {
    // ... existing cell building logic stays the same ...
    changeRows.push(outRi);
    stats.added++;
    // Remove the per-cell `if (text) stats.added++` lines
    // ...
} else if (row.type === 'deleted') {
    // ... existing cell building logic stays the same ...
    changeRows.push(outRi);
    stats.deleted++;
    // Remove the per-cell `if (text) stats.deleted++` lines
    // ...
}
```

- [ ] **Step 5: Update computeDiff stats aggregation**

In `computeDiff()`, ensure the totalStats aggregation still works (it does — it just sums the sheet-level stats).

- [ ] **Step 6: Verify manually**

1. Create two versions of a small XLSX file where only some cells in certain rows differ
2. Trigger diff (via VCS or file compare)
3. Confirm: modified rows show individual cells highlighted yellow, while unchanged cells in the same row remain white
4. Confirm: added rows are still fully green, deleted rows fully red

- [ ] **Step 7: Commit**

```bash
git add src/react/view/excel/excel_diff.ts
git commit -m "feat: cell-level diff highlighting instead of full-row coloring"
```

---

### Task 4: Implement arbitrary revision diff

**Files:**
- Modify: `src/provider/excelDiffProvider.ts`
- Modify: `src/extension.ts`
- Modify: `package.json`

- [ ] **Step 1: Add getGitLog and getSvnLog methods to ExcelDiffProvider**

In `src/provider/excelDiffProvider.ts`, add these methods after the existing `getSvnVersion()` method:

```typescript
private getGitLog(filePath: string, limit: number = 20): Promise<{ hash: string; message: string; date: string }[]> {
    return new Promise((resolve, reject) => {
        const cwd = dirname(filePath);
        execFile('git', ['log', '--oneline', '--follow', `-${limit}`, '--format=%H|%s|%ar', '--', filePath],
            { cwd, encoding: 'utf8' }, (err, stdout) => {
                if (err) return reject(err);
                const lines = stdout.trim().split('\n').filter(Boolean);
                const entries = lines.map(line => {
                    const [hash, message, date] = line.split('|');
                    return { hash, message: message || '', date: date || '' };
                });
                resolve(entries);
            });
    });
}

private getSvnLog(filePath: string, limit: number = 20): Promise<{ hash: string; message: string; date: string }[]> {
    return new Promise((resolve, reject) => {
        const cwd = dirname(filePath);
        execFile('svn', ['log', '-l', String(limit), '--xml', filePath],
            { cwd, encoding: 'utf8' }, (err, stdout) => {
                if (err) return reject(err);
                // Parse XML manually (simple regex — no external dep needed)
                const entries: { hash: string; message: string; date: string }[] = [];
                const logEntries = stdout.match(/<logentry[^>]*>[\s\S]*?<\/logentry>/g) || [];
                for (const entry of logEntries) {
                    const rev = entry.match(/revision="(\d+)"/)?.[1] || '';
                    const msg = entry.match(/<msg>([\s\S]*?)<\/msg>/)?.[1]?.trim() || '';
                    const date = entry.match(/<date>([\s\S]*?)<\/date>/)?.[1]?.substring(0, 10) || '';
                    entries.push({ hash: rev, message: msg, date });
                }
                resolve(entries);
            });
    });
}
```

- [ ] **Step 2: Add getGitFileAtRevision method**

```typescript
private getGitFileAtRevision(filePath: string, hash: string): Promise<Buffer> {
    return new Promise((resolve, reject) => {
        const cwd = dirname(filePath);
        execFile('git', ['rev-parse', '--show-toplevel'], { cwd, encoding: 'utf8' }, (err, stdout) => {
            if (err) return reject(err);
            const gitRoot = stdout.trim();
            const relPath = relative(gitRoot, filePath).replace(/\\/g, '/');
            execFile('git', ['show', `${hash}:${relPath}`], { cwd, encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 }, (err, stdout) => {
                if (err) return reject(err);
                resolve(stdout as unknown as Buffer);
            });
        });
    });
}

private getSvnFileAtRevision(filePath: string, revision: string): Promise<Buffer> {
    return new Promise((resolve, reject) => {
        const cwd = dirname(filePath);
        execFile('svn', ['cat', '-r', revision, filePath], { cwd, encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 }, (err, stdout) => {
            if (err) return reject(err);
            resolve(stdout as unknown as Buffer);
        });
    });
}
```

- [ ] **Step 3: Add diffWithRevision method**

```typescript
async diffWithRevision(uri?: vscode.Uri) {
    uri = uri || vscode.window.activeTextEditor?.document.uri;
    if (!uri) {
        vscode.window.showErrorMessage('No file selected for diff.');
        return;
    }

    const filePath = uri.fsPath;
    let vcsType: 'git' | 'svn' | null = null;
    let entries: { hash: string; message: string; date: string }[] = [];

    // Try git first, then svn
    try {
        entries = await this.getGitLog(filePath);
        vcsType = 'git';
    } catch {
        try {
            entries = await this.getSvnLog(filePath);
            vcsType = 'svn';
        } catch {
            vscode.window.showErrorMessage('File is not tracked by Git or SVN.');
            return;
        }
    }

    if (entries.length === 0) {
        vscode.window.showInformationMessage('No revision history found for this file.');
        return;
    }

    // Show QuickPick
    const items: (vscode.QuickPickItem & { hash: string })[] = entries.map(e => ({
        label: vcsType === 'git' ? e.hash.substring(0, 7) : `r${e.hash}`,
        description: e.message.length > 60 ? e.message.substring(0, 60) + '...' : e.message,
        detail: e.date,
        hash: e.hash,
    }));

    const selected = await vscode.window.showQuickPick(items, {
        placeHolder: 'Select a revision to compare with',
        matchOnDescription: true,
    });

    if (!selected) return;

    try {
        let baseBuffer: Buffer;
        if (vcsType === 'git') {
            baseBuffer = await this.getGitFileAtRevision(filePath, selected.hash);
        } else {
            baseBuffer = await this.getSvnFileAtRevision(filePath, selected.hash);
        }

        const baseLabel = selected.label;
        const currentLabel = 'Working';
        this.openDiffPanel(uri, baseBuffer, `${vscode.workspace.asRelativePath(uri)} (${baseLabel} ↔ ${currentLabel})`, baseLabel);
    } catch (err: any) {
        vscode.window.showErrorMessage(`Failed to get revision: ${err.message}`);
    }
}
```

- [ ] **Step 4: Register command in extension.ts**

In `src/extension.ts`, add after the existing `office.excel.diffWithFile` registration (line 32):

```typescript
vscode.commands.registerCommand('office.excel.diffWithRevision', (uri) => excelDiffProvider.diffWithRevision(uri)),
```

- [ ] **Step 5: Add command and menu contribution in package.json**

Add to `contributes.commands` array:

```json
{
    "command": "office.excel.diffWithRevision",
    "title": "Compare with Revision...",
    "category": "Office Viewer",
    "icon": "$(git-compare)"
}
```

Add to `contributes.menus.editor/title` array:

```json
{
    "command": "office.excel.diffWithRevision",
    "when": "resourceExtname =~ /\\.(xlsx|xls|csv|xlsm|ods)/i",
    "group": "navigation@-2"
}
```

Add to `contributes.menus.explorer/context` array:

```json
{
    "command": "office.excel.diffWithRevision",
    "when": "resourceExtname =~ /\\.(xlsx|xls|csv|xlsm|ods)/i",
    "group": "3_compare@3"
}
```

- [ ] **Step 6: Verify manually**

1. Open a git-tracked XLSX file
2. Right-click in explorer → "Compare with Revision..."
3. Confirm QuickPick shows recent commits with hash, message, date
4. Select one → confirm diff view opens with correct comparison
5. Test editor title bar button also works

- [ ] **Step 7: Commit**

```bash
git add src/provider/excelDiffProvider.ts src/extension.ts package.json
git commit -m "feat: support comparing Excel/CSV with arbitrary git/svn revision"
```

---

### Task 5: Implement large file lazy loading

**Files:**
- Modify: `src/react/view/excel/excel_reader.ts`
- Modify: `src/react/view/excel/Excel.tsx`

- [ ] **Step 1: Add loadSheetNames and loadSingleSheet to excel_reader.ts**

At the end of `src/react/view/excel/excel_reader.ts`, add:

```typescript
const LAZY_LOAD_THRESHOLD = 10 * 1024 * 1024; // 10MB

export function shouldLazyLoad(bufferSize: number): boolean {
    return bufferSize > LAZY_LOAD_THRESHOLD;
}

export function loadSheetNames(buffer: ArrayBuffer, ext: string): string[] {
    const ab = new Uint8Array(buffer).buffer;
    if (ext.toLowerCase() === '.csv') {
        return ['Sheet1']; // CSV always has one sheet
    }
    const wb = XLSX.read(ab, { type: 'array', bookSheets: true });
    return wb.SheetNames;
}

export function loadSingleSheet(
    buffer: ArrayBuffer,
    ext: string,
    sheetName: string,
    encoding: string = 'utf-8'
): ExcelData {
    const ab = new Uint8Array(buffer).buffer;

    let wb;
    if (ext.toLowerCase() === '.csv') {
        const text = new TextDecoder(encoding).decode(ab);
        wb = XLSX.read(text, { type: 'string', raw: true });
    } else {
        wb = XLSX.read(ab, { type: 'array', sheets: [sheetName], cellFormula: true, cellNF: true });
    }

    return convert(wb);
}
```

- [ ] **Step 2: Add lazy loading state and logic to Excel.tsx**

In `src/react/view/excel/Excel.tsx`, update imports:

```typescript
import { loadSheets, detectEncoding, shouldLazyLoad, loadSheetNames, loadSingleSheet } from "./excel_reader.ts";
```

Add new refs after the existing refs:

```typescript
const bufferRef = useRef<ArrayBuffer | null>(null);
const loadedSheetsRef = useRef<Set<string>>(new Set());
const allSheetNamesRef = useRef<string[]>([]);
const isLazyModeRef = useRef(false);
const currentEncodingRef = useRef<string>('utf-8');
```

- [ ] **Step 3: Modify renderExcel to support lazy loading**

Replace the body of `renderExcel` after the fetch resolves (inside `.then(res => { ... })`):

```typescript
fetch(path).then(response => response.arrayBuffer()).then(res => {
    let effectiveEncoding = encoding || 'utf-8';

    // Auto-detect encoding for CSV when no explicit encoding was set
    if (!isEncodingExplicit && ext?.match(/csv/i)) {
        const detected = detectEncoding(res);
        if (detected.confidence >= 0.7) {
            effectiveEncoding = detected.encoding;
            handler.emit('detectedEncoding', effectiveEncoding);
        }
    }

    currentEncodingRef.current = effectiveEncoding;
    isCSV.current = ext?.match(/csv/i) !== null;

    let excelData;
    if (shouldLazyLoad(res.byteLength)) {
        // Lazy mode: only load first sheet
        isLazyModeRef.current = true;
        bufferRef.current = res;
        const sheetNames = loadSheetNames(res, ext);
        allSheetNamesRef.current = sheetNames;
        excelData = loadSingleSheet(res, ext, sheetNames[0], effectiveEncoding);
        loadedSheetsRef.current = new Set([sheetNames[0]]);
    } else {
        // Normal mode: load all sheets
        isLazyModeRef.current = false;
        bufferRef.current = null;
        excelData = loadSheets(res, ext, effectiveEncoding);
        allSheetNamesRef.current = excelData.sheets.map(s => s.name);
        loadedSheetsRef.current = new Set(allSheetNamesRef.current);
    }

    const { sheets, maxLength, maxCols } = excelData;
    workbookRef.current = excelData.workbook || null;

    container.innerHTML = '';

    // In lazy mode, create placeholder sheets for unloaded tabs
    let sheetData = sheets;
    if (isLazyModeRef.current) {
        sheetData = allSheetNamesRef.current.map((name, idx) => {
            if (idx === 0) return sheets[0]; // First sheet is loaded
            return { name, rows: { 0: { cells: { 0: { text: 'Loading...' } } } } };
        });
    }

    const spreadSheet = new Spreadsheet(container, {
        showToolbar: false,
        row: {
            len: (maxLength || 100) + 50,
            height: 30,
        },
        col: {
            len: maxCols || 26,
        },
        view: {
            height: () => window.innerHeight - 2,
        }
    });
    spreadsheetRef.current = spreadSheet;

    spreadSheet.change(() => {
        if (!document.title.endsWith(DIRTY_MARKER)) {
            document.title = document.title + DIRTY_MARKER;
        }
    });

    // Hook into sheet tab switching for lazy loading
    if (isLazyModeRef.current && spreadSheet.bottombar) {
        const origSwap = spreadSheet.bottombar.swapFunc;
        spreadSheet.bottombar.swapFunc = (index: number) => {
            const sheetName = allSheetNamesRef.current[index];
            if (!loadedSheetsRef.current.has(sheetName) && bufferRef.current) {
                // Load this sheet on demand
                const singleData = loadSingleSheet(
                    bufferRef.current,
                    lastExtRef.current,
                    sheetName,
                    currentEncodingRef.current
                );
                loadedSheetsRef.current.add(sheetName);
                // Replace placeholder data in the spreadsheet's datas array
                const newSheetData = singleData.sheets[0];
                if (newSheetData && spreadSheet.datas[index]) {
                    const d = spreadSheet.datas[index];
                    d.rows.setData(newSheetData.rows);
                    if (newSheetData.cols) {
                        for (const ci of Object.keys(newSheetData.cols)) {
                            d.cols.setWidth(Number(ci), newSheetData.cols[ci].width);
                        }
                    }
                }
            }
            origSwap(index);
        };
    }

    if (keydownRef.current) {
        window.removeEventListener('keydown', keydownRef.current);
    }
    const onKeydown = (e: KeyboardEvent) => {
        if ((e.ctrlKey || e.metaKey) && e.code == "KeyS") {
            export_xlsx(spreadSheet, ext, workbookRef.current);
        }
        if ((e.ctrlKey || e.metaKey) && e.code == "KeyF") {
            e.preventDefault();
            setSearchVisible(true);
            setTimeout(() => searchInputRef.current?.focus(), 50);
        }
        if (e.code === 'Escape') {
            closeSearch();
        }
    };
    keydownRef.current = onKeydown;
    window.addEventListener('keydown', onKeydown);
    setLoading(false);
    spreadSheet.loadData(sheetData);
    const endTime = Date.now();
    console.log(`Excel file loaded successfully. Time elapsed: ${endTime - startTime}ms`);
}).catch(error => {
    console.error(`Failed to load Excel file: ${error.message}`);
    setLoading(false);
});
```

- [ ] **Step 4: Verify manually**

1. Find or create a >10MB XLSX file with multiple sheets
2. Open it in the extension
3. Confirm: loads quickly showing first sheet, status shows other sheet tabs
4. Click a different sheet tab → confirm it loads on demand (brief delay then renders)
5. For files <10MB → confirm all sheets load immediately (existing behavior)

- [ ] **Step 5: Commit**

```bash
git add src/react/view/excel/excel_reader.ts src/react/view/excel/Excel.tsx
git commit -m "feat: lazy-load sheets for large (>10MB) XLSX files"
```

---

### Task 6: Final integration verification

- [ ] **Step 1: Build the extension**

```bash
cd F:/Dev/vscode-office && npm run build
```

Verify no TypeScript errors.

- [ ] **Step 2: Test all features together**

1. Open a GBK CSV → auto-detection works
2. Open a large multi-sheet XLSX → lazy loading works
3. Right-click on a git-tracked XLSX → "Compare with Revision..." shows commit list
4. In diff view → cell-level highlighting shows only changed cells in yellow

- [ ] **Step 3: Commit any fixes**

If any issues found during integration testing, fix and commit separately.
