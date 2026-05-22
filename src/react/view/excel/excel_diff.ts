import { ExcelData } from './excel_reader';

// ================ Types ================

export interface DiffStats {
    added: number;
    deleted: number;
    modified: number;
    unchanged: number;
}

export interface DiffSheetData {
    name: string;
    rows: { [key: number]: { cells: { [key: number]: { text: string; style?: number } } } };
    styles: any[];
    maxCols: number;
    maxRows: number;
    changeRows: number[]; // Row indices that have changes (added/deleted/modified)
}

export interface DiffResult {
    leftSheets: DiffSheetData[];
    rightSheets: DiffSheetData[];
    stats: DiffStats;
}

type AlignedRow =
    | { type: 'match'; baseRi: number; currRi: number }
    | { type: 'added'; currRi: number }
    | { type: 'deleted'; baseRi: number };

// ================ Styles ================

const STYLE_ADDED = { bgcolor: '#d4edda' };
const STYLE_DELETED = { bgcolor: '#f8d7da', color: '#721c24' };
const STYLE_MODIFIED_CELL = { bgcolor: '#fff3cd' };
const STYLE_EMPTY_PLACEHOLDER = { bgcolor: '#f5f5f5' };
const STYLE_UNCHANGED_IN_MODIFIED_ROW = { bgcolor: '#ffffff' };

// ================ Helpers ================

/**
 * Generate a fingerprint for a row (concatenation of all cell texts).
 */
function rowFingerprint(row: any, maxCols: number): string {
    if (!row || !row.cells) return '';
    const parts: string[] = [];
    for (let ci = 0; ci < maxCols; ci++) {
        const cell = row.cells[ci];
        parts.push(cell?.text != null ? String(cell.text) : '');
    }
    return parts.join('|');
}

/**
 * Simple hash for chunked comparison — fast, not cryptographic.
 */
function simpleHash(str: string): number {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        const ch = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + ch;
        hash = hash | 0;
    }
    return hash;
}

/**
 * Compute LCS (Longest Common Subsequence) alignment between two arrays of fingerprints.
 * Uses Hunt-Szymanski for sparse matches or standard DP for dense ones.
 * Returns aligned row array.
 */
function lcsAlign(baseFP: string[], currFP: string[]): AlignedRow[] {
    const m = baseFP.length;
    const n = currFP.length;

    // For very large inputs, use positional comparison
    if (m > 5000 || n > 5000) {
        return positionalAlign(baseFP, currFP);
    }

    // Standard O(mn) LCS — limited to reasonable sizes by chunking
    const dp: number[][] = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));

    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            if (baseFP[i - 1] === currFP[j - 1]) {
                dp[i][j] = dp[i - 1][j - 1] + 1;
            } else {
                dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
            }
        }
    }

    // Backtrack to build alignment
    const aligned: AlignedRow[] = [];
    let i = m, j = n;
    const result: AlignedRow[] = [];

    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && baseFP[i - 1] === currFP[j - 1]) {
            result.push({ type: 'match', baseRi: i - 1, currRi: j - 1 });
            i--;
            j--;
        } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
            result.push({ type: 'added', currRi: j - 1 });
            j--;
        } else {
            result.push({ type: 'deleted', baseRi: i - 1 });
            i--;
        }
    }

    return result.reverse();
}

/**
 * Fallback for very large sheets: simple positional comparison.
 */
function positionalAlign(baseFP: string[], currFP: string[]): AlignedRow[] {
    const aligned: AlignedRow[] = [];
    const maxLen = Math.max(baseFP.length, currFP.length);

    for (let i = 0; i < maxLen; i++) {
        if (i < baseFP.length && i < currFP.length) {
            aligned.push({ type: 'match', baseRi: i, currRi: i });
        } else if (i >= baseFP.length) {
            aligned.push({ type: 'added', currRi: i });
        } else {
            aligned.push({ type: 'deleted', baseRi: i });
        }
    }

    return aligned;
}

/**
 * Chunked alignment: divide rows into blocks, identify changed blocks,
 * then only run LCS on changed blocks.
 */
function chunkedAlign(baseFP: string[], currFP: string[], chunkSize: number = 500): AlignedRow[] {
    // If small enough, just do full LCS
    if (baseFP.length <= chunkSize && currFP.length <= chunkSize) {
        return lcsAlign(baseFP, currFP);
    }

    // Build chunk-level hashes
    const baseChunks: { hash: number; start: number; end: number }[] = [];
    for (let i = 0; i < baseFP.length; i += chunkSize) {
        const end = Math.min(i + chunkSize, baseFP.length);
        const chunkStr = baseFP.slice(i, end).join('\n');
        baseChunks.push({ hash: simpleHash(chunkStr), start: i, end });
    }

    const currChunks: { hash: number; start: number; end: number }[] = [];
    for (let i = 0; i < currFP.length; i += chunkSize) {
        const end = Math.min(i + chunkSize, currFP.length);
        const chunkStr = currFP.slice(i, end).join('\n');
        currChunks.push({ hash: simpleHash(chunkStr), start: i, end });
    }

    // Match chunks by position and hash
    const aligned: AlignedRow[] = [];
    const maxChunks = Math.max(baseChunks.length, currChunks.length);

    for (let ci = 0; ci < maxChunks; ci++) {
        const bc = baseChunks[ci];
        const cc = currChunks[ci];

        if (bc && cc && bc.hash === cc.hash) {
            // Chunks are identical — all rows match positionally
            const len = Math.min(bc.end - bc.start, cc.end - cc.start);
            for (let r = 0; r < len; r++) {
                aligned.push({ type: 'match', baseRi: bc.start + r, currRi: cc.start + r });
            }
            // Handle remaining rows if chunk sizes differ
            for (let r = len; r < bc.end - bc.start; r++) {
                aligned.push({ type: 'deleted', baseRi: bc.start + r });
            }
            for (let r = len; r < cc.end - cc.start; r++) {
                aligned.push({ type: 'added', currRi: cc.start + r });
            }
        } else if (bc && cc) {
            // Chunks differ — run LCS on this block
            const blockBase = baseFP.slice(bc.start, bc.end);
            const blockCurr = currFP.slice(cc.start, cc.end);
            const blockAligned = lcsAlign(blockBase, blockCurr);
            // Remap indices back to global
            for (const row of blockAligned) {
                if (row.type === 'match') {
                    aligned.push({ type: 'match', baseRi: bc.start + row.baseRi, currRi: cc.start + row.currRi });
                } else if (row.type === 'added') {
                    aligned.push({ type: 'added', currRi: cc.start + row.currRi });
                } else {
                    aligned.push({ type: 'deleted', baseRi: bc.start + row.baseRi });
                }
            }
        } else if (bc) {
            // Only in base — all deleted
            for (let r = bc.start; r < bc.end; r++) {
                aligned.push({ type: 'deleted', baseRi: r });
            }
        } else if (cc) {
            // Only in current — all added
            for (let r = cc.start; r < cc.end; r++) {
                aligned.push({ type: 'added', currRi: r });
            }
        }
    }

    return aligned;
}

// ================ Core Diff ================

function getSheetRows(sheet: any): { [key: number]: any } {
    // Handle both ExcelData format (from loadSheets) and raw SheetData format
    if (sheet.rows && typeof sheet.rows === 'object') {
        return sheet.rows;
    }
    return {};
}

function getMaxRowIndex(rows: { [key: number]: any }): number {
    const keys = Object.keys(rows).map(Number).filter(n => !isNaN(n));
    return keys.length > 0 ? Math.max(...keys) : -1;
}

function diffSheet(
    baseSheet: any,
    currSheet: any,
    maxCols: number
): { left: DiffSheetData; right: DiffSheetData; stats: DiffStats } {
    const baseRows = getSheetRows(baseSheet);
    const currRows = getSheetRows(currSheet);
    const baseMaxRi = getMaxRowIndex(baseRows);
    const currMaxRi = getMaxRowIndex(currRows);

    // Build fingerprints
    const baseFP: string[] = [];
    for (let ri = 0; ri <= baseMaxRi; ri++) {
        baseFP.push(rowFingerprint(baseRows[ri], maxCols));
    }
    const currFP: string[] = [];
    for (let ri = 0; ri <= currMaxRi; ri++) {
        currFP.push(rowFingerprint(currRows[ri], maxCols));
    }

    // Align rows using chunked LCS
    const aligned = chunkedAlign(baseFP, currFP);

    // Build output sheet data with highlight styles
    const leftStyles = [STYLE_DELETED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
    const rightStyles = [STYLE_ADDED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
    const DELETED_IDX = 0, MODIFIED_CELL_IDX = 1, EMPTY_IDX = 2, UNCHANGED_CELL_IDX = 3;
    const ADDED_IDX = 0; // In right styles array

    const leftRows: { [key: number]: any } = {};
    const rightRows: { [key: number]: any } = {};
    const stats: DiffStats = { added: 0, deleted: 0, modified: 0, unchanged: 0 };
    const changeRows: number[] = [];

    let outRi = 0;
    for (const row of aligned) {
        if (row.type === 'match') {
            const baseRow = baseRows[row.baseRi];
            const currRow = currRows[row.currRi];
            let rowModified = false;

            const leftCells: any = {};
            const rightCells: any = {};

            // First pass: compare each cell individually
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
                // Second pass: assign white background to unchanged cells that have content
                for (let ci = 0; ci < maxCols; ci++) {
                    if (leftCells[ci] && leftCells[ci].style === undefined && leftCells[ci].text) {
                        leftCells[ci].style = UNCHANGED_CELL_IDX;
                    }
                    if (rightCells[ci] && rightCells[ci].style === undefined && rightCells[ci].text) {
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
        } else if (row.type === 'added') {
            const currRow = currRows[row.currRi];
            const leftCells: any = {};
            const rightCells: any = {};

            // Left side: empty placeholder row
            for (let ci = 0; ci < maxCols; ci++) {
                leftCells[ci] = { text: '', style: EMPTY_IDX };
            }

            // Right side: green highlighted
            for (let ci = 0; ci < maxCols; ci++) {
                const cell = currRow?.cells?.[ci];
                const text = cell?.text != null ? String(cell.text) : '';
                rightCells[ci] = { text, style: ADDED_IDX };
            }

            stats.added++;
            changeRows.push(outRi);
            leftRows[outRi] = { cells: leftCells };
            rightRows[outRi] = { cells: rightCells };
            outRi++;
        } else if (row.type === 'deleted') {
            const baseRow = baseRows[row.baseRi];
            const leftCells: any = {};
            const rightCells: any = {};

            // Left side: red highlighted
            for (let ci = 0; ci < maxCols; ci++) {
                const cell = baseRow?.cells?.[ci];
                const text = cell?.text != null ? String(cell.text) : '';
                leftCells[ci] = { text, style: DELETED_IDX };
            }

            // Right side: empty placeholder row
            for (let ci = 0; ci < maxCols; ci++) {
                rightCells[ci] = { text: '', style: EMPTY_IDX };
            }

            stats.deleted++;
            changeRows.push(outRi);
            leftRows[outRi] = { cells: leftCells };
            rightRows[outRi] = { cells: rightCells };
            outRi++;
        }
    }

    const sheetName = currSheet.name || baseSheet.name || 'Sheet1';
    return {
        left: { name: sheetName, rows: leftRows, styles: leftStyles, maxCols, maxRows: outRi, changeRows },
        right: { name: sheetName, rows: rightRows, styles: rightStyles, maxCols, maxRows: outRi, changeRows },
        stats,
    };
}

// ================ Public API ================

/**
 * Compute the diff between two ExcelData objects (as returned by loadSheets).
 * Returns left/right SheetData arrays ready for rendering in two Spreadsheet instances.
 */
export function computeDiff(base: ExcelData, current: ExcelData): DiffResult {
    const totalStats: DiffStats = { added: 0, deleted: 0, modified: 0, unchanged: 0 };
    const leftSheets: DiffSheetData[] = [];
    const rightSheets: DiffSheetData[] = [];

    // Match sheets by name
    const baseSheetMap = new Map<string, any>();
    for (const sheet of base.sheets) {
        baseSheetMap.set(sheet.name.toLowerCase(), sheet);
    }
    const currSheetMap = new Map<string, any>();
    for (const sheet of current.sheets) {
        currSheetMap.set(sheet.name.toLowerCase(), sheet);
    }

    const maxCols = Math.max(base.maxCols, current.maxCols);

    // Process matched + current-only sheets
    for (const currSheet of current.sheets) {
        const key = currSheet.name.toLowerCase();
        const baseSheet = baseSheetMap.get(key);

        if (baseSheet) {
            // Matched sheet — diff it
            const { left, right, stats } = diffSheet(baseSheet, currSheet, maxCols);
            leftSheets.push(left);
            rightSheets.push(right);
            totalStats.added += stats.added;
            totalStats.deleted += stats.deleted;
            totalStats.modified += stats.modified;
            totalStats.unchanged += stats.unchanged;
            baseSheetMap.delete(key);
        } else {
            // Only in current — all added
            const rows = getSheetRows(currSheet);
            const maxRi = getMaxRowIndex(rows);
            const rightRows: any = {};
            const leftRows: any = {};
            const rightStyles = [STYLE_ADDED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
            const leftStyles = [STYLE_DELETED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
            const changeRows: number[] = [];

            for (let ri = 0; ri <= maxRi; ri++) {
                const row = rows[ri];
                const rightCells: any = {};
                const leftCells: any = {};
                for (let ci = 0; ci < maxCols; ci++) {
                    const cell = row?.cells?.[ci];
                    const text = cell?.text != null ? String(cell.text) : '';
                    rightCells[ci] = { text, style: 0 }; // ADDED
                    leftCells[ci] = { text: '', style: 2 }; // EMPTY
                }
                totalStats.added++;
                rightRows[ri] = { cells: rightCells };
                leftRows[ri] = { cells: leftCells };
                changeRows.push(ri);
            }

            leftSheets.push({ name: currSheet.name, rows: leftRows, styles: leftStyles, maxCols, maxRows: maxRi + 1, changeRows });
            rightSheets.push({ name: currSheet.name, rows: rightRows, styles: rightStyles, maxCols, maxRows: maxRi + 1, changeRows });
        }
    }

    // Process base-only sheets (deleted)
    for (const [, baseSheet] of baseSheetMap) {
        const rows = getSheetRows(baseSheet);
        const maxRi = getMaxRowIndex(rows);
        const leftRows: any = {};
        const rightRows: any = {};
        const leftStyles = [STYLE_DELETED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
        const rightStyles = [STYLE_ADDED, STYLE_MODIFIED_CELL, STYLE_EMPTY_PLACEHOLDER, STYLE_UNCHANGED_IN_MODIFIED_ROW];
        const changeRows: number[] = [];

        for (let ri = 0; ri <= maxRi; ri++) {
            const row = rows[ri];
            const leftCells: any = {};
            const rightCells: any = {};
            for (let ci = 0; ci < maxCols; ci++) {
                const cell = row?.cells?.[ci];
                const text = cell?.text != null ? String(cell.text) : '';
                leftCells[ci] = { text, style: 0 }; // DELETED
                rightCells[ci] = { text: '', style: 2 }; // EMPTY
            }
            totalStats.deleted++;
            leftRows[ri] = { cells: leftCells };
            rightRows[ri] = { cells: rightCells };
            changeRows.push(ri);
        }

        leftSheets.push({ name: baseSheet.name, rows: leftRows, styles: leftStyles, maxCols, maxRows: maxRi + 1, changeRows });
        rightSheets.push({ name: baseSheet.name, rows: rightRows, styles: rightStyles, maxCols, maxRows: maxRi + 1, changeRows });
    }

    return { leftSheets, rightSheets, stats: totalStats };
}
