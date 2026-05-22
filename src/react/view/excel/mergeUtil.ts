export type CellStatus =
    | 'unchanged'
    | 'auto-ours'
    | 'auto-theirs'
    | 'auto-same'
    | 'conflict'
    | 'resolved-ours'
    | 'resolved-theirs'
    | 'resolved-both'
    | 'resolved-manual';

export interface MergedSheet {
    name: string;
    base: string[][];
    ours: string[][];
    theirs: string[][];
    result: string[][];
    status: CellStatus[][];
    rows: number;
    cols: number;
}

function cell(grid: string[][], r: number, c: number): string {
    const row = grid[r];
    if (!row) return '';
    const v = row[c];
    return v == null ? '' : String(v);
}

function classify(b: string, o: string, t: string): { status: CellStatus; value: string } {
    if (b === o && o === t) return { status: 'unchanged', value: o };
    if (o === t) return { status: 'auto-same', value: o };
    if (b === t) return { status: 'auto-ours', value: o };
    if (b === o) return { status: 'auto-theirs', value: t };
    return { status: 'conflict', value: o };
}

export function mergeSheet(name: string, base: string[][], ours: string[][], theirs: string[][]): MergedSheet {
    const rows = Math.max(base.length, ours.length, theirs.length);
    let cols = 0;
    for (const grid of [base, ours, theirs]) {
        for (const r of grid) if (r && r.length > cols) cols = r.length;
    }
    const result: string[][] = [];
    const status: CellStatus[][] = [];
    for (let r = 0; r < rows; r++) {
        const resultRow: string[] = [];
        const statusRow: CellStatus[] = [];
        for (let c = 0; c < cols; c++) {
            const b = cell(base, r, c);
            const o = cell(ours, r, c);
            const t = cell(theirs, r, c);
            const { status: s, value } = classify(b, o, t);
            resultRow.push(value);
            statusRow.push(s);
        }
        result.push(resultRow);
        status.push(statusRow);
    }
    return { name, base, ours, theirs, result, status, rows, cols };
}

export function applyResolution(sheet: MergedSheet, row: number, col: number, choice: 'ours' | 'theirs' | 'both' | string): void {
    if (!sheet.status[row] || !sheet.result[row]) return;
    const o = cell(sheet.ours, row, col);
    const t = cell(sheet.theirs, row, col);
    let value = sheet.result[row][col];
    let newStatus: CellStatus = sheet.status[row][col];
    if (choice === 'ours') { value = o; newStatus = 'resolved-ours'; }
    else if (choice === 'theirs') { value = t; newStatus = 'resolved-theirs'; }
    else if (choice === 'both') { value = `${o} | ${t}`; newStatus = 'resolved-both'; }
    else { value = choice; newStatus = 'resolved-manual'; }
    sheet.result[row][col] = value;
    sheet.status[row][col] = newStatus;
}

export function countUnresolved(sheet: MergedSheet): number {
    let n = 0;
    for (const row of sheet.status) for (const s of row) if (s === 'conflict') n++;
    return n;
}
