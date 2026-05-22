import { DiffSheetData } from './excel_diff';

/**
 * 把一个 sheet 的 rows 按变更行索引过滤出来。保留 row 0(表头)与所有 changeRows。
 */
export function filterChangeRows(sheet: DiffSheetData): DiffSheetData {
    const keep = new Set<number>([0, ...sheet.changeRows]);
    const newRows: DiffSheetData['rows'] = {};
    let newIdx = 0;
    const mapping: number[] = [];
    const sortedRows = Object.keys(sheet.rows).map(Number).filter(n => keep.has(n)).sort((a, b) => a - b);
    for (const oldIdx of sortedRows) {
        newRows[newIdx] = sheet.rows[oldIdx];
        mapping.push(oldIdx);
        newIdx++;
    }
    return {
        ...sheet,
        rows: newRows,
        maxRows: newIdx,
        changeRows: sheet.changeRows.map(orig => mapping.indexOf(orig)).filter(i => i >= 0),
    };
}

/**
 * 把双端 leftSheets / rightSheets 同时过滤(以左侧 changeRows 为准并合并右侧)
 */
export function filterDiff(leftSheets: DiffSheetData[], rightSheets: DiffSheetData[]): {
    left: DiffSheetData[]; right: DiffSheetData[];
} {
    const left: DiffSheetData[] = [];
    const right: DiffSheetData[] = [];
    for (let i = 0; i < leftSheets.length; i++) {
        const l = leftSheets[i];
        const r = rightSheets[i];
        if (!l || !r) continue;
        const combinedChange = new Set<number>([...l.changeRows, ...(r?.changeRows || [])]);
        const keep = new Set<number>([0, ...combinedChange]);
        const projectOne = (sheet: DiffSheetData): DiffSheetData => {
            const newRows: DiffSheetData['rows'] = {};
            const sortedRows = Object.keys(sheet.rows).map(Number).filter(n => keep.has(n)).sort((a, b) => a - b);
            let idx = 0;
            for (const oldIdx of sortedRows) {
                newRows[idx] = sheet.rows[oldIdx];
                idx++;
            }
            return { ...sheet, rows: newRows, maxRows: idx, changeRows: Array.from({ length: idx }, (_, i) => i).slice(1) };
        };
        left.push(projectOne(l));
        right.push(projectOne(r));
    }
    return { left, right };
}
