import * as XLSX from 'xlsx/dist/xlsx.mini.min.js';

interface SheetInfo {
    name: string;
    rows: any[];
    cols?: { [key: string]: { width: number } };
}

export interface ExcelData {
    sheets: SheetInfo[];
    maxCols: number;
    maxLength?: number;
}

const MIN_COL_WIDTH = 70;
const MAX_COL_WIDTH = 300;
const CHAR_WIDTH = 8;
const MAX_ROWS_TO_CHECK = 10;

const calculateColWidth = (rows: any[], colIndex: number): number => {
    let maxLength = 0;
    for (let i = 0; i < Math.min(rows.length, MAX_ROWS_TO_CHECK); i++) {
        const cell = rows[i][colIndex];
        if (cell) {
            const length = String(cell).length;
            if (length > maxLength) {
                maxLength = length;
            }
        }
    }
    const width = maxLength * CHAR_WIDTH;
    return Math.min(Math.max(width, MIN_COL_WIDTH), MAX_COL_WIDTH);
};

const convert = wb => {
    const sheets: SheetInfo[] = [];
    let maxLength = 0;
    let maxCols = 26;
    wb.SheetNames.forEach(name => {
        const sheet: SheetInfo = { name, rows: [] };
        const ws = wb.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(ws, { raw: false, header: 1 });
        if (maxLength < rows.length) maxLength = rows.length

        // 计算列宽
        const cols = {};
        for (let i = 0; i < rows[0]?.length || 0; i++) {
            const width = calculateColWidth(rows, i);
            cols[i] = { width };
        }
        sheet.cols = cols;

        sheet.rows = rows.reduce((map, row, i) => {
            const cells = row.reduce((colMap, column, j) => {
                colMap[j] = { text: column }
                return colMap
            }, {});
            map[i] = { cells }
            const colLen = Object.keys(cells).length;
            if (colLen > maxCols) {
                maxCols = colLen;
            }
            return map
        }, {})

        sheets.push(sheet);
    });
    return { sheets, maxLength, maxCols };
};

export function loadSheets(buffer: ArrayBuffer, ext: string, encoding: string = 'utf-8'): ExcelData {
    const ab = new Uint8Array(buffer).buffer
    const wb = ext.toLowerCase() == ".csv" ? XLSX.read(new TextDecoder(encoding).decode(ab), { type: "string", raw: true }) : XLSX.read(ab, { type: "array" });
    return convert(wb);
}