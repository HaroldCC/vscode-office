import * as XLSX from 'xlsx/dist/xlsx.mini.min.js';
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
};

function normalizeEncoding(raw: string): string {
    const lower = raw.toLowerCase().replace(/[^a-z0-9-]/g, '');
    return ENCODING_NORMALIZE_MAP[lower] || ENCODING_NORMALIZE_MAP[raw.toLowerCase()] || 'utf-8';
}

export function detectEncoding(buffer: ArrayBuffer): { encoding: string; confidence: number } {
    const bytes = new Uint8Array(buffer);
    const n = bytes.length;

    // 1. BOM 优先 — 命中即 100% 确信
    if (n >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
        return { encoding: 'utf-8', confidence: 1 };
    }
    if (n >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
        return { encoding: 'utf-16le', confidence: 1 };
    }
    if (n >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
        return { encoding: 'utf-16be', confidence: 1 };
    }

    // 取 64KB 样本(比之前 4KB 更可靠,尤其对短行 CSV 前几 KB 是 ASCII 数字的场景)
    const sampleSize = Math.min(65536, n);
    const sample = bytes.subarray(0, sampleSize);

    // 2. 纯 ASCII(无 >= 0x80 字节)→ UTF-8(超集)
    let hasHighByte = false;
    for (let i = 0; i < sampleSize; i++) {
        if (sample[i] >= 0x80) { hasHighByte = true; break; }
    }
    if (!hasHighByte) return { encoding: 'utf-8', confidence: 1 };

    // 3. 严格 UTF-8 校验 — 把字节按 UTF-8 多字节规则扫描,所有合法 → 高度可信 UTF-8
    if (isValidUtf8(sample)) {
        return { encoding: 'utf-8', confidence: 0.95 };
    }

    // 4. jschardet 兜底(它对 GB18030/Big5/Shift_JIS 等亚洲编码识别较好)
    let result;
    try {
        result = jschardet.detect(Buffer.from(sample));
    } catch { result = null; }
    if (result?.encoding && result.confidence >= 0.8) {
        return { encoding: normalizeEncoding(result.encoding), confidence: result.confidence };
    }

    // 5. GBK 启发:中文 Windows CSV 最常见。检测大量 0x81-0xFE 双字节序列
    if (looksLikeGbk(sample)) {
        return { encoding: 'gb18030', confidence: 0.7 };
    }

    // 6. 最后仍不确定,用 jschardet 的结果(即使 confidence 不高)
    if (result?.encoding) {
        return { encoding: normalizeEncoding(result.encoding), confidence: result.confidence };
    }
    return { encoding: 'utf-8', confidence: 0 };
}

/** 严格 UTF-8 字节序列校验。返回 true 当且仅当 sample 完全符合 UTF-8 编码规则。 */
function isValidUtf8(b: Uint8Array): boolean {
    const n = b.length;
    let i = 0;
    while (i < n) {
        const c = b[i];
        if (c < 0x80) { i++; continue; }
        let need: number;
        if ((c & 0xE0) === 0xC0) need = 1;       // 110xxxxx → 2-byte
        else if ((c & 0xF0) === 0xE0) need = 2;  // 1110xxxx → 3-byte
        else if ((c & 0xF8) === 0xF0) need = 3;  // 11110xxx → 4-byte
        else return false;
        if (i + need >= n) {
            // 尾部不完整(可能样本截断);允许只剩最后一个不完整字符
            return i + need < n + need;
        }
        for (let k = 1; k <= need; k++) {
            if ((b[i + k] & 0xC0) !== 0x80) return false;
        }
        i += need + 1;
    }
    return true;
}

/** GBK/GB18030 启发:大部分高位字节出现在 0x81-0xFE 范围且成对(双字节字符) */
function looksLikeGbk(b: Uint8Array): boolean {
    const n = b.length;
    let total = 0, gbkPairs = 0, bad = 0;
    let i = 0;
    while (i < n) {
        const c = b[i];
        if (c < 0x80) { i++; continue; }
        total++;
        if (c >= 0x81 && c <= 0xFE && i + 1 < n) {
            const c2 = b[i + 1];
            // GBK 双字节范围:第二个字节 0x40-0x7E 或 0x80-0xFE
            if ((c2 >= 0x40 && c2 <= 0x7E) || (c2 >= 0x80 && c2 <= 0xFE)) {
                gbkPairs++;
                i += 2;
                continue;
            }
        }
        bad++;
        i++;
    }
    if (total === 0) return false;
    return gbkPairs / total > 0.85 && bad / total < 0.05;
}

/**
 * 检测 CSV 分隔符。在前 50 行内,统计每个候选分隔符的总出现次数,
 * 若领先的分隔符出现次数 ≥ 第二名 × 1.5 且 ≥ 行数,采用之,否则 fallback ','。
 */
export function detectCsvDelimiter(text: string): string {
    const candidates = [',', ';', '\t', '|'];
    const lines = text.split(/\r?\n/, 50).filter(l => l.length > 0);
    if (lines.length === 0) return ',';
    const counts: number[] = candidates.map(d => {
        let total = 0;
        for (const line of lines) {
            for (let i = 0; i < line.length; i++) if (line[i] === d) total++;
        }
        return total;
    });
    const sorted = candidates.map((d, i) => ({ d, n: counts[i] })).sort((a, b) => b.n - a.n);
    const [top, second] = sorted;
    if (top.n >= lines.length && top.n >= (second?.n || 0) * 1.5) {
        return top.d;
    }
    return ',';
}

interface SheetInfo {
    name: string;
    rows: any;
    cols?: { [key: string]: { width: number } };
    merges?: string[];
    styles?: any[];
}

export interface ExcelData {
    sheets: SheetInfo[];
    maxCols: number;
    maxLength?: number;
    workbook?: any; // Keep reference for incremental save
}

const MIN_COL_WIDTH = 70;
const MAX_COL_WIDTH = 300;
const CHAR_WIDTH = 8;
const MAX_ROWS_TO_CHECK = 10;

const calculateColWidth = (rows: any[], colIndex: number): number => {
    let maxLength = 0;
    for (let i = 0; i < Math.min(rows.length, MAX_ROWS_TO_CHECK); i++) {
        const cell = rows[i]?.[colIndex];
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

/**
 * Convert XLSX merge ranges (e.g., {s:{r:0,c:0},e:{r:2,c:3}}) to
 * x-spreadsheet merge format (e.g., "A1:D3")
 */
function convertMerges(wsMerges: any[]): string[] {
    if (!wsMerges || wsMerges.length === 0) return [];
    return wsMerges.map(range => {
        const startCell = XLSX.utils.encode_cell(range.s);
        const endCell = XLSX.utils.encode_cell(range.e);
        return `${startCell}:${endCell}`;
    });
}

/**
 * Get cell merge dimensions for x-spreadsheet cell format: [rowspan-1, colspan-1]
 */
function getMergeDimensions(wsMerges: any[], ri: number, ci: number): [number, number] | null {
    if (!wsMerges) return null;
    for (const range of wsMerges) {
        if (range.s.r === ri && range.s.c === ci) {
            const rowSpan = range.e.r - range.s.r;
            const colSpan = range.e.c - range.s.c;
            if (rowSpan > 0 || colSpan > 0) {
                return [rowSpan, colSpan];
            }
        }
    }
    return null;
}

/**
 * Check if a cell is hidden by a merge (not the top-left anchor).
 */
function isMergedSlave(wsMerges: any[], ri: number, ci: number): boolean {
    if (!wsMerges) return false;
    for (const range of wsMerges) {
        if (ri >= range.s.r && ri <= range.e.r && ci >= range.s.c && ci <= range.e.c) {
            if (ri !== range.s.r || ci !== range.s.c) {
                return true;
            }
        }
    }
    return false;
}

/**
 * Extract column widths from XLSX worksheet.
 */
function extractColWidths(ws: any, maxCols: number): { [key: number]: { width: number } } {
    const cols: { [key: number]: { width: number } } = {};
    const wsCols = ws['!cols'];
    if (wsCols) {
        for (let i = 0; i < Math.min(wsCols.length, maxCols); i++) {
            if (wsCols[i] && wsCols[i].wpx) {
                cols[i] = { width: Math.min(Math.max(wsCols[i].wpx, MIN_COL_WIDTH), MAX_COL_WIDTH) };
            } else if (wsCols[i] && wsCols[i].wch) {
                cols[i] = { width: Math.min(Math.max(wsCols[i].wch * CHAR_WIDTH, MIN_COL_WIDTH), MAX_COL_WIDTH) };
            }
        }
    }
    return cols;
}

const convert = (wb: any) => {
    const sheets: SheetInfo[] = [];
    let maxLength = 0;
    let maxCols = 26;

    wb.SheetNames.forEach((name: string) => {
        const sheet: SheetInfo = { name, rows: {} };
        const ws = wb.Sheets[name];

        // Get sheet range
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        const rowCount = range.e.r + 1;
        const colCount = range.e.c + 1;

        if (maxLength < rowCount) maxLength = rowCount;
        if (colCount > maxCols) maxCols = colCount;

        // Extract merges
        const wsMerges = ws['!merges'] || [];
        sheet.merges = convertMerges(wsMerges);

        // Extract column widths (prefer XLSX metadata, fallback to heuristic)
        const xlsxCols = extractColWidths(ws, colCount);

        // Build rows with full cell data
        const rows: any = {};
        for (let ri = 0; ri <= range.e.r; ri++) {
            const cells: any = {};
            let hasContent = false;

            for (let ci = 0; ci <= range.e.c; ci++) {
                const cellRef = XLSX.utils.encode_cell({ r: ri, c: ci });
                const cell = ws[cellRef];

                // Skip cells hidden by merges
                if (isMergedSlave(wsMerges, ri, ci)) continue;

                if (cell) {
                    const cellData: any = {};

                    // Text value
                    if (cell.t === 'n') {
                        // Number: use formatted value if available, else raw
                        cellData.text = cell.w || String(cell.v);
                    } else if (cell.t === 'b') {
                        cellData.text = cell.v ? 'TRUE' : 'FALSE';
                    } else if (cell.t === 'd') {
                        cellData.text = cell.w || String(cell.v);
                    } else {
                        cellData.text = cell.w || cell.v || '';
                    }

                    // Formula preservation
                    if (cell.f) {
                        cellData.formula = cell.f;
                    }

                    // Merge dimensions
                    const merge = getMergeDimensions(wsMerges, ri, ci);
                    if (merge) {
                        cellData.merge = merge;
                    }

                    cells[ci] = cellData;
                    hasContent = true;
                }
            }

            if (hasContent) {
                rows[ri] = { cells };
            }
        }

        sheet.rows = rows;

        // Use XLSX column widths if available, else calculate heuristically
        if (Object.keys(xlsxCols).length > 0) {
            sheet.cols = xlsxCols;
        } else {
            // Fallback: calculate from content
            const jsonRows = XLSX.utils.sheet_to_json(ws, { raw: false, header: 1 });
            const cols: { [key: number]: { width: number } } = {};
            for (let i = 0; i < colCount; i++) {
                cols[i] = { width: calculateColWidth(jsonRows, i) };
            }
            sheet.cols = cols;
        }

        sheets.push(sheet);
    });

    return { sheets, maxLength, maxCols, workbook: wb };
};

export function loadSheets(buffer: ArrayBuffer, ext: string, encoding: string = 'utf-8'): ExcelData {
    const ab = new Uint8Array(buffer).buffer;
    const options: any = { type: "array", cellFormula: true, cellNF: true };

    let wb;
    if (ext.toLowerCase() === ".csv") {
        const text = new TextDecoder(encoding).decode(ab);
        const fs = detectCsvDelimiter(text);
        wb = XLSX.read(text, { type: "string", raw: true, FS: fs });
    } else {
        wb = XLSX.read(ab, options);
    }

    return convert(wb);
}

const LAZY_LOAD_THRESHOLD = 10 * 1024 * 1024; // 10MB

export function shouldLazyLoad(bufferSize: number): boolean {
    return bufferSize > LAZY_LOAD_THRESHOLD;
}

export function loadSheetNames(buffer: ArrayBuffer, ext: string): string[] {
    if (ext.toLowerCase() === '.csv') {
        return ['Sheet1'];
    }
    const ab = new Uint8Array(buffer).buffer;
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
        const fs = detectCsvDelimiter(text);
        wb = XLSX.read(text, { type: 'string', raw: true, FS: fs });
    } else {
        wb = XLSX.read(ab, { type: 'array', sheets: [sheetName], cellFormula: true, cellNF: true });
    }

    return convert(wb);
}
