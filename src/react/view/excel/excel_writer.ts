import { handler } from "../../util/vscode";
import * as XLSX from 'xlsx/dist/xlsx.mini.min.js';
import Spreadsheet from './x-spreadsheet/index';

/**
 * Convert x-spreadsheet sheet data to XLSX worksheet.
 * Preserves formulas and merge information.
 */
function dataToSheet(xws: any): any {
    const aoa: any[][] = [[]];
    const rowobj = xws.rows;
    const merges: any[] = [];

    for (let ri = 0; ri < (rowobj.len || 1000); ri++) {
        const row = rowobj[ri];
        if (!row) continue;
        aoa[ri] = aoa[ri] || [];

        Object.keys(row.cells).forEach(function (k) {
            const idx = +k;
            if (isNaN(idx)) return;
            const cell = row.cells[k];

            // Preserve formula if present
            if (cell.formula) {
                aoa[ri][idx] = { f: cell.formula, t: 's', v: cell.text || '' };
            } else {
                aoa[ri][idx] = cell.text;
            }

            // Collect merges
            if (cell.merge) {
                const [rowSpan, colSpan] = cell.merge;
                if (rowSpan > 0 || colSpan > 0) {
                    merges.push({
                        s: { r: ri, c: idx },
                        e: { r: ri + rowSpan, c: idx + colSpan }
                    });
                }
            }
        });
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Re-apply formulas (aoa_to_sheet doesn't handle formula objects well)
    for (let ri = 0; ri < aoa.length; ri++) {
        if (!aoa[ri]) continue;
        for (let ci = 0; ci < aoa[ri].length; ci++) {
            const val = aoa[ri][ci];
            if (val && typeof val === 'object' && val.f) {
                const cellRef = XLSX.utils.encode_cell({ r: ri, c: ci });
                ws[cellRef] = { t: 's', v: val.v || '', f: val.f };
            }
        }
    }

    // Apply merges
    if (merges.length > 0) {
        ws['!merges'] = merges;
    }

    return ws;
}

function xtos(sdata: any[]) {
    const out = XLSX.utils.book_new();
    sdata.forEach(function (xws) {
        const ws = dataToSheet(xws);
        XLSX.utils.book_append_sheet(out, ws, xws.name);
    });
    return out;
}

/**
 * Incremental save: if we have the original workbook, update cells in-place
 * to preserve styles, charts, images, and other metadata that we can't render.
 */
function incrementalSave(spreadSheet: Spreadsheet, originalWb: any, extName: string) {
    const sdata = spreadSheet.getData();

    sdata.forEach((sheetData: any, idx: number) => {
        const sheetName = sheetData.name || originalWb.SheetNames[idx];
        const ws = originalWb.Sheets[sheetName];
        if (!ws) return;

        const rowobj = sheetData.rows;
        for (let ri = 0; ri < (rowobj.len || 1000); ri++) {
            const row = rowobj[ri];
            if (!row) continue;

            Object.keys(row.cells).forEach(function (k) {
                const ci = +k;
                if (isNaN(ci)) return;
                const cell = row.cells[k];
                const cellRef = XLSX.utils.encode_cell({ r: ri, c: ci });

                if (cell.formula) {
                    ws[cellRef] = { t: 's', v: cell.text || '', f: cell.formula };
                } else if (cell.text != null && cell.text !== '') {
                    // Preserve existing cell type if possible
                    const existing = ws[cellRef];
                    if (existing) {
                        existing.v = cell.text;
                        if (existing.w) existing.w = cell.text;
                    } else {
                        ws[cellRef] = { t: 's', v: cell.text };
                    }
                }

                // Update merges
                if (cell.merge) {
                    const [rowSpan, colSpan] = cell.merge;
                    if (!ws['!merges']) ws['!merges'] = [];
                    // Remove old merge for this cell if exists
                    ws['!merges'] = ws['!merges'].filter((m: any) =>
                        !(m.s.r === ri && m.s.c === ci)
                    );
                    if (rowSpan > 0 || colSpan > 0) {
                        ws['!merges'].push({
                            s: { r: ri, c: ci },
                            e: { r: ri + rowSpan, c: ci + colSpan }
                        });
                    }
                }
            });
        }
    });

    const buffer = XLSX.write(originalWb, { bookType: extName, type: "array" });
    const array = [...new Uint8Array(buffer)];
    handler.emit('save', array);
}

/**
 * Export spreadsheet data back to file format.
 * Uses incremental save when original workbook is available (preserves more metadata).
 * Falls back to full rebuild otherwise.
 * @param encoding - For CSV files, encode output in the same encoding as the source.
 */
export function export_xlsx(spreadSheet: Spreadsheet, extName: string, originalWb?: any, encoding?: string) {
    extName = extName.replace('.', '');
    if (extName === 'xlsx' || extName === 'xls' || extName === 'ods') {
        if (originalWb) {
            incrementalSave(spreadSheet, originalWb, extName);
        } else {
            const new_wb = xtos(spreadSheet.getData());
            const buffer = XLSX.write(new_wb, { bookType: extName, type: "array" });
            const array = [...new Uint8Array(buffer)];
            handler.emit('save', array);
        }
    } else if (extName === "csv") {
        const csvContent = XLSX.utils.sheet_to_csv(dataToSheet(spreadSheet.getData()[0]));
        if (encoding && encoding !== 'utf-8') {
            handler.emit('save', { text: csvContent, encoding });
        } else {
            handler.emit('save', csvContent);
        }
    }
}
