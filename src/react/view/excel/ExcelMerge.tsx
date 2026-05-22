import { message, Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler } from "../../util/vscode.ts";
import './ExcelMerge.less';
import { loadSheets } from "./excel_reader.ts";
import { export_xlsx } from "./excel_writer.ts";
import Spreadsheet from './x-spreadsheet/index';
import { mergeSheet, applyResolution, countUnresolved, MergedSheet, CellStatus } from "./mergeUtil.ts";

const STATUS_BG: Record<CellStatus, string | undefined> = {
    'unchanged': undefined,
    'auto-ours': '#d4edda',
    'auto-theirs': '#cce5ff',
    'auto-same': '#d4edda',
    'conflict': '#fff3cd',
    'resolved-ours': '#d4edda',
    'resolved-theirs': '#cce5ff',
    'resolved-both': '#e2d9f3',
    'resolved-manual': '#d1ecf1',
};

function base64ToBuffer(b64: string): ArrayBuffer {
    const bin = atob(b64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8.buffer;
}

function sheetInfoToGrid(sheet: any): string[][] {
    const grid: string[][] = [];
    const rows = sheet?.rows || {};
    const rowKeys = Object.keys(rows).filter(k => k !== 'len').map(Number).filter(n => !isNaN(n));
    const maxRow = rowKeys.length ? Math.max(...rowKeys) : -1;
    for (let r = 0; r <= maxRow; r++) {
        const rowObj = rows[r];
        const cells = rowObj?.cells || {};
        const cellKeys = Object.keys(cells).map(Number).filter(n => !isNaN(n));
        const maxCol = cellKeys.length ? Math.max(...cellKeys) : -1;
        const arr: string[] = [];
        for (let c = 0; c <= maxCol; c++) {
            const cv = cells[c]?.text;
            arr.push(cv == null ? '' : String(cv));
        }
        grid.push(arr);
    }
    return grid;
}

function gridToSheetData(name: string, grid: string[][], statuses?: CellStatus[][]): any {
    const rows: any = {};
    const styleSet = new Map<string, number>();
    const styles: any[] = [];
    for (let r = 0; r < grid.length; r++) {
        const cells: any = {};
        for (let c = 0; c < grid[r].length; c++) {
            const cell: any = { text: grid[r][c] };
            const s = statuses?.[r]?.[c];
            if (s) {
                const bg = STATUS_BG[s];
                if (bg) {
                    let idx = styleSet.get(bg);
                    if (idx === undefined) {
                        idx = styles.length;
                        styles.push({ bgcolor: bg });
                        styleSet.set(bg, idx);
                    }
                    cell.style = idx;
                }
            }
            cells[c] = cell;
        }
        rows[r] = { cells };
    }
    return { name, rows, styles };
}

export default function ExcelMerge() {
    const [loading, setLoading] = useState(true);
    const [sheets, setSheets] = useState<MergedSheet[]>([]);
    const [activeSheet, setActiveSheet] = useState(0);
    const [unresolved, setUnresolved] = useState(0);
    const [autoResolved, setAutoResolved] = useState(0);
    const [conflictIndex, setConflictIndex] = useState(-1);
    const [relPath, setRelPath] = useState('');

    const resultRef = useRef<any>(null);
    const baseRef = useRef<any>(null);
    const oursRef = useRef<any>(null);
    const theirsRef = useRef<any>(null);
    const extRef = useRef<string>('.xlsx');
    const encodingRef = useRef<string>('utf-8');
    const workbookRef = useRef<any>(null);
    const sheetsRef = useRef<MergedSheet[]>([]);

    const computeStats = useCallback((mergedAll: MergedSheet[]) => {
        let un = 0, ar = 0;
        for (const ms of mergedAll) {
            for (const row of ms.status) {
                for (const s of row) {
                    if (s === 'conflict') un++;
                    else if (s === 'auto-ours' || s === 'auto-theirs' || s === 'auto-same') ar++;
                }
            }
        }
        setUnresolved(un);
        setAutoResolved(ar);
    }, []);

    const renderAll = useCallback(() => {
        const ms = sheetsRef.current[activeSheet];
        if (!ms) return;
        renderGrid('merge-base', ms.base, ms.name, baseRef, undefined, false);
        renderGrid('merge-ours', ms.ours, ms.name, oursRef, ms.status.map(row => row.map(s =>
            s === 'auto-ours' || s === 'resolved-ours' ? 'auto-ours' : s === 'conflict' ? 'conflict' : 'unchanged' as CellStatus
        )), false);
        renderGrid('merge-theirs', ms.theirs, ms.name, theirsRef, ms.status.map(row => row.map(s =>
            s === 'auto-theirs' || s === 'resolved-theirs' ? 'auto-theirs' : s === 'conflict' ? 'conflict' : 'unchanged' as CellStatus
        )), false);
        renderGrid('merge-result', ms.result, ms.name, resultRef, ms.status, true);
    }, [activeSheet]);

    function renderGrid(elId: string, grid: string[][], name: string, ref: any, statuses: CellStatus[][] | undefined, editable: boolean) {
        const container = document.getElementById(elId);
        if (!container) return;
        container.innerHTML = '';
        const maxRows = grid.length;
        const maxCols = Math.max(0, ...grid.map(r => r.length));
        const ss = new Spreadsheet(container, {
            mode: editable ? 'edit' : 'read',
            showToolbar: false,
            showBottomBar: false,
            row: { len: maxRows + 20, height: 28 },
            col: { len: Math.max(maxCols, 8) },
            view: {
                height: () => (window.innerHeight / 2) - 40,
                width: () => (window.innerWidth / 3) - 4,
            },
        });
        ss.loadData([gridToSheetData(name, grid, statuses)]);
        ref.current = ss;
    }

    useEffect(() => {
        handler.on('openMerge', (content: any) => {
            const { base, ours, theirs, ext, encoding, relPath: rp } = content;
            extRef.current = ext;
            encodingRef.current = encoding || 'utf-8';
            setRelPath(rp);
            setLoading(true);
            setTimeout(() => {
                try {
                    const baseExcel = base ? loadSheets(base64ToBuffer(base), ext, encoding) : { sheets: [], maxCols: 0 } as any;
                    const oursExcel = loadSheets(base64ToBuffer(ours), ext, encoding);
                    const theirsExcel = loadSheets(base64ToBuffer(theirs), ext, encoding);
                    workbookRef.current = oursExcel.workbook || null;

                    const sheetNames = Array.from(new Set([
                        ...oursExcel.sheets.map((s: any) => s.name),
                        ...theirsExcel.sheets.map((s: any) => s.name),
                    ]));
                    const merged: MergedSheet[] = [];
                    for (const sn of sheetNames) {
                        const b = baseExcel.sheets.find((s: any) => s.name === sn);
                        const o = oursExcel.sheets.find((s: any) => s.name === sn);
                        const t = theirsExcel.sheets.find((s: any) => s.name === sn);
                        const baseGrid = b ? sheetInfoToGrid(b) : [];
                        const oursGrid = o ? sheetInfoToGrid(o) : [];
                        const theirsGrid = t ? sheetInfoToGrid(t) : [];
                        merged.push(mergeSheet(sn, baseGrid, oursGrid, theirsGrid));
                    }
                    sheetsRef.current = merged;
                    setSheets(merged);
                    computeStats(merged);
                    setConflictIndex(-1);
                    setLoading(false);
                    setTimeout(() => renderAll(), 30);
                } catch (err) {
                    console.error('Merge load failed:', err);
                    message.error({ duration: 3, content: 'Failed to load merge data' });
                    setLoading(false);
                }
            }, 16);
        }).on('resolveDone', () => {
            message.success({ duration: 2, content: 'Marked as resolved (git add)' });
        }).emit('init');
    }, [renderAll, computeStats]);

    useEffect(() => {
        if (!loading) renderAll();
    }, [activeSheet, loading, renderAll]);

    const takeOurs = useCallback(() => applyToSelection('ours'), []);
    const takeTheirs = useCallback(() => applyToSelection('theirs'), []);
    const takeBoth = useCallback(() => applyToSelection('both'), []);

    function applyToSelection(choice: 'ours' | 'theirs' | 'both') {
        const ms = sheetsRef.current[activeSheet];
        if (!ms) return;
        const ss = resultRef.current;
        const sel = ss?.sheet?.selector;
        if (!sel || sel.ri == null || sel.ci == null || sel.ri < 0 || sel.ci < 0) {
            message.warning({ duration: 2, content: 'Select a cell in RESULT first' });
            return;
        }
        applyResolution(ms, sel.ri, sel.ci, choice);
        computeStats(sheetsRef.current);
        renderAll();
    }

    const nextConflict = useCallback(() => {
        const ms = sheetsRef.current[activeSheet];
        if (!ms) return;
        // 重新扫描未解决冲突(因为已被解决的已不再是 conflict status)
        const positions: { r: number; c: number }[] = [];
        for (let r = 0; r < ms.status.length; r++) {
            for (let c = 0; c < ms.status[r].length; c++) {
                if (ms.status[r][c] === 'conflict') positions.push({ r, c });
            }
        }
        if (positions.length === 0) {
            message.success({ duration: 2, content: 'No more unresolved conflicts on this sheet' });
            return;
        }
        const next = conflictIndex + 1 >= positions.length ? 0 : conflictIndex + 1;
        setConflictIndex(next);
        const pos = positions[next];
        const ss = resultRef.current;
        if (ss?.sheet?.selector) {
            ss.sheet.selector.set(pos.r, pos.c);
            ss.sheet.table.render();
        }
    }, [activeSheet, conflictIndex]);

    const markResolved = useCallback(() => {
        const all = sheetsRef.current;
        const stillUnresolved = all.reduce((acc, s) => acc + countUnresolved(s), 0);
        if (stillUnresolved > 0) {
            message.warning({ duration: 3, content: `${stillUnresolved} conflicts remain unresolved` });
            return;
        }
        // 把 result 写回 xlsx/csv
        const ss = resultRef.current;
        if (!ss) return;
        // 用 result 数据组装 workbook,通过 export_xlsx 写出 → 它内部走 message back
        // 但 export_xlsx 走 handler.emit('save'),我们要走 'markResolved' 走 git add
        // 因此手动构造 buffer 数组发回 ext 端
        const ext = extRef.current;
        if (ext.toLowerCase() === '.csv') {
            const ms = all[0];
            const text = ms.result.map(row => row.map(cellCsvEscape).join(',')).join('\n');
            handler.emit('markResolved', { content: { text, encoding: encodingRef.current } });
        } else {
            // 借助 export_xlsx 把 ss 数据序列化为 xlsx 数组 buffer,但它直接调用 emit('save')
            // 我们覆写:临时改 handler.emit hook 不可行,直接构造 workbook by XLSX
            const XLSX = (window as any).XLSX || require('xlsx/dist/xlsx.mini.min.js');
            const wb = XLSX.utils.book_new();
            for (const ms of all) {
                const aoa = ms.result;
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                XLSX.utils.book_append_sheet(wb, ws, ms.name);
            }
            const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
            handler.emit('markResolved', { content: Array.from(new Uint8Array(out)) });
        }
    }, []);

    function cellCsvEscape(v: string): string {
        if (v == null) return '';
        if (/[",\n]/.test(v)) return `"${v.replace(/"/g, '""')}"`;
        return v;
    }

    return (
        <div className="excel-merge-viewer">
            <Spin spinning={loading} fullscreen={true} />
            <div className="excel-merge-toolbar">
                <div className="toolbar-left">
                    <span className="merge-title">Merge: {relPath}</span>
                    {sheets.length > 1 && (
                        <select value={activeSheet} onChange={e => setActiveSheet(Number(e.target.value))} className="sheet-select">
                            {sheets.map((s, i) => (
                                <option key={i} value={i}>{s.name}</option>
                            ))}
                        </select>
                    )}
                </div>
                <div className="toolbar-center">
                    <span className="merge-stat stat-conflict">{unresolved} unresolved</span>
                    <span className="merge-stat stat-auto">{autoResolved} auto-resolved</span>
                </div>
                <div className="toolbar-right">
                    <button className="merge-btn" onClick={nextConflict} title="Next conflict">Next ▼</button>
                    <button className="merge-btn ours-btn" onClick={takeOurs} title="Use OURS for selected cell">Take Ours</button>
                    <button className="merge-btn theirs-btn" onClick={takeTheirs} title="Use THEIRS for selected cell">Take Theirs</button>
                    <button className="merge-btn both-btn" onClick={takeBoth} title="Concatenate ours and theirs">Take Both</button>
                    <button className="merge-btn resolved-btn" onClick={markResolved} title="Write result and git add">Mark Resolved</button>
                </div>
            </div>
            <div className="excel-merge-panels three-col">
                <div className="merge-panel">
                    <div className="panel-header">BASE (common ancestor)</div>
                    <div id="merge-base" className="panel-content"></div>
                </div>
                <div className="merge-panel">
                    <div className="panel-header ours-header">OURS (current branch)</div>
                    <div id="merge-ours" className="panel-content"></div>
                </div>
                <div className="merge-panel">
                    <div className="panel-header theirs-header">THEIRS (incoming)</div>
                    <div id="merge-theirs" className="panel-content"></div>
                </div>
            </div>
            <div className="excel-merge-result">
                <div className="panel-header result-header">RESULT (editable — what will be written)</div>
                <div id="merge-result" className="panel-content"></div>
            </div>
        </div>
    );
}
