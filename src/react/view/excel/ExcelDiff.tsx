import { message, Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler } from "../../util/vscode.ts";
import './ExcelDiff.less';
import { loadSheets, detectEncoding } from "./excel_reader.ts";
import { export_xlsx } from "./excel_writer.ts";
import { computeDiff, DiffSheetData } from "./excel_diff.ts";
import Spreadsheet from './x-spreadsheet/index';

const ENCODINGS = [
    'utf-8', 'gbk', 'gb2312', 'gb18030', 'big5', 'shift_jis', 'euc-jp',
    'euc-kr', 'iso-8859-1', 'windows-1252', 'utf-16le', 'utf-16be',
];

type Ref =
    | { kind: 'working' }
    | { kind: 'staged' }
    | { kind: 'head' }
    | { kind: 'commit'; hash: string; message?: string; date?: string }
    | { kind: 'stash'; index: number; message?: string }
    | { kind: 'file'; uri: any }
    | { kind: 'svn-working' }
    | { kind: 'svn-revision'; revision: string; message?: string; date?: string };

function refLabel(r: Ref | null | undefined): string {
    if (!r) return '';
    switch (r.kind) {
        case 'working': case 'svn-working': return 'Working';
        case 'staged': return 'Staged';
        case 'head': return 'HEAD';
        case 'commit': return r.hash.substring(0, 7);
        case 'stash': return `stash@{${r.index}}`;
        case 'file': {
            const p = r.uri?.path || r.uri?.fsPath || '';
            return p.split(/[\\/]/).pop() || 'file';
        }
        case 'svn-revision': return `r${r.revision}`;
        default: return '?';
    }
}

interface BlamePopover {
    sheet: string;
    row: number;
    col: number;
    entry?: { hash: string; author: string; date: string; message: string };
    error?: string;
}

function base64ToBuffer(b64: string): ArrayBuffer {
    const bin = atob(b64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8.buffer;
}

export default function ExcelDiff() {
    const [loading, setLoading] = useState(true);
    const [stats, setStats] = useState<{ added: number; deleted: number; modified: number; unchanged: number } | null>(null);
    const [encoding, setEncoding] = useState('utf-8');
    const [changeIndex, setChangeIndex] = useState(-1);
    const [totalChanges, setTotalChanges] = useState(0);
    const [leftRefState, setLeftRefState] = useState<Ref | null>(null);
    const [rightRefState, setRightRefState] = useState<Ref | null>(null);
    const [blamePopover, setBlamePopover] = useState<BlamePopover | null>(null);

    const leftRef = useRef<any>(null);
    const rightRef = useRef<any>(null);
    const syncingScroll = useRef(false);
    const changeRowsRef = useRef<number[]>([]);
    const diffDataRef = useRef<{ leftData: string; rightData: string; ext: string } | null>(null);
    const sheetsRef = useRef<{ left: DiffSheetData[]; right: DiffSheetData[] } | null>(null);
    const currentBufferRef = useRef<ArrayBuffer | null>(null);
    const workbookRef = useRef<any>(null);

    const renderFromBase64 = useCallback((leftData: string, rightData: string, fileExt: string, enc: string, autoDetect: boolean = false) => {
        setLoading(true);
        const leftContainer = document.getElementById('diff-left');
        const rightContainer = document.getElementById('diff-right');

        setTimeout(() => {
            try {
                const leftBuffer = base64ToBuffer(leftData);
                const rightBuffer = base64ToBuffer(rightData);

                let effectiveEncoding = enc;
                if (autoDetect && fileExt?.match(/csv/i)) {
                    const detected = detectEncoding(rightBuffer);
                    if (detected.confidence >= 0.7) {
                        effectiveEncoding = detected.encoding;
                        setEncoding(effectiveEncoding);
                    }
                }

                currentBufferRef.current = rightBuffer;

                const baseExcel = loadSheets(leftBuffer, fileExt, effectiveEncoding);
                const currentExcel = loadSheets(rightBuffer, fileExt, effectiveEncoding);
                workbookRef.current = currentExcel.workbook || null;
                const diffResult = computeDiff(baseExcel, currentExcel);

                setStats(diffResult.stats);
                sheetsRef.current = { left: diffResult.leftSheets, right: diffResult.rightSheets };

                const rows = diffResult.leftSheets[0]?.changeRows || [];
                changeRowsRef.current = rows;
                setTotalChanges(rows.length);
                setChangeIndex(-1);

                renderSide(leftContainer!, diffResult.leftSheets, 'left');
                renderSide(rightContainer!, diffResult.rightSheets, 'right', true);
                setupScrollSync(leftContainer!, rightContainer!);

                setLoading(false);
            } catch (err) {
                console.error('Diff failed:', err);
                setLoading(false);
            }
        }, 16);
    }, []);

    useEffect(() => {
        handler.on("openDiff", (content: any) => {
            const { leftRef, rightRef, leftData, rightData, ext, encoding: enc } = content;
            diffDataRef.current = { leftData, rightData, ext };
            setLeftRefState(leftRef);
            setRightRefState(rightRef);
            setEncoding(enc || 'utf-8');
            renderFromBase64(leftData, rightData, ext, enc || 'utf-8', true);
        }).on("updateSide", (payload: any) => {
            const { side, ref, data } = payload;
            if (!diffDataRef.current) return;
            if (side === 'left') {
                diffDataRef.current.leftData = data;
                setLeftRefState(ref);
            } else {
                diffDataRef.current.rightData = data;
                setRightRefState(ref);
            }
            const dd = diffDataRef.current;
            renderFromBase64(dd.leftData, dd.rightData, dd.ext, encoding);
        }).on("blameResult", (payload: any) => {
            setBlamePopover({
                sheet: payload.sheet,
                row: payload.row,
                col: payload.col,
                entry: payload.entry || undefined,
                error: payload.error,
            });
        }).on("saveDone", () => {
            message.success({ duration: 1, content: 'Save done' });
            if (diffDataRef.current) {
                const { leftData, rightData, ext } = diffDataRef.current;
                setTimeout(() => renderFromBase64(leftData, rightData, ext, encoding), 300);
            }
        }).emit("init");

        const onKeydown = (e: KeyboardEvent) => {
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyS') {
                e.preventDefault();
                const ss = rightRef.current;
                if (ss && diffDataRef.current) {
                    export_xlsx(ss, diffDataRef.current.ext, workbookRef.current, encoding);
                }
            }
        };
        window.addEventListener('keydown', onKeydown);
        return () => window.removeEventListener('keydown', onKeydown);
    }, [renderFromBase64, encoding]);

    function renderSide(container: HTMLElement, sheets: DiffSheetData[], side: 'left' | 'right', editable: boolean = false) {
        container.innerHTML = '';
        const maxRows = Math.max(...sheets.map(s => s.maxRows));
        const maxCols = Math.max(...sheets.map(s => s.maxCols));

        const spreadSheet = new Spreadsheet(container, {
            mode: editable ? 'edit' : 'read',
            showToolbar: false,
            showBottomBar: sheets.length > 1,
            row: { len: maxRows + 20, height: 30 },
            col: { len: maxCols },
            view: {
                height: () => window.innerHeight - 80,
                width: () => (window.innerWidth / 2) - 2,
            },
        });

        const sheetData = sheets.map(sheet => ({
            name: sheet.name,
            rows: sheet.rows,
            styles: sheet.styles,
        }));

        spreadSheet.loadData(sheetData);

        if (side === 'left') leftRef.current = spreadSheet;
        else rightRef.current = spreadSheet;

        if (side === 'right' && spreadSheet.bottombar) {
            const origSwap = spreadSheet.bottombar.swapFunc;
            spreadSheet.bottombar.swapFunc = (index: number) => {
                origSwap(index);
                const newChangeRows = sheets[index]?.changeRows || [];
                changeRowsRef.current = newChangeRows;
                setTotalChanges(newChangeRows.length);
                setChangeIndex(-1);
                const lss = leftRef.current;
                if (lss && lss.bottombar) {
                    const d = lss.datas[index];
                    if (d) lss.sheet.resetData(d);
                }
            };
        }
    }

    function setupScrollSync(leftEl: HTMLElement, rightEl: HTMLElement) {
        const observeScroll = (source: HTMLElement, target: HTMLElement) => {
            const sourceScrollEls = source.querySelectorAll('.x-spreadsheet-scrollbar');
            const targetScrollEls = target.querySelectorAll('.x-spreadsheet-scrollbar');
            sourceScrollEls.forEach((el, idx) => {
                el.addEventListener('scroll', () => {
                    if (syncingScroll.current) return;
                    syncingScroll.current = true;
                    const targetEl = targetScrollEls[idx] as HTMLElement;
                    if (targetEl) {
                        targetEl.scrollTop = el.scrollTop;
                        targetEl.scrollLeft = el.scrollLeft;
                    }
                    requestAnimationFrame(() => { syncingScroll.current = false; });
                });
            });
        };
        setTimeout(() => {
            observeScroll(leftEl, rightEl);
            observeScroll(rightEl, leftEl);
        }, 200);
    }

    const navigateToRow = useCallback((rowIndex: number) => {
        const ss = rightRef.current;
        if (!ss) return;
        const data = ss.data;
        if (!data) return;

        const { rows, freeze } = data;
        const [fri] = freeze;
        const pixelOffset = rows.sumHeight(fri, rowIndex);

        if (ss.sheet.verticalScrollbar) {
            ss.sheet.verticalScrollbar.move({ top: Math.max(0, pixelOffset - rows.getHeight(rowIndex)) });
        }

        setTimeout(() => {
            ss.sheet.selector.set(rowIndex, 0);
            ss.sheet.table.render();

            const lss = leftRef.current;
            if (lss && lss.sheet.verticalScrollbar) {
                lss.sheet.verticalScrollbar.move({ top: Math.max(0, pixelOffset - rows.getHeight(rowIndex)) });
                setTimeout(() => {
                    lss.sheet.selector.set(rowIndex, 0);
                    lss.sheet.table.render();
                }, 20);
            }
        }, 20);
    }, []);

    const goNextChange = useCallback(() => {
        const rows = changeRowsRef.current;
        if (rows.length === 0) return;
        const next = changeIndex + 1 >= rows.length ? 0 : changeIndex + 1;
        setChangeIndex(next);
        navigateToRow(rows[next]);
    }, [changeIndex, navigateToRow]);

    const goPrevChange = useCallback(() => {
        const rows = changeRowsRef.current;
        if (rows.length === 0) return;
        const prev = changeIndex - 1 < 0 ? rows.length - 1 : changeIndex - 1;
        setChangeIndex(prev);
        navigateToRow(rows[prev]);
    }, [changeIndex, navigateToRow]);

    const handleEncodingChange = useCallback((newEncoding: string) => {
        setEncoding(newEncoding);
        if (diffDataRef.current) {
            const { leftData, rightData, ext } = diffDataRef.current;
            renderFromBase64(leftData, rightData, ext, newEncoding);
        }
    }, [renderFromBase64]);

    const onPickRef = useCallback((side: 'left' | 'right') => {
        handler.emit("pickRef", { side });
    }, []);

    const onBlameClick = useCallback(() => {
        const ss = rightRef.current;
        if (!ss) return;
        const sel = ss.sheet?.selector;
        if (!sel) return;
        const ri = sel.ri;
        const ci = sel.ci;
        if (ri == null || ci == null || ri < 0 || ci < 0) {
            message.warning({ duration: 2, content: 'Select a cell first' });
            return;
        }
        const sheetName = ss.data?.name || 'Sheet1';
        setBlamePopover(null);
        handler.emit("requestBlame", { sheet: sheetName, row: ri, col: ci });
    }, []);

    return (
        <div className="excel-diff-viewer">
            <Spin spinning={loading} fullscreen={true} />
            <div className="excel-diff-toolbar">
                <div className="toolbar-left">
                    <button className="ref-switch" onClick={() => onPickRef('left')} title="Change left side reference">
                        {refLabel(leftRefState)} <span className="ref-caret">▼</span>
                    </button>
                    <span className="ref-arrow">↔</span>
                    <button className="ref-switch" onClick={() => onPickRef('right')} title="Change right side reference">
                        {refLabel(rightRefState)} <span className="ref-caret">▼</span>
                    </button>
                    <div className="legend-items">
                        <span className="legend-item">
                            <span className="legend-color legend-added"></span>Added
                        </span>
                        <span className="legend-item">
                            <span className="legend-color legend-deleted"></span>Deleted
                        </span>
                        <span className="legend-item">
                            <span className="legend-color legend-modified"></span>Modified
                        </span>
                    </div>
                    {stats && (
                        <div className="legend-stats">
                            <span className="stat-added">+{stats.added}</span>
                            <span className="stat-deleted">-{stats.deleted}</span>
                            <span className="stat-modified">~{stats.modified}</span>
                        </div>
                    )}
                </div>
                <div className="toolbar-right">
                    <div className="nav-group">
                        <button className="nav-btn" onClick={goPrevChange} title="Previous Change">&#x25B2;</button>
                        <span className="nav-indicator">
                            {totalChanges > 0 ? `${changeIndex + 1}/${totalChanges}` : '0/0'}
                        </span>
                        <button className="nav-btn" onClick={goNextChange} title="Next Change">&#x25BC;</button>
                    </div>
                    <button className="nav-btn blame-btn" onClick={onBlameClick} title="Blame selected cell">
                        Blame
                    </button>
                    <div className="encoding-group">
                        <select
                            value={encoding}
                            onChange={(e) => handleEncodingChange(e.target.value)}
                            title="File Encoding"
                        >
                            {ENCODINGS.map(enc => (
                                <option key={enc} value={enc}>{enc.toUpperCase()}</option>
                            ))}
                        </select>
                    </div>
                </div>
            </div>
            <div className="excel-diff-panels">
                <div className="excel-diff-panel">
                    <div className="panel-header">BASE ({refLabel(leftRefState)})</div>
                    <div id="diff-left" className="panel-content"></div>
                </div>
                <div className="excel-diff-divider"></div>
                <div className="excel-diff-panel">
                    <div className="panel-header">CURRENT ({refLabel(rightRefState)} - Editable, Ctrl+S to save)</div>
                    <div id="diff-right" className="panel-content"></div>
                </div>
            </div>
            {blamePopover && (
                <div className="blame-popover" onClick={() => setBlamePopover(null)}>
                    {blamePopover.error ? (
                        <span>{blamePopover.error}</span>
                    ) : blamePopover.entry ? (
                        <span>
                            <strong>{blamePopover.entry.hash.substring(0, 7)}</strong>
                            {' · '}{blamePopover.entry.author}
                            {' · '}{blamePopover.entry.date.substring(0, 10)}
                            {' — '}{blamePopover.entry.message}
                        </span>
                    ) : (
                        <span>No blame info found for this cell.</span>
                    )}
                    <span className="blame-close" title="Dismiss"> × </span>
                </div>
            )}
        </div>
    );
}
