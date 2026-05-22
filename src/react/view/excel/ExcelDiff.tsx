import { Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler } from "../../util/vscode.ts";
import './ExcelDiff.less';
import { loadSheets } from "./excel_reader.ts";
import { computeDiff, DiffSheetData } from "./excel_diff.ts";
import Spreadsheet from './x-spreadsheet/index';

const ENCODINGS = [
    'utf-8', 'gbk', 'gb2312', 'gb18030', 'big5', 'shift_jis', 'euc-jp',
    'euc-kr', 'iso-8859-1', 'windows-1252', 'utf-16le', 'utf-16be',
];

export default function ExcelDiff() {
    const [loading, setLoading] = useState(true);
    const [stats, setStats] = useState<{ added: number; deleted: number; modified: number; unchanged: number } | null>(null);
    const [encoding, setEncoding] = useState('utf-8');
    const [changeIndex, setChangeIndex] = useState(-1);
    const [totalChanges, setTotalChanges] = useState(0);

    const leftRef = useRef<any>(null);
    const rightRef = useRef<any>(null);
    const syncingScroll = useRef(false);
    const changeRowsRef = useRef<number[]>([]);
    const diffDataRef = useRef<{ currentPath: string; baseData: string; ext: string } | null>(null);
    const sheetsRef = useRef<{ left: DiffSheetData[]; right: DiffSheetData[] } | null>(null);

    const renderDiff = useCallback((currentPath: string, baseData: string, ext: string, enc: string) => {
        setLoading(true);
        const leftContainer = document.getElementById('diff-left');
        const rightContainer = document.getElementById('diff-right');

        fetch(currentPath)
            .then(res => res.arrayBuffer())
            .then(currentBuffer => {
                const binaryStr = atob(baseData);
                const bytes = new Uint8Array(binaryStr.length);
                for (let i = 0; i < binaryStr.length; i++) {
                    bytes[i] = binaryStr.charCodeAt(i);
                }
                const baseBuffer = bytes.buffer;

                const baseExcel = loadSheets(baseBuffer, ext, enc);
                const currentExcel = loadSheets(currentBuffer, ext, enc);
                const diffResult = computeDiff(baseExcel, currentExcel);

                setStats(diffResult.stats);
                sheetsRef.current = { left: diffResult.leftSheets, right: diffResult.rightSheets };

                // Collect all change rows from current sheet (first sheet)
                const rows = diffResult.leftSheets[0]?.changeRows || [];
                changeRowsRef.current = rows;
                setTotalChanges(rows.length);
                setChangeIndex(-1);

                renderSide(leftContainer!, diffResult.leftSheets, 'left');
                renderSide(rightContainer!, diffResult.rightSheets, 'right');
                setupScrollSync(leftContainer!, rightContainer!);

                setLoading(false);
            })
            .catch(err => {
                console.error('Diff failed:', err);
                setLoading(false);
            });
    }, []);

    useEffect(() => {
        handler.on("openDiff", ({ currentPath, baseData, ext, encoding: enc }) => {
            diffDataRef.current = { currentPath, baseData, ext };
            setEncoding(enc || 'utf-8');
            renderDiff(currentPath, baseData, ext, enc || 'utf-8');
        }).emit("init");
    }, [renderDiff]);

    function renderSide(container: HTMLElement, sheets: DiffSheetData[], side: 'left' | 'right') {
        container.innerHTML = '';
        const maxRows = Math.max(...sheets.map(s => s.maxRows));
        const maxCols = Math.max(...sheets.map(s => s.maxCols));

        const spreadSheet = new Spreadsheet(container, {
            mode: 'read',
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

        // Listen for sheet tab switch (right side drives navigation update)
        if (side === 'right' && spreadSheet.bottombar) {
            const origSwap = spreadSheet.bottombar.swapFunc;
            spreadSheet.bottombar.swapFunc = (index: number) => {
                origSwap(index);
                // Update changeRows for the new active sheet
                const newChangeRows = sheets[index]?.changeRows || [];
                changeRowsRef.current = newChangeRows;
                setTotalChanges(newChangeRows.length);
                setChangeIndex(-1);
                // Sync left side tab switch
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

        // Calculate pixel offset for the target row
        const { rows, freeze } = data;
        const [fri] = freeze;
        const pixelOffset = rows.sumHeight(fri, rowIndex);

        // Drive the scrollbar to move — this triggers the full scroll pipeline
        if (ss.sheet.verticalScrollbar) {
            ss.sheet.verticalScrollbar.move({ top: Math.max(0, pixelOffset - rows.getHeight(rowIndex)) });
        }

        // Select the cell
        setTimeout(() => {
            ss.sheet.selector.set(rowIndex, 0);
            ss.sheet.table.render();

            // Sync left side
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
            const { currentPath, baseData, ext } = diffDataRef.current;
            renderDiff(currentPath, baseData, ext, newEncoding);
        }
    }, [renderDiff]);

    return (
        <div className="excel-diff-viewer">
            <Spin spinning={loading} fullscreen={true} />
            <div className="excel-diff-toolbar">
                <div className="toolbar-left">
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
                    <div className="panel-header">BASE (HEAD)</div>
                    <div id="diff-left" className="panel-content"></div>
                </div>
                <div className="excel-diff-divider"></div>
                <div className="excel-diff-panel">
                    <div className="panel-header">CURRENT (Working)</div>
                    <div id="diff-right" className="panel-content"></div>
                </div>
            </div>
        </div>
    );
}
