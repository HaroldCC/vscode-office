import { message, Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler } from "../../util/vscode.ts";
import VSCodeLogo from "../vscode.tsx";
import './Excel.less';
import { loadSheets, detectEncoding, shouldLazyLoad, loadSheetNames, loadSingleSheet } from "./excel_reader.ts";
import { export_xlsx } from "./excel_writer.ts";
import Spreadsheet from './x-spreadsheet/index';

const DIRTY_MARKER = ' ●'; // ● symbol for dirty indicator

interface SearchMatch {
    ri: number;
    ci: number;
    sheetIndex: number;
}

export default function Excel() {
    const [loading, setLoading] = useState(true)
    const [searchVisible, setSearchVisible] = useState(false)
    const [searchText, setSearchText] = useState('')
    const [matchCount, setMatchCount] = useState(0)
    const [currentMatch, setCurrentMatch] = useState(0)
    const [searchDirty, setSearchDirty] = useState(true)
    const isCSV = useRef<boolean>(false)
    const keydownRef = useRef<((e: KeyboardEvent) => void) | null>(null)
    const lastPathRef = useRef<string>('')
    const lastExtRef = useRef<string>('')
    const spreadsheetRef = useRef<any>(null)
    const workbookRef = useRef<any>(null)
    const matchesRef = useRef<SearchMatch[]>([])
    const searchInputRef = useRef<HTMLInputElement>(null)
    const bufferRef = useRef<ArrayBuffer | null>(null);
    const loadedSheetsRef = useRef<Set<string>>(new Set());
    const allSheetNamesRef = useRef<string[]>([]);
    const isLazyModeRef = useRef(false);
    const currentEncodingRef = useRef<string>('utf-8');

    const doSearch = useCallback((text: string) => {
        setSearchDirty(false);
        const ss = spreadsheetRef.current;
        if (!ss || !text) {
            matchesRef.current = [];
            setMatchCount(0);
            setCurrentMatch(0);
            return;
        }

        const matches: SearchMatch[] = [];
        const lowerText = text.toLowerCase();

        ss.datas.forEach((data: any, sheetIndex: number) => {
            const rows = data.rows;
            for (const riStr of Object.keys(rows._)) {
                const ri = parseInt(riStr);
                const row = rows._[ri];
                if (!row || !row.cells) continue;
                for (const ciStr of Object.keys(row.cells)) {
                    const ci = parseInt(ciStr);
                    const cell = row.cells[ci];
                    if (cell && cell.text != null) {
                        const cellText = String(cell.text).toLowerCase();
                        if (cellText.includes(lowerText)) {
                            matches.push({ ri, ci, sheetIndex });
                        }
                    }
                }
            }
        });

        matchesRef.current = matches;
        setMatchCount(matches.length);
        if (matches.length > 0) {
            setCurrentMatch(1);
            navigateToMatch(0);
        } else {
            setCurrentMatch(0);
        }
    }, []);

    const navigateToMatch = useCallback((index: number) => {
        const ss = spreadsheetRef.current;
        const matches = matchesRef.current;
        if (!ss || !matches.length || index < 0 || index >= matches.length) return;

        const match = matches[index];
        const data = ss.datas[match.sheetIndex];

        // Switch sheet if needed
        if (ss.bottombar && ss.data !== data) {
            ss.sheet.resetData(data);
        }

        // Scroll to make the cell visible
        const { rows, cols, freeze } = data;
        const [fri, fci] = freeze;
        const targetTop = rows.sumHeight(fri, match.ri);
        const targetLeft = cols.sumWidth(fci, match.ci);

        data.scroll.ri = match.ri > fri ? match.ri - 1 : 0;
        data.scroll.ci = match.ci > fci ? match.ci - 1 : 0;
        data.scroll.y = targetTop > 0 ? targetTop - rows.getHeight(match.ri) : 0;
        data.scroll.x = targetLeft > 0 ? targetLeft - cols.getWidth(match.ci) : 0;

        // Select the cell
        data.selector.setIndexes(match.ri, match.ci);
        data.selector.range = data.calSelectedRangeByStart(match.ri, match.ci);

        // Re-render
        ss.sheet.table.render();
        ss.sheet.selector.set(match.ri, match.ci);
    }, []);

    const nextMatch = useCallback(() => {
        if (searchDirty || matchesRef.current.length === 0) {
            doSearch(searchText);
            return;
        }
        const next = currentMatch >= matchesRef.current.length ? 1 : currentMatch + 1;
        setCurrentMatch(next);
        navigateToMatch(next - 1);
    }, [currentMatch, navigateToMatch, searchDirty, searchText, doSearch]);

    const prevMatch = useCallback(() => {
        if (searchDirty || matchesRef.current.length === 0) {
            doSearch(searchText);
            return;
        }
        const prev = currentMatch <= 1 ? matchesRef.current.length : currentMatch - 1;
        setCurrentMatch(prev);
        navigateToMatch(prev - 1);
    }, [currentMatch, navigateToMatch, searchDirty, searchText, doSearch]);

    const closeSearch = useCallback(() => {
        setSearchVisible(false);
        setSearchText('');
        matchesRef.current = [];
        setMatchCount(0);
        setCurrentMatch(0);
        setSearchDirty(true);
    }, []);

    useEffect(() => {
        const container = document.getElementById('container');

        const renderExcel = (path: string, ext: string, encoding: string = 'utf-8', isEncodingExplicit: boolean = false) => {
            const startTime = Date.now();
            console.log('Loading Excel file...');
            setLoading(true);
            fetch(path).then(response => response.arrayBuffer()).then(res => {
                let effectiveEncoding = encoding;
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
                let sheetData;

                if (shouldLazyLoad(res.byteLength)) {
                    isLazyModeRef.current = true;
                    bufferRef.current = res;
                    const sheetNames = loadSheetNames(res, ext);
                    allSheetNamesRef.current = sheetNames;
                    excelData = loadSingleSheet(res, ext, sheetNames[0], effectiveEncoding);
                    loadedSheetsRef.current = new Set([sheetNames[0]]);

                    // Create placeholder sheets for unloaded tabs
                    sheetData = sheetNames.map((name, idx) => {
                        if (idx === 0) return excelData.sheets[0];
                        return { name, rows: { 0: { cells: { 0: { text: 'Loading...' } } } } };
                    });
                } else {
                    isLazyModeRef.current = false;
                    bufferRef.current = null;
                    excelData = loadSheets(res, ext, effectiveEncoding);
                    sheetData = excelData.sheets;
                    allSheetNamesRef.current = excelData.sheets.map(s => s.name);
                    loadedSheetsRef.current = new Set(allSheetNamesRef.current);
                }

                const { maxLength, maxCols } = excelData;
                workbookRef.current = excelData.workbook || null;

                container.innerHTML = '';
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
                        if (sheetName && !loadedSheetsRef.current.has(sheetName) && bufferRef.current) {
                            const singleData = loadSingleSheet(
                                bufferRef.current,
                                lastExtRef.current,
                                sheetName,
                                currentEncodingRef.current
                            );
                            loadedSheetsRef.current.add(sheetName);
                            const loadedSheet = singleData.sheets[0];
                            if (loadedSheet && spreadSheet.datas[index]) {
                                const d = spreadSheet.datas[index];
                                // Update the DataProxy's row data
                                if (d.rows && d.rows.setData) {
                                    d.rows.setData(loadedSheet.rows);
                                } else {
                                    d.rows._ = loadedSheet.rows;
                                }
                                // Update column widths
                                if (loadedSheet.cols) {
                                    for (const ci of Object.keys(loadedSheet.cols)) {
                                        if (d.cols && d.cols.setWidth) {
                                            d.cols.setWidth(Number(ci), loadedSheet.cols[ci].width);
                                        }
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
                        export_xlsx(spreadSheet, ext, workbookRef.current, currentEncodingRef.current);
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
                setLoading(false)
            });
        };

        handler.on("open", ({ path, ext, encoding, isExplicit }) => {
            lastPathRef.current = path;
            lastExtRef.current = ext;
            renderExcel(path, ext, encoding, isExplicit === true);
        }).on("changeEncoding", (encoding: string) => {
            if (lastPathRef.current) {
                renderExcel(lastPathRef.current, lastExtRef.current, encoding, true);
            }
        }).on("saveDone", () => {
            message.success({
                duration: 1,
                content: 'Save done',
            });
            // Clear dirty marker
            if (document.title.endsWith(DIRTY_MARKER)) {
                document.title = document.title.slice(0, -DIRTY_MARKER.length);
            }
        }).emit("init")

        return () => {
            if (keydownRef.current) {
                window.removeEventListener('keydown', keydownRef.current);
            }
        }
    }, [])

    return (
        <div className='excel-viewer'>
            <Spin spinning={loading} fullscreen={true}>
            </Spin>
            {searchVisible && (
                <div className="excel-search-bar"
                    onClick={(e) => e.stopPropagation()}
                    onMouseDown={(e) => e.stopPropagation()}
                    onKeyDown={(e) => e.stopPropagation()}
                    onKeyUp={(e) => e.stopPropagation()}
                    onKeyPress={(e) => e.stopPropagation()}
                >
                    <input
                        ref={searchInputRef}
                        type="text"
                        placeholder="Search..."
                        value={searchText}
                        autoFocus
                        onFocus={() => {
                            const ss = spreadsheetRef.current;
                            if (ss?.sheet) ss.sheet.focusing = false;
                        }}
                        onChange={(e) => {
                            setSearchText(e.target.value);
                            setSearchDirty(true);
                        }}
                        onKeyDown={(e) => {
                            e.stopPropagation();
                            if (e.key === 'Enter') {
                                if (e.shiftKey) prevMatch();
                                else nextMatch();
                            }
                            if (e.key === 'Escape') closeSearch();
                        }}
                    />
                    <span className="excel-search-count">
                        {matchCount > 0 ? `${currentMatch}/${matchCount}` : (searchDirty ? '' : 'No results')}
                    </span>
                    <button onClick={() => doSearch(searchText)} title="Search (Enter)">&#x1F50D;</button>
                    <button onClick={prevMatch} title="Previous (Shift+Enter)">&#x25B2;</button>
                    <button onClick={nextMatch} title="Next (Enter)">&#x25BC;</button>
                    <button onClick={closeSearch} title="Close (Esc)">&#x2715;</button>
                </div>
            )}
            <div id='container'></div>
            {
                isCSV.current ? <VSCodeLogo /> : null
            }
        </div>
    )
}
