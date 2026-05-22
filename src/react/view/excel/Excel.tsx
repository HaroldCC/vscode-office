import { message, Spin } from "antd";
import { useEffect, useRef, useState } from "react";
import { handler } from "../../util/vscode.ts";
import VSCodeLogo from "../vscode.tsx";
import './Excel.less';
import { loadSheets } from "./excel_reader.ts";
import { export_xlsx } from "./excel_writer.ts";
import Spreadsheet from './x-spreadsheet/index';

export default function Excel() {
    const [loading, setLoading] = useState(true)
    const isCSV = useRef<boolean>(false)
    const keydownRef = useRef<((e: KeyboardEvent) => void) | null>(null)
    const lastPathRef = useRef<string>('')
    const lastExtRef = useRef<string>('')
    useEffect(() => {
        const container = document.getElementById('container');

        const renderExcel = (path: string, ext: string, encoding: string = 'utf-8') => {
            const startTime = Date.now();
            console.log('Loading Excel file...');
            setLoading(true);
            fetch(path).then(response => response.arrayBuffer()).then(res => {
                const { sheets, maxLength, maxCols } = loadSheets(res, ext, encoding);
                isCSV.current = ext?.match(/csv/i) !== null;
                container.innerHTML = ''
                const spreadSheet = new Spreadsheet(container, {
                    showToolbar: false,
                    row: {
                        len: maxLength + 50,
                        height: 30,
                    },
                    col: {
                        len: maxCols,
                    },
                    view: {
                        height: () => window.innerHeight - 2,
                    }
                });
                if (keydownRef.current) {
                    window.removeEventListener('keydown', keydownRef.current);
                }
                const onKeydown = (e: KeyboardEvent) => {
                    if ((e.ctrlKey || e.metaKey) && e.code == "KeyS") {
                        export_xlsx(spreadSheet, ext);
                    }
                };
                keydownRef.current = onKeydown;
                window.addEventListener('keydown', onKeydown);
                setLoading(false)
                spreadSheet.loadData(sheets);
                const endTime = Date.now();
                console.log(`Excel file loaded successfully. Time elapsed: ${endTime - startTime}ms`);
            }).catch(error => {
                console.error(`Failed to load Excel file: ${error.message}`);
                setLoading(false)
            });
        };

        handler.on("open", ({ path, ext, encoding }) => {
            lastPathRef.current = path;
            lastExtRef.current = ext;
            renderExcel(path, ext, encoding);
        }).on("changeEncoding", (encoding: string) => {
            if (lastPathRef.current) {
                renderExcel(lastPathRef.current, lastExtRef.current, encoding);
            }
        }).on("saveDone", () => {
            message.success({
                duration: 1,
                content: 'Save done',
            })
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
            <div id='container'></div>
            {
                isCSV.current ? <VSCodeLogo /> : null
            }
        </div>
    )
}