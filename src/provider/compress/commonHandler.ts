import * as vscode from 'vscode';
import * as iconv from 'iconv-lite';
import { Handler } from "@/common/handler";
import { EncodingStatusBar } from "@/common/encodingStatusBar";
import { Uri, workspace } from 'vscode';
import { extname } from 'path';

const fileSaveTimes: Record<string, number> = {};

export function handleCommonEvent(uri: Uri, handler: Handler, encodingStatusBar?: EncodingStatusBar) {
    const send = () => {
        const now = Date.now();
        const lastSaveTime = fileSaveTimes[uri.toString()];
        if (lastSaveTime && now - lastSaveTime < 100) {
            return;
        }
        const encoding = encodingStatusBar?.getEncoding(uri.toString()) || 'utf-8';
        const isExplicit = encodingStatusBar?.isExplicitlySet(uri.toString()) ?? false;
        handler.emit("open", {
            ext: extname(uri.fsPath),
            path: handler.panel.webview.asWebviewUri(uri).with({ query: `nonce=${now.toString()}` }).toString(),
            encoding,
            isExplicit,
        })
    }
    handler
        .on("editInVSCode", (full: boolean) => {
            const side = full ? vscode.ViewColumn.Active : vscode.ViewColumn.Beside;
            vscode.commands.executeCommand('vscode.openWith', uri, "default", side);
        })
        .on("init", send)
        .on("fileChange", send)
        .on('detectedEncoding', (encoding: string) => {
            if (encodingStatusBar) {
                encodingStatusBar.setDetectedEncoding(uri.toString(), encoding);
            }
        })
        .on("save", async (content) => {
            let res: Uint8Array;
            if (content && typeof content === 'object' && content.text && content.encoding) {
                // CSV with non-UTF-8 encoding: encode text using iconv-lite
                res = iconv.encode(content.text, content.encoding);
            } else if (Array.isArray(content)) {
                res = new Uint8Array(content);
            } else {
                res = new TextEncoder().encode(content);
            }
            await workspace.fs.writeFile(uri, res)
            fileSaveTimes[uri.toString()] = Date.now();
            handler.emit("saveDone")
        })
        .on('developerTool', () => vscode.commands.executeCommand('workbench.action.toggleDevTools'))
        .on('dispose', () => {
            delete fileSaveTimes[uri.toString()];
        })
}