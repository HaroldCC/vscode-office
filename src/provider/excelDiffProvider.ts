import * as iconv from 'iconv-lite';
import * as vscode from 'vscode';
import { ReactApp } from '../common/reactApp';
import { VcsResolver } from './vcs/vcsResolver';
import { showRevisionPicker } from './vcs/revisionPicker';
import { BlameProvider } from './vcs/blameProvider';
import { Ref, refLabel } from './vcs/types';

export class ExcelDiffProvider {

    private blameProvider = new BlameProvider();

    constructor(private context: vscode.ExtensionContext) { }

    async diffWithVCS(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }
        const info = await VcsResolver.detect(uri);
        if (info.kind === null) {
            vscode.window.showErrorMessage('File is not tracked by Git or SVN.');
            return;
        }
        const refA: Ref = info.kind === 'git' ? { kind: 'head' } : { kind: 'svn-revision', revision: 'PREV' };
        const refB: Ref = info.kind === 'git' ? { kind: 'working' } : { kind: 'svn-working' };
        await this.openDiff(uri, refA, refB);
    }

    async diffWithFile(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }
        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: true, canSelectMany: false,
            filters: { 'Excel/CSV Files': ['xlsx', 'xls', 'xlsm', 'csv', 'ods'] },
            title: 'Select file to compare with',
        });
        if (!selected || selected.length === 0) return;
        const refA: Ref = { kind: 'file', uri: selected[0] };
        const refB: Ref = { kind: 'working' };
        await this.openDiff(uri, refA, refB);
    }

    /** 被外部 diff 工具(Tortoise)以 URI handler 调起 */
    async diffWithExternalFiles(baseUri: vscode.Uri, mineUri: vscode.Uri, baseLabel: string, mineLabel: string) {
        const refA: Ref = { kind: 'file', uri: baseUri, ...(baseLabel ? { label: baseLabel } as any : {}) };
        const refB: Ref = { kind: 'file', uri: mineUri, ...(mineLabel ? { label: mineLabel } as any : {}) };
        await this.openDiff(mineUri, refA, refB);
    }

    async diffWithRevision(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }
        const ref = await showRevisionPicker(uri);
        if (!ref) return;
        const info = await VcsResolver.detect(uri);
        const refB: Ref = info.kind === 'svn' ? { kind: 'svn-working' } : { kind: 'working' };
        await this.openDiff(uri, ref, refB);
    }

    private async openDiff(currentUri: vscode.Uri, refA: Ref, refB: Ref) {
        // 立即创建 panel,避免点击后"很久才反应"
        const title = `${vscode.workspace.asRelativePath(currentUri)} (${refLabel(refA)} ↔ ${refLabel(refB)})`;
        const panel = vscode.window.createWebviewPanel(
            'excelDiff',
            title,
            vscode.ViewColumn.Active,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [
                    vscode.Uri.file(this.context.extensionPath),
                    vscode.Uri.joinPath(currentUri, '..'),
                ],
            }
        );
        await ReactApp.view(panel.webview, { route: 'excel-diff' });

        // 后台并行加载两端 buffer,加载完成前 webview 自行显示 loading spinner
        let bufA: Buffer | undefined;
        let bufB: Buffer | undefined;
        let loadErr: string | undefined;
        const loadPromise = (async () => {
            try {
                [bufA, bufB] = await Promise.all([
                    VcsResolver.resolveBuffer(currentUri, refA),
                    VcsResolver.resolveBuffer(currentUri, refB),
                ]);
            } catch (err: any) {
                loadErr = err.message || String(err);
            }
        })();

        this.bind(panel, currentUri, refA, refB, async () => {
            await loadPromise;
            if (loadErr) {
                vscode.window.showErrorMessage(`Failed to load diff content: ${loadErr}`);
                return undefined;
            }
            return { bufA: bufA!, bufB: bufB! };
        });
    }

    private bind(
        panel: vscode.WebviewPanel,
        currentUri: vscode.Uri,
        initialRefA: Ref,
        initialRefB: Ref,
        getInitialBufs: () => Promise<{ bufA: Buffer; bufB: Buffer } | undefined>,
    ) {
        let refA = initialRefA;
        let refB = initialRefB;
        let bufA: Buffer | undefined;
        let bufB: Buffer | undefined;
        const ext = currentUri.fsPath.match(/\.[^.]+$/)?.[0] || '.xlsx';

        const sendOpen = (leftData: Buffer, rightData: Buffer) => {
            panel.webview.postMessage({
                type: 'openDiff',
                content: {
                    leftRef: refA,
                    rightRef: refB,
                    leftData: leftData.toString('base64'),
                    rightData: rightData.toString('base64'),
                    ext,
                    encoding: 'utf-8',
                }
            });
        };

        panel.webview.onDidReceiveMessage(async (msg) => {
            if (msg.type === 'init') {
                const initial = await getInitialBufs();
                if (!initial) return;
                bufA = initial.bufA;
                bufB = initial.bufB;
                sendOpen(bufA, bufB);
                return;
            }
            if (msg.type === 'pickRef') {
                const side: 'left' | 'right' = msg.side === 'left' ? 'left' : 'right';
                const newRef = await showRevisionPicker(currentUri);
                if (!newRef) return;
                let newBuf: Buffer;
                try {
                    newBuf = await VcsResolver.resolveBuffer(currentUri, newRef);
                } catch (err: any) {
                    vscode.window.showErrorMessage(`Failed to load: ${err.message}`);
                    return;
                }
                if (side === 'left') { refA = newRef; bufA = newBuf; }
                else { refB = newRef; bufB = newBuf; }
                panel.webview.postMessage({
                    type: 'updateSide',
                    side,
                    ref: newRef,
                    data: newBuf.toString('base64'),
                });
                panel.title = `${vscode.workspace.asRelativePath(currentUri)} (${refLabel(refA)} ↔ ${refLabel(refB)})`;
                return;
            }
            if (msg.type === 'requestBlame') {
                const result = await this.blameProvider.getCellBlame(
                    currentUri, msg.sheet, msg.row, msg.col
                );
                panel.webview.postMessage({
                    type: 'blameResult',
                    sheet: msg.sheet,
                    row: msg.row,
                    col: msg.col,
                    entry: result && 'hash' in result ? result : null,
                    error: result && 'error' in result ? result.error : undefined,
                });
                return;
            }
            if (msg.type === 'save') {
                try {
                    const content = msg.content;
                    let data: Uint8Array;
                    if (content && typeof content === 'object' && content.text && content.encoding) {
                        data = iconv.encode(content.text, content.encoding);
                    } else if (typeof content === 'string') {
                        data = Buffer.from(content, 'utf-8');
                    } else if (Array.isArray(content)) {
                        data = new Uint8Array(content);
                    } else if (content && content.data) {
                        data = new Uint8Array(content.data);
                    } else {
                        data = new Uint8Array(content);
                    }
                    await vscode.workspace.fs.writeFile(currentUri, data);
                    panel.webview.postMessage({ type: 'saveDone' });
                } catch (err: any) {
                    vscode.window.showErrorMessage(`Save failed: ${err.message}`);
                }
                return;
            }
        });
    }
}
