import { execFile } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';
import { ReactApp } from '../../common/reactApp';
import { MergeLoader } from './mergeLoader';
import { CliFallback } from '../vcs/cliFallback';

export class MergeEditorProvider {
    constructor(private context: vscode.ExtensionContext) { }

    async openConflictEditor(uri?: vscode.Uri): Promise<void> {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for merge.');
            return;
        }

        let bufs;
        try {
            bufs = await MergeLoader.load(uri);
        } catch (err: any) {
            vscode.window.showErrorMessage(`Failed to load merge sides: ${err.message}`);
            return;
        }

        const ext = uri.fsPath.match(/\.[^.]+$/)?.[0] || '.xlsx';
        const title = `Merge: ${vscode.workspace.asRelativePath(uri)}`;
        const panel = vscode.window.createWebviewPanel(
            'excelMerge',
            title,
            vscode.ViewColumn.Active,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [
                    vscode.Uri.file(this.context.extensionPath),
                    vscode.Uri.joinPath(uri, '..'),
                ],
            }
        );
        await ReactApp.view(panel.webview, { route: 'excel-merge' });

        panel.webview.onDidReceiveMessage(async (msg) => {
            if (msg.type === 'init') {
                panel.webview.postMessage({
                    type: 'openMerge',
                    content: {
                        base: bufs.base.toString('base64'),
                        ours: bufs.ours.toString('base64'),
                        theirs: bufs.theirs.toString('base64'),
                        ext,
                        encoding: 'utf-8',
                        relPath: vscode.workspace.asRelativePath(uri!),
                    }
                });
                return;
            }
            if (msg.type === 'markResolved') {
                try {
                    const data = msg.content;
                    let bytes: Uint8Array;
                    if (Array.isArray(data)) bytes = new Uint8Array(data);
                    else if (data?.data) bytes = new Uint8Array(data.data);
                    else if (typeof data === 'string') bytes = Buffer.from(data, 'utf-8');
                    else bytes = new Uint8Array(data);
                    await vscode.workspace.fs.writeFile(uri!, bytes);
                    await this.gitAdd(uri!);
                    panel.webview.postMessage({ type: 'resolveDone' });
                    vscode.window.showInformationMessage(`Marked ${path.basename(uri!.fsPath)} as resolved.`);
                } catch (err: any) {
                    vscode.window.showErrorMessage(`Mark resolved failed: ${err.message}`);
                }
            }
        });
    }

    private async gitAdd(uri: vscode.Uri): Promise<void> {
        const info = await CliFallback.gitInfo(uri);
        if (!info) return;
        await new Promise<void>((resolve, reject) => {
            execFile('git', ['add', '--', info.relPath], { cwd: info.root, encoding: 'utf8' },
                (err) => err ? reject(err) : resolve());
        });
    }
}
