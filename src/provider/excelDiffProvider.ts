import { execFile } from 'child_process';
import { dirname, relative } from 'path';
import * as vscode from 'vscode';
import { ReactApp } from '../common/reactApp';

export class ExcelDiffProvider {

    private extensionPath: string;

    constructor(private context: vscode.ExtensionContext) {
        this.extensionPath = context.extensionPath;
    }

    /**
     * Compare current file with VCS (git/svn) HEAD version.
     */
    async diffWithVCS(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }

        const filePath = uri.fsPath;

        try {
            const baseBuffer = await this.getVCSVersion(filePath);
            if (!baseBuffer) return;

            const baseLabel = 'HEAD';
            const currentLabel = 'Working';
            this.openDiffPanel(uri, baseBuffer, `${vscode.workspace.asRelativePath(uri)} (${baseLabel} ↔ ${currentLabel})`);
        } catch (err: any) {
            vscode.window.showErrorMessage(`Failed to get VCS version: ${err.message}`);
        }
    }

    /**
     * Compare current file with another user-selected file.
     */
    async diffWithFile(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }

        const ext = uri.fsPath.match(/\.[^.]+$/)?.[0] || '*';
        const filters: { [key: string]: string[] } = {
            'Excel/CSV Files': ['xlsx', 'xls', 'xlsm', 'csv', 'ods'],
        };

        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: true,
            canSelectMany: false,
            filters,
            title: 'Select file to compare with',
        });

        if (!selected || selected.length === 0) return;

        const baseUri = selected[0];
        const baseBuffer = await vscode.workspace.fs.readFile(baseUri);

        const baseName = vscode.workspace.asRelativePath(baseUri);
        const currentName = vscode.workspace.asRelativePath(uri);
        this.openDiffPanel(uri, Buffer.from(baseBuffer), `${currentName} ↔ ${baseName}`, baseName);
    }

    /**
     * Open the diff webview panel.
     */
    private async openDiffPanel(currentUri: vscode.Uri, baseBuffer: Buffer, title: string, baseLabel?: string) {
        const panel = vscode.window.createWebviewPanel(
            'excelDiff',
            title,
            vscode.ViewColumn.Active,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [
                    vscode.Uri.file(this.extensionPath),
                    vscode.Uri.joinPath(currentUri, '..'),
                ],
            }
        );

        const webview = panel.webview;
        await ReactApp.view(webview, { route: 'excel-diff' });

        // Wait a bit for React to mount, then send data
        const ext = currentUri.fsPath.match(/\.[^.]+$/)?.[0] || '.xlsx';
        const currentPath = webview.asWebviewUri(currentUri)
            .with({ query: `nonce=${Date.now()}` }).toString();
        const baseData = baseBuffer.toString('base64');

        // Listen for messages from the webview
        panel.webview.onDidReceiveMessage(async (msg) => {
            if (msg.type === 'init') {
                panel.webview.postMessage({
                    type: 'openDiff',
                    content: {
                        currentPath,
                        baseData,
                        ext,
                        encoding: 'utf-8',
                        baseLabel: baseLabel || 'HEAD',
                        currentLabel: 'Working',
                    }
                });
            } else if (msg.type === 'save') {
                try {
                    const content = msg.content;
                    let data: Uint8Array;
                    if (typeof content === 'string') {
                        data = Buffer.from(content, 'utf-8');
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
            }
        });
    }

    /**
     * Attempt to get the HEAD version of a file from git or svn.
     */
    private async getVCSVersion(filePath: string): Promise<Buffer | null> {
        // Try git first
        try {
            return await this.getGitVersion(filePath);
        } catch {
            // Not a git repo or git not available
        }

        // Try svn
        try {
            return await this.getSvnVersion(filePath);
        } catch {
            // Not an svn repo or svn not available
        }

        vscode.window.showErrorMessage(
            'File is not tracked by Git or SVN. Use "Compare with Another File..." instead.'
        );
        return null;
    }

    private getGitVersion(filePath: string): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            const cwd = dirname(filePath);

            // First get git root
            execFile('git', ['rev-parse', '--show-toplevel'], { cwd, encoding: 'utf8' }, (err, stdout) => {
                if (err) return reject(err);

                const gitRoot = stdout.trim();
                const relPath = relative(gitRoot, filePath).replace(/\\/g, '/');

                // Get HEAD version
                execFile('git', ['show', `HEAD:${relPath}`], { cwd, encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 }, (err, stdout) => {
                    if (err) return reject(err);
                    resolve(stdout as unknown as Buffer);
                });
            });
        });
    }

    private getSvnVersion(filePath: string): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            const cwd = dirname(filePath);
            execFile('svn', ['cat', '-r', 'PREV', filePath], { cwd, encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 }, (err, stdout) => {
                if (err) return reject(err);
                resolve(stdout as unknown as Buffer);
            });
        });
    }

    private getGitLog(filePath: string, limit: number = 20): Promise<{ hash: string; message: string; date: string }[]> {
        return new Promise((resolve, reject) => {
            const cwd = dirname(filePath);
            execFile('git', ['log', '--oneline', '--follow', `-${limit}`, '--format=%H|%s|%ar', '--', filePath],
                { cwd, encoding: 'utf8' }, (err, stdout) => {
                    if (err) return reject(err);
                    const lines = stdout.trim().split('\n').filter(Boolean);
                    const entries = lines.map(line => {
                        const [hash, ...rest] = line.split('|');
                        const message = rest.slice(0, -1).join('|');
                        const date = rest[rest.length - 1] || '';
                        return { hash, message, date };
                    });
                    resolve(entries);
                });
        });
    }

    private getSvnLog(filePath: string, limit: number = 20): Promise<{ hash: string; message: string; date: string }[]> {
        return new Promise((resolve, reject) => {
            const cwd = dirname(filePath);
            execFile('svn', ['log', '-l', String(limit), '--xml', filePath],
                { cwd, encoding: 'utf8' }, (err, stdout) => {
                    if (err) return reject(err);
                    const entries: { hash: string; message: string; date: string }[] = [];
                    const logEntries = stdout.match(/<logentry[^>]*>[\s\S]*?<\/logentry>/g) || [];
                    for (const entry of logEntries) {
                        const rev = entry.match(/revision="(\d+)"/)?.[1] || '';
                        const msg = entry.match(/<msg>([\s\S]*?)<\/msg>/)?.[1]?.trim() || '';
                        const date = entry.match(/<date>([\s\S]*?)<\/date>/)?.[1]?.substring(0, 10) || '';
                        entries.push({ hash: rev, message: msg, date });
                    }
                    resolve(entries);
                });
        });
    }

    private getGitFileAtRevision(filePath: string, hash: string): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            const cwd = dirname(filePath);
            execFile('git', ['rev-parse', '--show-toplevel'], { cwd, encoding: 'utf8' }, (err, stdout) => {
                if (err) return reject(err);
                const gitRoot = stdout.trim();
                const relPath = relative(gitRoot, filePath).replace(/\\/g, '/');
                execFile('git', ['show', `${hash}:${relPath}`], { cwd, encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 }, (err, stdout) => {
                    if (err) return reject(err);
                    resolve(stdout as unknown as Buffer);
                });
            });
        });
    }

    private getSvnFileAtRevision(filePath: string, revision: string): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            const cwd = dirname(filePath);
            execFile('svn', ['cat', '-r', revision, filePath], { cwd, encoding: 'buffer', maxBuffer: 50 * 1024 * 1024 }, (err, stdout) => {
                if (err) return reject(err);
                resolve(stdout as unknown as Buffer);
            });
        });
    }

    async diffWithRevision(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }

        const filePath = uri.fsPath;
        let vcsType: 'git' | 'svn' | null = null;
        let entries: { hash: string; message: string; date: string }[] = [];

        try {
            entries = await this.getGitLog(filePath);
            vcsType = 'git';
        } catch {
            try {
                entries = await this.getSvnLog(filePath);
                vcsType = 'svn';
            } catch {
                vscode.window.showErrorMessage('File is not tracked by Git or SVN.');
                return;
            }
        }

        if (entries.length === 0) {
            vscode.window.showInformationMessage('No revision history found for this file.');
            return;
        }

        const items: (vscode.QuickPickItem & { hash: string })[] = entries.map(e => ({
            label: vcsType === 'git' ? e.hash.substring(0, 7) : `r${e.hash}`,
            description: e.message.length > 60 ? e.message.substring(0, 60) + '...' : e.message,
            detail: e.date,
            hash: e.hash,
        }));

        const selected = await vscode.window.showQuickPick(items, {
            placeHolder: 'Select a revision to compare with',
            matchOnDescription: true,
        });

        if (!selected) return;

        try {
            let baseBuffer: Buffer;
            if (vcsType === 'git') {
                baseBuffer = await this.getGitFileAtRevision(filePath, selected.hash);
            } else {
                baseBuffer = await this.getSvnFileAtRevision(filePath, selected.hash);
            }

            const baseLabel = selected.label;
            const currentLabel = 'Working';
            this.openDiffPanel(uri, baseBuffer, `${vscode.workspace.asRelativePath(uri)} (${baseLabel} ↔ ${currentLabel})`, baseLabel);
        } catch (err: any) {
            vscode.window.showErrorMessage(`Failed to get revision: ${err.message}`);
        }
    }
}
