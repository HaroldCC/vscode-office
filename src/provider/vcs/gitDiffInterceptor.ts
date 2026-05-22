import * as vscode from 'vscode';
import { ExcelDiffProvider } from '../excelDiffProvider';

const OFFICE_RE = /\.(xlsx|xlsm|xls|csv|ods)$/i;

/**
 * 拦截 vscode 原生 git 视图打开 xlsx/csv 时的 textual diff:
 * 当活动编辑器变成 `git:` 或 diff editor 且涉及 office 文件,
 * 关掉该 tab 并以本插件的 diff 面板替代。
 *
 * 取代原本的独立 SCM 面板。
 */
export class GitDiffInterceptor {

    private disposables: vscode.Disposable[] = [];
    private handling = new Set<string>();

    constructor(private excelDiff: ExcelDiffProvider) { }

    activate(ctx: vscode.ExtensionContext): void {
        this.disposables.push(
            vscode.window.onDidChangeActiveTextEditor(ed => this.onActive(ed)),
            vscode.window.tabGroups.onDidChangeTabs(e => {
                for (const tab of e.opened) this.maybeInterceptTab(tab);
            }),
        );
        // 首次激活时扫描已有 tabs
        for (const group of vscode.window.tabGroups.all) {
            for (const tab of group.tabs) this.maybeInterceptTab(tab);
        }
        ctx.subscriptions.push({ dispose: () => this.disposables.forEach(d => d.dispose()) });
    }

    private onActive(ed?: vscode.TextEditor): void {
        if (!ed) return;
        const uri = ed.document.uri;
        // 原生 git diff 在 git: scheme 打开右侧
        if (uri.scheme !== 'git' && uri.scheme !== 'gitfs') return;
        if (!OFFICE_RE.test(uri.path)) return;
        const realPath = this.gitUriToFsPath(uri);
        if (!realPath) return;
        if (this.handling.has(realPath)) return;
        this.handling.add(realPath);
        setTimeout(() => this.handling.delete(realPath), 1500);

        // 关掉这个 textual diff tab(异步,避免触发 onDidChange 死循环)
        vscode.commands.executeCommand('workbench.action.closeActiveEditor').then(() => {
            this.excelDiff.diffWithVCS(vscode.Uri.file(realPath));
        });
    }

    private maybeInterceptTab(tab: vscode.Tab): void {
        const input = tab.input as any;
        if (!input) return;
        // TabInputTextDiff: { original: Uri, modified: Uri }
        if (input.original && input.modified) {
            const modUri: vscode.Uri = input.modified;
            const origUri: vscode.Uri = input.original;
            const target = OFFICE_RE.test(modUri.path) ? modUri : OFFICE_RE.test(origUri.path) ? origUri : undefined;
            if (!target) return;
            // 找出真实磁盘文件
            const fsPath = this.gitUriToFsPath(target) ||
                (target.scheme === 'file' ? target.fsPath : undefined);
            if (!fsPath) return;
            if (this.handling.has(fsPath)) return;
            this.handling.add(fsPath);
            setTimeout(() => this.handling.delete(fsPath), 1500);

            vscode.window.tabGroups.close(tab).then(() => {
                this.excelDiff.diffWithVCS(vscode.Uri.file(fsPath));
            });
        }
    }

    /** git: URI 的 query 中包含 {"path":...,"ref":...} */
    private gitUriToFsPath(uri: vscode.Uri): string | undefined {
        if (uri.scheme === 'file') return uri.fsPath;
        try {
            const q = JSON.parse(uri.query);
            if (q?.path) return q.path;
        } catch { /* ignore */ }
        // fallback:把 git: scheme 的 path 当作 fsPath
        if (uri.path) return uri.path;
        return undefined;
    }
}
