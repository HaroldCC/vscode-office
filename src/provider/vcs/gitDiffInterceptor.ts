import * as vscode from 'vscode';
import { ExcelDiffProvider } from '../excelDiffProvider';

const OFFICE_RE = /\.(xlsx|xlsm|xls|csv|ods)$/i;

/**
 * 拦截 vscode 原生 git 视图打开 xlsx/csv 时的 textual diff:
 * 监听 tab 打开事件,识别 TabInputTextDiff 涉及 office 文件 → 关掉并
 * 替换为本插件的 diff 面板。
 */
export class GitDiffInterceptor {

    private disposables: vscode.Disposable[] = [];
    private intercepted = new Set<string>();
    private active = false;

    constructor(private excelDiff: ExcelDiffProvider) { }

    activate(ctx: vscode.ExtensionContext): void {
        // 启动几秒后才激活,避免误关用户在启动恢复阶段的 tab
        setTimeout(() => { this.active = true; }, 800);
        this.disposables.push(
            vscode.window.tabGroups.onDidChangeTabs(e => {
                for (const tab of e.opened) this.maybeInterceptTab(tab);
                for (const tab of e.changed) this.maybeInterceptTab(tab);
            }),
        );
        ctx.subscriptions.push({ dispose: () => this.disposables.forEach(d => d.dispose()) });
    }

    private maybeInterceptTab(tab: vscode.Tab): void {
        if (!this.active) return;
        const input = tab.input;
        if (!input || typeof input !== 'object') return;

        // TabInputTextDiff: 有 original 与 modified 两个 Uri
        const original = (input as any).original as vscode.Uri | undefined;
        const modified = (input as any).modified as vscode.Uri | undefined;
        if (!original || !modified || !(modified instanceof vscode.Uri)) return;

        const candidatePath = OFFICE_RE.test(modified.path) ? modified
            : OFFICE_RE.test(original.path) ? original
            : undefined;
        if (!candidatePath) return;

        const fsPath = this.resolveFsPath(candidatePath, original, modified);
        if (!fsPath) return;
        if (this.intercepted.has(fsPath)) return;
        this.intercepted.add(fsPath);
        // 同一文件 3 秒内不重复处理(防止 tabs.onDidChange 重复事件)
        setTimeout(() => this.intercepted.delete(fsPath), 3000);

        // 关闭原生 textual diff tab 再开本插件 diff
        vscode.window.tabGroups.close(tab, true).then(() => {
            this.excelDiff.diffWithVCS(vscode.Uri.file(fsPath));
        }, () => {
            // 关闭失败也尝试打开
            this.excelDiff.diffWithVCS(vscode.Uri.file(fsPath));
        });
    }

    /**
     * 从 git: scheme 的 Uri 解析真实磁盘路径。
     * vscode.git 的 query 是 JSON: { path: "/abs/...", ref: "HEAD" }
     */
    private resolveFsPath(uri: vscode.Uri, original: vscode.Uri, modified: vscode.Uri): string | undefined {
        // 优先使用 file scheme 那端
        if (modified.scheme === 'file') return modified.fsPath;
        if (original.scheme === 'file') return original.fsPath;
        // 否则解析 git: query
        for (const u of [uri, modified, original]) {
            if (u.scheme === 'git' || u.scheme === 'gitfs') {
                try {
                    const q = JSON.parse(u.query);
                    if (q?.path) return q.path;
                } catch { /* ignore */ }
            }
        }
        // fallback: 在已知 workspace 中寻找匹配 basename 的文件
        return undefined;
    }
}
