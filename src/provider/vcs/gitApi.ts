import * as vscode from 'vscode';
import * as path from 'path';
import type { GitExtension, API, Repository } from '../../types/vscode-git';

export class GitApiHelper {
    private static _instance: GitApiHelper | undefined;
    private api: API | undefined;
    private initialized = false;
    private initPromise: Promise<API | undefined> | undefined;

    static get instance(): GitApiHelper {
        if (!this._instance) this._instance = new GitApiHelper();
        return this._instance;
    }

    async ensureInit(): Promise<API | undefined> {
        if (this.initialized) return this.api;
        if (this.initPromise) return this.initPromise;
        this.initPromise = (async () => {
            const ext = vscode.extensions.getExtension<GitExtension>('vscode.git');
            if (!ext) { this.initialized = true; return undefined; }
            if (!ext.isActive) {
                try { await ext.activate(); } catch { /* ignore */ }
            }
            try {
                this.api = ext.exports.getAPI(1);
            } catch { /* version mismatch */ }
            this.initialized = true;
            return this.api;
        })();
        return this.initPromise;
    }

    repoFor(uri: vscode.Uri): Repository | undefined {
        if (!this.api) return undefined;
        const target = uri.fsPath;
        let best: Repository | undefined;
        let bestLen = -1;
        const sep = path.sep;
        for (const r of this.api.repositories) {
            const root = r.rootUri.fsPath;
            if (target === root || target.startsWith(root + sep)) {
                if (root.length > bestLen) { best = r; bestLen = root.length; }
            }
        }
        return best;
    }

    repos(): Repository[] {
        return this.api?.repositories ?? [];
    }

    onDidOpenRepository(cb: (r: Repository) => void): vscode.Disposable {
        if (!this.api) return new vscode.Disposable(() => { });
        return this.api.onDidOpenRepository(cb);
    }

    onDidCloseRepository(cb: (r: Repository) => void): vscode.Disposable {
        if (!this.api) return new vscode.Disposable(() => { });
        return this.api.onDidCloseRepository(cb);
    }
}
