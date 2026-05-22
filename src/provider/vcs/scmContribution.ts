import * as vscode from 'vscode';
import * as path from 'path';
import { GitApiHelper } from './gitApi';
import type { Repository, Change } from '../../types/vscode-git';

const OFFICE_EXT_DEFAULT = ['xlsx', 'xlsm', 'xls', 'csv', 'ods'];

function statusText(s: number): string {
    const map: Record<number, string> = {
        0: 'Index Modified', 1: 'Index Added', 2: 'Index Deleted',
        5: 'Modified', 6: 'Deleted', 7: 'Untracked',
        16: 'Both Added', 17: 'Both Deleted', 18: 'Both Modified',
    };
    return map[s] ?? 'Changed';
}

export class OfficeScmContribution {
    private scs = new Map<string, vscode.SourceControl>();
    private groups = new Map<string, vscode.SourceControlResourceGroup>();
    private disposables: vscode.Disposable[] = [];

    async activate(ctx: vscode.ExtensionContext): Promise<void> {
        const enabled = vscode.workspace.getConfiguration('vscode-office').get<boolean>('scm.enableOfficePanel', true);
        if (!enabled) return;

        const git = GitApiHelper.instance;
        const api = await git.ensureInit();
        if (!api) return;

        for (const repo of api.repositories) this.attach(repo);
        this.disposables.push(git.onDidOpenRepository(r => this.attach(r)));
        this.disposables.push(git.onDidCloseRepository(r => this.detach(r)));

        ctx.subscriptions.push({
            dispose: () => {
                this.disposables.forEach(d => d.dispose());
                for (const sc of this.scs.values()) sc.dispose();
                this.scs.clear();
                this.groups.clear();
            }
        });
    }

    private get includedExts(): string[] {
        return vscode.workspace.getConfiguration('vscode-office').get<string[]>('scm.includedExtensions', OFFICE_EXT_DEFAULT);
    }

    private isOfficeFile(uri: vscode.Uri): boolean {
        const ext = (uri.fsPath.match(/\.([^.]+)$/)?.[1] || '').toLowerCase();
        return this.includedExts.includes(ext);
    }

    private attach(repo: Repository) {
        const root = repo.rootUri.fsPath;
        if (this.scs.has(root)) return;
        const sc = vscode.scm.createSourceControl('office-excel-diff', `Office Files (${path.basename(root)})`, repo.rootUri);
        const group = sc.createResourceGroup('changes', 'Changes');
        sc.quickDiffProvider = { provideOriginalResource: (uri) => this.provideQuickDiff(uri) };
        this.scs.set(root, sc);
        this.groups.set(root, group);
        this.disposables.push(repo.state.onDidChange(() => this.refresh(repo)));
        this.refresh(repo);
    }

    private detach(repo: Repository) {
        const root = repo.rootUri.fsPath;
        const sc = this.scs.get(root);
        if (sc) sc.dispose();
        this.scs.delete(root);
        this.groups.delete(root);
    }

    private refresh(repo: Repository) {
        const root = repo.rootUri.fsPath;
        const group = this.groups.get(root);
        if (!group) return;
        const changes = [...repo.state.indexChanges, ...repo.state.workingTreeChanges];
        const seen = new Set<string>();
        const officeChanges: Change[] = [];
        for (const c of changes) {
            const key = c.uri.fsPath;
            if (seen.has(key)) continue;
            seen.add(key);
            if (this.isOfficeFile(c.uri)) officeChanges.push(c);
        }
        group.resourceStates = officeChanges.map(c => ({
            resourceUri: c.uri,
            command: { command: 'office.excel.diffWithVCS', title: 'Compare with HEAD', arguments: [c.uri] },
            decorations: {
                strikeThrough: c.status === 2 || c.status === 6,
                tooltip: statusText(c.status),
            }
        }));
    }

    private provideQuickDiff(uri: vscode.Uri): vscode.Uri | undefined {
        if (!/\.csv$/i.test(uri.fsPath)) return undefined;
        return uri.with({ scheme: 'git', path: uri.path, query: JSON.stringify({ path: uri.fsPath, ref: '' }) });
    }
}
