import * as vscode from 'vscode';
import { GitApiHelper } from '../vcs/gitApi';
import type { Repository, Change } from '../../types/vscode-git';

const OFFICE_RE = /\.(xlsx|xlsm|xls|csv|ods)$/i;

export class MergeDetector {
    private disposables: vscode.Disposable[] = [];
    private promptedFiles = new Set<string>();

    async activate(ctx: vscode.ExtensionContext): Promise<void> {
        const api = await GitApiHelper.instance.ensureInit();
        if (!api) return;
        for (const repo of api.repositories) this.watch(repo);
        this.disposables.push(GitApiHelper.instance.onDidOpenRepository(r => this.watch(r)));
        ctx.subscriptions.push({ dispose: () => this.disposables.forEach(d => d.dispose()) });
    }

    private watch(repo: Repository) {
        this.disposables.push(repo.state.onDidChange(() => this.check(repo)));
        this.check(repo);
    }

    private check(repo: Repository) {
        const enabled = vscode.workspace.getConfiguration('vscode-office').get<boolean>('merge.autoPromptOnConflict', true);
        if (!enabled) return;
        const mergeChanges = repo.state.mergeChanges || [];
        const officeConflicts: Change[] = mergeChanges.filter(c => OFFICE_RE.test(c.uri.fsPath));
        if (officeConflicts.length === 0) {
            // 清空 prompted set 当无冲突时,使下次能再次提示
            this.promptedFiles.clear();
            return;
        }
        const fresh = officeConflicts.filter(c => !this.promptedFiles.has(c.uri.fsPath));
        if (fresh.length === 0) return;
        for (const c of fresh) this.promptedFiles.add(c.uri.fsPath);

        vscode.window.showWarningMessage(
            `${officeConflicts.length} office file${officeConflicts.length > 1 ? 's' : ''} have merge conflicts. Open Office Merge Editor?`,
            'Open All',
            'Later',
        ).then(choice => {
            if (choice === 'Open All') {
                for (const c of officeConflicts) {
                    vscode.commands.executeCommand('office.merge.openConflictEditor', c.uri);
                }
            }
        });
    }
}
