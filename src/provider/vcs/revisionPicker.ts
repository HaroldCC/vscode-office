import * as vscode from 'vscode';
import { VcsResolver } from './vcsResolver';
import { Ref, refLabel } from './types';

interface RefItem extends vscode.QuickPickItem {
    ref?: Ref;
    isBrowse?: boolean;
}

export async function showRevisionPicker(uri: vscode.Uri, opts?: { excludeKind?: Ref['kind'] }): Promise<Ref | undefined> {
    const refs = await VcsResolver.listAvailableRefs(uri);
    const items: RefItem[] = [];

    const wt = refs.filter(r => r.kind === 'working' || r.kind === 'staged' || r.kind === 'svn-working');
    if (wt.length) {
        items.push({ label: 'Working Tree', kind: vscode.QuickPickItemKind.Separator });
        for (const r of wt) {
            if (opts?.excludeKind === r.kind) continue;
            items.push({
                label: refLabel(r),
                description: r.kind === 'staged' ? 'git index' : 'on disk',
                ref: r,
            });
        }
    }

    const commits = refs.filter(r => r.kind === 'commit' || r.kind === 'svn-revision');
    if (commits.length) {
        items.push({ label: 'History', kind: vscode.QuickPickItemKind.Separator });
        for (const r of commits) {
            const lbl = r.kind === 'commit' ? r.hash.substring(0, 7) : `r${(r as any).revision}`;
            items.push({
                label: lbl,
                description: (r as any).message || '',
                detail: (r as any).date || '',
                ref: r,
            });
        }
    }

    const stashes = refs.filter(r => r.kind === 'stash');
    if (stashes.length) {
        items.push({ label: 'Stash', kind: vscode.QuickPickItemKind.Separator });
        for (const r of stashes) {
            items.push({
                label: `stash@{${(r as any).index}}`,
                description: (r as any).message || '',
                ref: r,
            });
        }
    }

    items.push({ label: 'External', kind: vscode.QuickPickItemKind.Separator });
    items.push({ label: '$(file) Browse file…', isBrowse: true });

    const sel = await vscode.window.showQuickPick(items, {
        placeHolder: 'Select revision to compare',
        matchOnDescription: true,
    });
    if (!sel) return undefined;

    if (sel.isBrowse) {
        const picked = await vscode.window.showOpenDialog({
            canSelectFiles: true,
            canSelectMany: false,
            filters: { 'Excel/CSV': ['xlsx', 'xls', 'xlsm', 'csv', 'ods'] },
        });
        if (!picked || picked.length === 0) return undefined;
        return { kind: 'file', uri: picked[0] };
    }
    return sel.ref;
}
