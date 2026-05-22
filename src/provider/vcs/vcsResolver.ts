import * as path from 'path';
import * as vscode from 'vscode';
import { CliFallback } from './cliFallback';
import { Ref, RepoInfo } from './types';

export class VcsResolver {

    static async detect(uri: vscode.Uri): Promise<RepoInfo> {
        const gitInfo = await CliFallback.gitInfo(uri);
        if (gitInfo) return { kind: 'git', root: gitInfo.root, relPath: gitInfo.relPath };
        const svnInfo = await CliFallback.svnInfo(uri);
        if (svnInfo) return { kind: 'svn', root: path.dirname(uri.fsPath), relPath: path.basename(uri.fsPath) };
        return { kind: null, root: path.dirname(uri.fsPath), relPath: path.basename(uri.fsPath) };
    }

    static async resolveBuffer(uri: vscode.Uri, ref: Ref): Promise<Buffer> {
        switch (ref.kind) {
            case 'working': {
                const data = await vscode.workspace.fs.readFile(uri);
                return Buffer.from(data);
            }
            case 'file': {
                const data = await vscode.workspace.fs.readFile(ref.uri);
                return Buffer.from(data);
            }
            case 'staged':
                return CliFallback.gitShow(uri, ':');
            case 'head':
                return CliFallback.gitShow(uri, 'HEAD');
            case 'commit':
                return CliFallback.gitShow(uri, ref.hash);
            case 'stash':
                return CliFallback.gitShow(uri, `stash@{${ref.index}}`);
            case 'svn-working': {
                const data = await vscode.workspace.fs.readFile(uri);
                return Buffer.from(data);
            }
            case 'svn-revision':
                return CliFallback.svnCat(uri, ref.revision);
        }
    }

    static async listAvailableRefs(uri: vscode.Uri): Promise<Ref[]> {
        const info = await this.detect(uri);
        if (info.kind === 'git') {
            const commits = await CliFallback.gitLog(uri, 20).catch(() => []);
            const stashes = await CliFallback.gitStashList(info.root).catch(() => []);
            return [
                { kind: 'working' },
                { kind: 'staged' },
                { kind: 'head' },
                ...commits.map(c => ({ kind: 'commit', hash: c.hash, message: c.message, date: c.date }) as Ref),
                ...stashes.map((s, i) => ({ kind: 'stash', index: i, message: s.message }) as Ref),
            ];
        }
        if (info.kind === 'svn') {
            const revs = await CliFallback.svnLog(uri, 20).catch(() => []);
            return [
                { kind: 'svn-working' },
                ...revs.map(r => ({ kind: 'svn-revision', revision: r.hash, message: r.message, date: r.date }) as Ref),
            ];
        }
        return [{ kind: 'working' }];
    }
}
