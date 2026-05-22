import { execFile } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';
import { GitApiHelper } from './gitApi';
import { LogEntry, BlameEntry } from './types';

const MAX_BUFFER = 50 * 1024 * 1024;

function execBuf(cmd: string, args: string[], cwd: string): Promise<Buffer> {
    return new Promise((resolve, reject) => {
        execFile(cmd, args, { cwd, encoding: 'buffer' as any, maxBuffer: MAX_BUFFER },
            (err, stdout) => err ? reject(err) : resolve(stdout as Buffer));
    });
}

function execStr(cmd: string, args: string[], cwd: string): Promise<string> {
    return new Promise((resolve, reject) => {
        execFile(cmd, args, { cwd, encoding: 'utf8', maxBuffer: MAX_BUFFER },
            (err, stdout) => err ? reject(err) : resolve(stdout as string));
    });
}

export class CliFallback {

    static async gitInfo(uri: vscode.Uri): Promise<{ root: string; relPath: string } | null> {
        const api = await GitApiHelper.instance.ensureInit();
        const repo = api ? GitApiHelper.instance.repoFor(uri) : undefined;
        if (repo) {
            const root = repo.rootUri.fsPath;
            return { root, relPath: path.relative(root, uri.fsPath).replace(/\\/g, '/') };
        }
        try {
            const cwd = path.dirname(uri.fsPath);
            const root = (await execStr('git', ['rev-parse', '--show-toplevel'], cwd)).trim();
            return { root, relPath: path.relative(root, uri.fsPath).replace(/\\/g, '/') };
        } catch {
            return null;
        }
    }

    static async gitShow(uri: vscode.Uri, ref: string): Promise<Buffer> {
        const info = await this.gitInfo(uri);
        if (!info) throw new Error('Not a git repository');
        const spec = ref === ':' ? `:${info.relPath}` : `${ref}:${info.relPath}`;
        return execBuf('git', ['show', spec], info.root);
    }

    static async gitLog(uri: vscode.Uri, limit: number): Promise<LogEntry[]> {
        const info = await this.gitInfo(uri);
        if (!info) throw new Error('Not a git repository');
        const out = await execStr('git',
            ['log', '--follow', `-${limit}`, '--format=%H%x1f%s%x1f%ar', '--', info.relPath],
            info.root);
        return out.trim().split('\n').filter(Boolean).map(line => {
            const [hash, message, date] = line.split('\x1f');
            return { hash, message: message || '', date: date || '' };
        });
    }

    static async gitStashList(root: string): Promise<{ message: string }[]> {
        try {
            const out = await execStr('git', ['stash', 'list', '--format=%gs'], root);
            return out.trim().split('\n').filter(Boolean).map(message => ({ message }));
        } catch {
            return [];
        }
    }

    static async gitCommitInfo(uri: vscode.Uri, hash: string): Promise<BlameEntry | null> {
        const info = await this.gitInfo(uri);
        if (!info) return null;
        try {
            const out = await execStr('git',
                ['show', '-s', `--format=%H%x1f%an%x1f%ae%x1f%aI%x1f%s`, hash],
                info.root);
            const [h, author, email, date, message] = out.trim().split('\x1f');
            return { hash: h, author, email, date, message };
        } catch {
            return null;
        }
    }

    static async svnInfo(uri: vscode.Uri): Promise<{ revision: string } | null> {
        try {
            const out = await execStr('svn', ['info', '--show-item', 'revision', uri.fsPath], path.dirname(uri.fsPath));
            return { revision: out.trim() };
        } catch {
            return null;
        }
    }

    static async svnCat(uri: vscode.Uri, revision: string): Promise<Buffer> {
        return execBuf('svn', ['cat', '-r', revision, uri.fsPath], path.dirname(uri.fsPath));
    }

    static async svnLog(uri: vscode.Uri, limit: number): Promise<LogEntry[]> {
        const out = await execStr('svn', ['log', '-l', String(limit), '--xml', uri.fsPath], path.dirname(uri.fsPath));
        const entries: LogEntry[] = [];
        const blocks = out.match(/<logentry[^>]*>[\s\S]*?<\/logentry>/g) || [];
        for (const b of blocks) {
            const rev = b.match(/revision="(\d+)"/)?.[1] || '';
            const msg = (b.match(/<msg>([\s\S]*?)<\/msg>/)?.[1] || '').trim();
            const date = (b.match(/<date>([\s\S]*?)<\/date>/)?.[1] || '').substring(0, 10);
            entries.push({ hash: rev, message: msg, date });
        }
        return entries;
    }
}
