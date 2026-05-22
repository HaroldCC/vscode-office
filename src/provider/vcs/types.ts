import * as vscode from 'vscode';

export type Ref =
    | { kind: 'working' }
    | { kind: 'staged' }
    | { kind: 'head' }
    | { kind: 'commit'; hash: string; label?: string; message?: string; date?: string }
    | { kind: 'stash'; index: number; message?: string }
    | { kind: 'file'; uri: vscode.Uri; label?: string }
    | { kind: 'svn-working' }
    | { kind: 'svn-revision'; revision: string; message?: string; date?: string };

export interface BlameEntry {
    hash: string;
    author: string;
    email: string;
    date: string;
    message: string;
}

export interface RepoInfo {
    kind: 'git' | 'svn' | null;
    root: string;
    relPath: string;
}

export interface LogEntry {
    hash: string;
    message: string;
    date: string;
}

export function refLabel(ref: Ref): string {
    switch (ref.kind) {
        case 'working': return 'Working';
        case 'staged': return 'Staged';
        case 'head': return 'HEAD';
        case 'commit': return ref.label || ref.hash.substring(0, 7);
        case 'stash': return `stash@{${ref.index}}`;
        case 'file': return ref.label || ref.uri.path.split('/').pop() || 'file';
        case 'svn-working': return 'Working';
        case 'svn-revision': return `r${ref.revision}`;
    }
}

export function refsEqual(a: Ref, b: Ref): boolean {
    if (a.kind !== b.kind) return false;
    switch (a.kind) {
        case 'commit': return a.hash === (b as any).hash;
        case 'stash': return a.index === (b as any).index;
        case 'file': return a.uri.toString() === (b as any).uri.toString();
        case 'svn-revision': return a.revision === (b as any).revision;
        default: return true;
    }
}
