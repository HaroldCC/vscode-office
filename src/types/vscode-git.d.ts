import { Uri, Event, Disposable, SourceControlInputBox } from 'vscode';

export interface GitExtension {
    readonly enabled: boolean;
    readonly onDidChangeEnablement: Event<boolean>;
    getAPI(version: 1): API;
}

export interface API {
    readonly state: 'uninitialized' | 'initialized';
    readonly onDidChangeState: Event<'uninitialized' | 'initialized'>;
    readonly repositories: Repository[];
    readonly onDidOpenRepository: Event<Repository>;
    readonly onDidCloseRepository: Event<Repository>;
}

export const enum Status {
    INDEX_MODIFIED = 0,
    INDEX_ADDED = 1,
    INDEX_DELETED = 2,
    INDEX_RENAMED = 3,
    INDEX_COPIED = 4,
    MODIFIED = 5,
    DELETED = 6,
    UNTRACKED = 7,
    IGNORED = 8,
    INTENT_TO_ADD = 9,
    INTENT_TO_RENAME = 10,
    TYPE_CHANGED = 11,
    ADDED_BY_US = 12,
    ADDED_BY_THEM = 13,
    DELETED_BY_US = 14,
    DELETED_BY_THEM = 15,
    BOTH_ADDED = 16,
    BOTH_DELETED = 17,
    BOTH_MODIFIED = 18
}

export interface Change {
    readonly uri: Uri;
    readonly originalUri: Uri;
    readonly renameUri: Uri | undefined;
    readonly status: Status;
}

export interface RepositoryState {
    readonly HEAD: Branch | undefined;
    readonly workingTreeChanges: Change[];
    readonly indexChanges: Change[];
    readonly mergeChanges: Change[];
    readonly onDidChange: Event<void>;
}

export interface Branch {
    readonly name?: string;
    readonly commit?: string;
    readonly type?: number;
}

export interface Repository {
    readonly rootUri: Uri;
    readonly inputBox: SourceControlInputBox;
    readonly state: RepositoryState;
    readonly ui: { readonly selected: boolean };
    show(path: string, ref?: string): Promise<string>;
    log(options?: { readonly path?: string; readonly maxEntries?: number }): Promise<Commit[]>;
    getCommit(ref: string): Promise<Commit>;
}

export interface Commit {
    readonly hash: string;
    readonly message: string;
    readonly parents: string[];
    readonly authorDate?: Date;
    readonly authorName?: string;
    readonly authorEmail?: string;
    readonly commitDate?: Date;
}
