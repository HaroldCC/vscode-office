# VCS Deep Integration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 重构 vscode-office 的 VCS 访问层,接入 vscode.git Extension API、支持任意 Ref 对比、单元格 blame、独立 SCM 面板。

**Architecture:** 新建 `src/provider/vcs/` 模块,以统一 `Ref` 模型贯穿。GitApiHelper 单例负责仓库识别与状态;CliFallback 负责 git/svn 内容读取(`git show`/`svn cat`);VcsResolver 是所有 diff 入口的唯一收束。webview 协议扩展为支持顶部动态切换 ref 与按需 blame。

**Tech Stack:** TypeScript 4.x, vscode API 1.64+, vscode.git Extension API v1, child_process(CLI fallback), XLSX 0.18.5(已有), React 18(webview)

参考 spec: [docs/superpowers/specs/2026-05-22-vcs-deep-integration-design.md](../specs/2026-05-22-vcs-deep-integration-design.md)

**项目无单元测试框架** — 验证以 ad-hoc node 脚本(放 `scratch/` 目录,不提交)+ Extension Development Host(F5)手测为主。每个任务都给出明确的 manual 验证步骤。

---

## File Structure(锁定)

```
src/
├── types/
│   └── vscode-git.d.ts                       [新] vscode.git API 类型定义
├── provider/
│   ├── vcs/
│   │   ├── types.ts                          [新] Ref / BlameEntry / RepoInfo
│   │   ├── gitApi.ts                         [新] GitApiHelper 单例
│   │   ├── cliFallback.ts                    [新] execFile 包装
│   │   ├── vcsResolver.ts                    [新] resolveBuffer / detect / listAvailableRefs
│   │   ├── revisionPicker.ts                 [新] showRevisionPicker QuickPick
│   │   ├── blameProvider.ts                  [新] getCellBlame + LRU
│   │   └── scmContribution.ts                [新] 独立 SourceControl
│   └── excelDiffProvider.ts                  [改] 瘦身,委托新模块
├── extension.ts                              [改] async activate + 初始化新模块
└── react/view/excel/
    ├── ExcelDiff.tsx                         [改] ref switcher + blame UI + 新协议
    └── ExcelDiff.less                        [改] switcher 样式
package.json                                  [改] 3 项新配置
```

---

## Task 1: vscode.git 类型定义文件

**Files:**
- Create: `src/types/vscode-git.d.ts`

- [ ] **Step 1: 创建类型定义文件**

内容来源:vscode 仓库 `extensions/git/src/api/git.d.ts`(MIT)。仅保留本项目用到的成员,避免巨大类型噪音。

```typescript
// src/types/vscode-git.d.ts
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
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无新增 error(此文件目前还没有 import 它的位置)

- [ ] **Step 3: Commit**

```bash
git add src/types/vscode-git.d.ts
git commit -m "feat(vcs): add vscode.git Extension API type definitions"
```

---

## Task 2: vcs/types.ts(核心数据模型)

**Files:**
- Create: `src/provider/vcs/types.ts`

- [ ] **Step 1: 创建类型文件**

```typescript
// src/provider/vcs/types.ts
import * as vscode from 'vscode';

export type Ref =
    | { kind: 'working' }
    | { kind: 'staged' }
    | { kind: 'head' }
    | { kind: 'commit'; hash: string; label?: string; message?: string; date?: string }
    | { kind: 'stash'; index: number; message?: string }
    | { kind: 'file'; uri: vscode.Uri }
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
    root: string;     // 绝对路径
    relPath: string;  // posix path 相对 root
}

export interface LogEntry {
    hash: string;
    message: string;
    date: string;
}

/** 给 UI 显示用的简短 label */
export function refLabel(ref: Ref): string {
    switch (ref.kind) {
        case 'working': return 'Working';
        case 'staged': return 'Staged';
        case 'head': return 'HEAD';
        case 'commit': return ref.label || ref.hash.substring(0, 7);
        case 'stash': return `stash@{${ref.index}}`;
        case 'file': return ref.uri.path.split('/').pop() || 'file';
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
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/vcs/types.ts
git commit -m "feat(vcs): add unified Ref/BlameEntry/RepoInfo types"
```

---

## Task 3: gitApi.ts(GitApiHelper 单例)

**Files:**
- Create: `src/provider/vcs/gitApi.ts`

- [ ] **Step 1: 实现**

```typescript
// src/provider/vcs/gitApi.ts
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
        if (!this.api) return new vscode.Disposable(() => {});
        return this.api.onDidOpenRepository(cb);
    }

    onDidCloseRepository(cb: (r: Repository) => void): vscode.Disposable {
        if (!this.api) return new vscode.Disposable(() => {});
        return this.api.onDidCloseRepository(cb);
    }
}
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/vcs/gitApi.ts
git commit -m "feat(vcs): GitApiHelper singleton wrapping vscode.git API"
```

---

## Task 4: cliFallback.ts(CLI 包装)

**Files:**
- Create: `src/provider/vcs/cliFallback.ts`

- [ ] **Step 1: 实现**

```typescript
// src/provider/vcs/cliFallback.ts
import { execFile, ExecFileOptionsWithBufferEncoding, ExecFileOptionsWithStringEncoding } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';
import { GitApiHelper } from './gitApi';
import { LogEntry, BlameEntry, RepoInfo } from './types';

const MAX_BUFFER = 50 * 1024 * 1024;

function execBuf(cmd: string, args: string[], cwd: string): Promise<Buffer> {
    return new Promise((resolve, reject) => {
        execFile(cmd, args, { cwd, encoding: 'buffer', maxBuffer: MAX_BUFFER } as ExecFileOptionsWithBufferEncoding,
            (err, stdout) => err ? reject(err) : resolve(stdout as Buffer));
    });
}

function execStr(cmd: string, args: string[], cwd: string): Promise<string> {
    return new Promise((resolve, reject) => {
        execFile(cmd, args, { cwd, encoding: 'utf8', maxBuffer: MAX_BUFFER } as ExecFileOptionsWithStringEncoding,
            (err, stdout) => err ? reject(err) : resolve(stdout));
    });
}

export class CliFallback {

    /** 解析 uri 所在 git 仓库根目录 + 相对路径。优先用 vscode.git API,fallback 到 rev-parse */
    static async gitInfo(uri: vscode.Uri): Promise<{ root: string; relPath: string } | null> {
        const api = await GitApiHelper.instance.ensureInit();
        const repo = api ? GitApiHelper.instance.repoFor(uri) : undefined;
        if (repo) {
            const root = repo.rootUri.fsPath;
            return { root, relPath: path.relative(root, uri.fsPath).replace(/\\/g, '/') };
        }
        // fallback CLI(用户安装了 git 但 vscode.git 没识别)
        try {
            const cwd = path.dirname(uri.fsPath);
            const root = (await execStr('git', ['rev-parse', '--show-toplevel'], cwd)).trim();
            return { root, relPath: path.relative(root, uri.fsPath).replace(/\\/g, '/') };
        } catch {
            return null;
        }
    }

    /** ref 可以是 'HEAD' / ':' (staged) / '<hash>' / 'stash@{n}' */
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

    /**
     * 单行 blame。返回最后修改该行的 commit 信息。
     * line 是 1-based。porcelain 输出固定格式,我们只取第一段元数据。
     */
    static async gitBlameLine(uri: vscode.Uri, line: number): Promise<BlameEntry | null> {
        const info = await this.gitInfo(uri);
        if (!info) return null;
        try {
            const out = await execStr('git',
                ['blame', '-L', `${line},${line}`, '--porcelain', '--', info.relPath],
                info.root);
            return this.parseBlamePorcelain(out);
        } catch {
            return null;
        }
    }

    /** 获取某个 commit 的简要信息(给 BlameProvider 用) */
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

    private static parseBlamePorcelain(out: string): BlameEntry | null {
        const lines = out.split(/\r?\n/);
        if (lines.length === 0) return null;
        const headerParts = lines[0].split(' ');
        const hash = headerParts[0];
        if (!hash || hash.length < 7) return null;
        const meta: Record<string, string> = {};
        for (let i = 1; i < lines.length; i++) {
            const l = lines[i];
            if (l.startsWith('\t')) break;
            const sp = l.indexOf(' ');
            if (sp > 0) meta[l.substring(0, sp)] = l.substring(sp + 1);
        }
        return {
            hash,
            author: meta['author'] || '',
            email: (meta['author-mail'] || '').replace(/[<>]/g, ''),
            date: meta['author-time'] ? new Date(parseInt(meta['author-time'], 10) * 1000).toISOString() : '',
            message: meta['summary'] || '',
        };
    }

    // ---- SVN ----

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
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/vcs/cliFallback.ts
git commit -m "feat(vcs): CliFallback for git/svn content reads + log + blame"
```

---

## Task 5: vcsResolver.ts(统一入口)

**Files:**
- Create: `src/provider/vcs/vcsResolver.ts`

- [ ] **Step 1: 实现**

```typescript
// src/provider/vcs/vcsResolver.ts
import * as path from 'path';
import * as vscode from 'vscode';
import { GitApiHelper } from './gitApi';
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
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/vcs/vcsResolver.ts
git commit -m "feat(vcs): VcsResolver unifies ref → buffer resolution"
```

---

## Task 6: revisionPicker.ts(QuickPick UI)

**Files:**
- Create: `src/provider/vcs/revisionPicker.ts`

- [ ] **Step 1: 实现**

```typescript
// src/provider/vcs/revisionPicker.ts
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
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/vcs/revisionPicker.ts
git commit -m "feat(vcs): RevisionPicker QuickPick for unified ref selection"
```

---

## Task 7: blameProvider.ts(单元格 blame)

**Files:**
- Create: `src/provider/vcs/blameProvider.ts`

- [ ] **Step 1: 实现**

```typescript
// src/provider/vcs/blameProvider.ts
import * as vscode from 'vscode';
import * as XLSX from 'xlsx';
import * as iconv from 'iconv-lite';
import { VcsResolver } from './vcsResolver';
import { CliFallback } from './cliFallback';
import { BlameEntry } from './types';

type SheetMap = Map<string /* sheetName */, string[][]>;

class LruCache<K, V> {
    private map = new Map<K, V>();
    constructor(private capacity: number) {}
    get(key: K): V | undefined {
        const v = this.map.get(key);
        if (v !== undefined) { this.map.delete(key); this.map.set(key, v); }
        return v;
    }
    set(key: K, value: V): void {
        if (this.map.has(key)) this.map.delete(key);
        this.map.set(key, value);
        if (this.map.size > this.capacity) {
            const first = this.map.keys().next().value;
            if (first !== undefined) this.map.delete(first);
        }
    }
}

export class BlameProvider {
    private cache = new LruCache<string, SheetMap>(20);

    async getCellBlame(uri: vscode.Uri, sheetName: string, row: number, col: number): Promise<BlameEntry | { error: string } | null> {
        const info = await VcsResolver.detect(uri);
        if (info.kind !== 'git') return { error: 'Blame only supported for git-tracked files' };

        const cfg = vscode.workspace.getConfiguration('vscode-office');
        const depth = Math.max(1, Math.min(100, cfg.get<number>('blame.depth', 20)));

        const workingBuf = await VcsResolver.resolveBuffer(uri, { kind: 'working' });
        const workingValue = this.cellOf(this.parse(workingBuf, uri), sheetName, row, col);

        const commits = await CliFallback.gitLog(uri, depth).catch(() => []);
        if (commits.length === 0) return null;

        let lastSame: { hash: string } | null = null;

        for (const c of commits) {
            const key = `${c.hash}:${info.relPath}`;
            let parsed = this.cache.get(key);
            if (!parsed) {
                try {
                    const buf = await VcsResolver.resolveBuffer(uri, { kind: 'commit', hash: c.hash });
                    parsed = this.parse(buf, uri);
                    this.cache.set(key, parsed);
                } catch {
                    continue;
                }
            }
            const val = this.cellOf(parsed, sheetName, row, col);
            if (val !== workingValue) {
                if (!lastSame) return null;
                return await CliFallback.gitCommitInfo(uri, lastSame.hash);
            }
            lastSame = { hash: c.hash };
        }
        if (lastSame) return await CliFallback.gitCommitInfo(uri, lastSame.hash);
        return { error: `Older than ${depth} commits, blame unavailable` };
    }

    private parse(buf: Buffer, uri: vscode.Uri): SheetMap {
        const ext = (uri.fsPath.match(/\.[^.]+$/)?.[0] || '').toLowerCase();
        const result: SheetMap = new Map();
        if (ext === '.csv') {
            const text = this.decodeCsv(buf);
            const rows = text.split(/\r?\n/).map(line => this.parseCsvLine(line));
            result.set('Sheet1', rows);
            return result;
        }
        try {
            const wb = XLSX.read(buf, { type: 'buffer' });
            for (const sn of wb.SheetNames) {
                const sh = wb.Sheets[sn];
                const rows = XLSX.utils.sheet_to_json<string[]>(sh, { header: 1, raw: false, defval: '' }) as any as string[][];
                result.set(sn, rows);
            }
        } catch { /* malformed history entry */ }
        return result;
    }

    private decodeCsv(buf: Buffer): string {
        // 尝试 utf-8;失败 fallback iconv windows-1252
        try { return buf.toString('utf-8'); }
        catch { return iconv.decode(buf, 'windows-1252'); }
    }

    private parseCsvLine(line: string): string[] {
        // 简化 CSV parser:足够 blame 单元格比较;转义逗号在双引号内
        const out: string[] = [];
        let cur = '';
        let inQ = false;
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (inQ) {
                if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
                else if (ch === '"') { inQ = false; }
                else { cur += ch; }
            } else {
                if (ch === ',') { out.push(cur); cur = ''; }
                else if (ch === '"' && cur === '') { inQ = true; }
                else { cur += ch; }
            }
        }
        out.push(cur);
        return out;
    }

    private cellOf(sheets: SheetMap, sheetName: string, row: number, col: number): string {
        const rows = sheets.get(sheetName) || sheets.values().next().value;
        if (!rows || !rows[row]) return '';
        const v = rows[row][col];
        return v == null ? '' : String(v);
    }
}
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/vcs/blameProvider.ts
git commit -m "feat(vcs): cell-level blame via history replay with LRU cache"
```

---

## Task 8: scmContribution.ts(独立 SCM 面板)

**Files:**
- Create: `src/provider/vcs/scmContribution.ts`

- [ ] **Step 1: 实现**

```typescript
// src/provider/vcs/scmContribution.ts
import * as vscode from 'vscode';
import * as path from 'path';
import { GitApiHelper } from './gitApi';
import type { Repository, Change, Status } from '../../types/vscode-git';

const OFFICE_EXT_DEFAULT = ['xlsx', 'xlsm', 'xls', 'csv', 'ods'];

function statusText(s: Status): string {
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
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/vcs/scmContribution.ts
git commit -m "feat(vcs): independent SCM panel for xlsx/csv changes"
```

---

## Task 9: 重构 excelDiffProvider.ts(瘦身,委托新模块)

**Files:**
- Modify: `src/provider/excelDiffProvider.ts`

- [ ] **Step 1: 完整替换文件内容**

```typescript
// src/provider/excelDiffProvider.ts
import * as iconv from 'iconv-lite';
import * as vscode from 'vscode';
import { ReactApp } from '../common/reactApp';
import { VcsResolver } from './vcs/vcsResolver';
import { showRevisionPicker } from './vcs/revisionPicker';
import { BlameProvider } from './vcs/blameProvider';
import { Ref, refLabel } from './vcs/types';

export class ExcelDiffProvider {

    private blameProvider = new BlameProvider();

    constructor(private context: vscode.ExtensionContext) {}

    /** HEAD ↔ Working(git) / HEAD-equivalent ↔ Working(svn) */
    async diffWithVCS(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }
        const info = await VcsResolver.detect(uri);
        if (info.kind === null) {
            vscode.window.showErrorMessage('File is not tracked by Git or SVN.');
            return;
        }
        const refA: Ref = info.kind === 'git' ? { kind: 'head' } : { kind: 'svn-revision', revision: 'PREV' };
        const refB: Ref = info.kind === 'git' ? { kind: 'working' } : { kind: 'svn-working' };
        await this.openDiff(uri, refA, refB);
    }

    /** Compare with another file selected by user */
    async diffWithFile(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }
        const selected = await vscode.window.showOpenDialog({
            canSelectFiles: true, canSelectMany: false,
            filters: { 'Excel/CSV Files': ['xlsx', 'xls', 'xlsm', 'csv', 'ods'] },
            title: 'Select file to compare with',
        });
        if (!selected || selected.length === 0) return;
        const refA: Ref = { kind: 'file', uri: selected[0] };
        const refB: Ref = { kind: 'working' };
        await this.openDiff(uri, refA, refB);
    }

    /** Compare with arbitrary git/svn revision via QuickPick */
    async diffWithRevision(uri?: vscode.Uri) {
        uri = uri || vscode.window.activeTextEditor?.document.uri;
        if (!uri) {
            vscode.window.showErrorMessage('No file selected for diff.');
            return;
        }
        const ref = await showRevisionPicker(uri);
        if (!ref) return;
        const info = await VcsResolver.detect(uri);
        const refB: Ref = info.kind === 'svn' ? { kind: 'svn-working' } : { kind: 'working' };
        await this.openDiff(uri, ref, refB);
    }

    private async openDiff(currentUri: vscode.Uri, refA: Ref, refB: Ref) {
        let bufA: Buffer, bufB: Buffer;
        try {
            [bufA, bufB] = await Promise.all([
                VcsResolver.resolveBuffer(currentUri, refA),
                VcsResolver.resolveBuffer(currentUri, refB),
            ]);
        } catch (err: any) {
            vscode.window.showErrorMessage(`Failed to load diff content: ${err.message}`);
            return;
        }

        const title = `${vscode.workspace.asRelativePath(currentUri)} (${refLabel(refA)} ↔ ${refLabel(refB)})`;
        const panel = vscode.window.createWebviewPanel(
            'excelDiff',
            title,
            vscode.ViewColumn.Active,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [
                    vscode.Uri.file(this.context.extensionPath),
                    vscode.Uri.joinPath(currentUri, '..'),
                ],
            }
        );
        await ReactApp.view(panel.webview, { route: 'excel-diff' });
        this.bind(panel, currentUri, refA, refB, bufA, bufB);
    }

    private bind(
        panel: vscode.WebviewPanel,
        currentUri: vscode.Uri,
        initialRefA: Ref,
        initialRefB: Ref,
        bufA: Buffer,
        bufB: Buffer,
    ) {
        let refA = initialRefA;
        let refB = initialRefB;
        const ext = currentUri.fsPath.match(/\.[^.]+$/)?.[0] || '.xlsx';

        const sendOpen = (leftData: Buffer, rightData: Buffer) => {
            panel.webview.postMessage({
                type: 'openDiff',
                content: {
                    leftRef: refA,
                    rightRef: refB,
                    leftData: leftData.toString('base64'),
                    rightData: rightData.toString('base64'),
                    ext,
                    encoding: 'utf-8',
                }
            });
        };

        panel.webview.onDidReceiveMessage(async (msg) => {
            if (msg.type === 'init') {
                sendOpen(bufA, bufB);
                return;
            }
            if (msg.type === 'pickRef') {
                const side: 'left' | 'right' = msg.side === 'left' ? 'left' : 'right';
                const newRef = await showRevisionPicker(currentUri);
                if (!newRef) return;
                let newBuf: Buffer;
                try {
                    newBuf = await VcsResolver.resolveBuffer(currentUri, newRef);
                } catch (err: any) {
                    vscode.window.showErrorMessage(`Failed to load: ${err.message}`);
                    return;
                }
                if (side === 'left') { refA = newRef; bufA = newBuf; }
                else { refB = newRef; bufB = newBuf; }
                panel.webview.postMessage({
                    type: 'updateSide',
                    side,
                    ref: newRef,
                    data: newBuf.toString('base64'),
                });
                panel.title = `${vscode.workspace.asRelativePath(currentUri)} (${refLabel(refA)} ↔ ${refLabel(refB)})`;
                return;
            }
            if (msg.type === 'requestBlame') {
                const result = await this.blameProvider.getCellBlame(
                    currentUri, msg.sheet, msg.row, msg.col
                );
                panel.webview.postMessage({
                    type: 'blameResult',
                    sheet: msg.sheet,
                    row: msg.row,
                    col: msg.col,
                    entry: result && 'hash' in result ? result : null,
                    error: result && 'error' in result ? result.error : undefined,
                });
                return;
            }
            if (msg.type === 'save') {
                try {
                    const content = msg.content;
                    let data: Uint8Array;
                    if (content && typeof content === 'object' && content.text && content.encoding) {
                        data = iconv.encode(content.text, content.encoding);
                    } else if (typeof content === 'string') {
                        data = Buffer.from(content, 'utf-8');
                    } else if (Array.isArray(content)) {
                        data = new Uint8Array(content);
                    } else if (content && content.data) {
                        data = new Uint8Array(content.data);
                    } else {
                        data = new Uint8Array(content);
                    }
                    await vscode.workspace.fs.writeFile(currentUri, data);
                    panel.webview.postMessage({ type: 'saveDone' });
                } catch (err: any) {
                    vscode.window.showErrorMessage(`Save failed: ${err.message}`);
                }
                return;
            }
        });
    }
}
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/provider/excelDiffProvider.ts
git commit -m "refactor(diff): excelDiffProvider delegates to vcs/* modules"
```

---

## Task 10: extension.ts 接入新模块

**Files:**
- Modify: `src/extension.ts`

- [ ] **Step 1: 修改 activate 为 async,并初始化 ScmContribution**

替换文件内容(基于现有 [src/extension.ts](../../src/extension.ts) 的精确改造):

```typescript
import * as vscode from 'vscode';
import { EncodingStatusBar } from './common/encodingStatusBar';
import { MarkdownEditorProvider } from './provider/markdownEditorProvider';
import { OfficeViewerProvider } from './provider/officeViewerProvider';
import { ExcelDiffProvider } from './provider/excelDiffProvider';
import { HtmlService } from './service/htmlService';
import { MarkdownService } from './service/markdownService';
import { Output } from './common/Output';
import { FileUtil } from './common/fileUtil';
import { ReactApp } from './common/reactApp';
import { GitApiHelper } from './provider/vcs/gitApi';
import { OfficeScmContribution } from './provider/vcs/scmContribution';
const httpExt = require('./bundle/extension');

export async function activate(context: vscode.ExtensionContext) {
    keepOriginDiff();
    activeHTTP(context);
    const viewOption = { webviewOptions: { retainContextWhenHidden: true, enableFindWidget: false } };
    const viewOptionWithFind = { webviewOptions: { retainContextWhenHidden: true, enableFindWidget: true } };
    FileUtil.init(context);
    ReactApp.init(context);
    const encodingStatusBar = new EncodingStatusBar(context.globalState);
    const markdownService = new MarkdownService(context);
    const viewerInstance = new OfficeViewerProvider(context, encodingStatusBar);
    const markdownEditorProvider = new MarkdownEditorProvider(context, encodingStatusBar);
    const excelDiffProvider = new ExcelDiffProvider(context);

    // VCS subsystem: warm up git API + register Office SCM panel
    GitApiHelper.instance.ensureInit().catch((err) => Output.debug('git api init: ' + err));
    const scmContribution = new OfficeScmContribution();
    scmContribution.activate(context).catch((err) => Output.debug('scm init: ' + err));

    context.subscriptions.push(
        encodingStatusBar.registerCommand(),
        vscode.commands.registerCommand('office.quickOpen', () => vscode.commands.executeCommand('workbench.action.quickOpen')),
        vscode.commands.registerCommand('office.markdown.switch', (uri) => { markdownService.switchEditor(uri); }),
        vscode.commands.registerCommand('office.markdown.paste', () => { markdownService.loadClipboardImage(); }),
        vscode.commands.registerCommand('office.html.preview', uri => HtmlService.previewHtml(uri, context)),
        vscode.commands.registerCommand('office.excel.diffWithVCS', (uri) => excelDiffProvider.diffWithVCS(uri)),
        vscode.commands.registerCommand('office.excel.diffWithFile', (uri) => excelDiffProvider.diffWithFile(uri)),
        vscode.commands.registerCommand('office.excel.diffWithRevision', (uri) => excelDiffProvider.diffWithRevision(uri)),
        vscode.window.registerCustomEditorProvider("cweijan.markdownViewer", markdownEditorProvider, viewOptionWithFind),
        vscode.window.registerCustomEditorProvider("cweijan.markdownViewer.optional", markdownEditorProvider, viewOptionWithFind),
        ...viewerInstance.bindCustomEditors(viewOption)
    );
}

export function deactivate() {}

async function activeHTTP(context: vscode.ExtensionContext) {
    try {
        httpExt.activate(context);
    } catch (error) {
        Output.debug(error);
    }
}

function keepOriginDiff() {
    try {
        const config = vscode.workspace.getConfiguration("workbench");
        const configKey = 'editorAssociations';
        const editorAssociations = config.get(configKey) as Record<string, string | undefined>;
        const key = '{git,gitlens,git-graph}:/**/*.{md,csv,svg}';
        if (editorAssociations && editorAssociations[key]) {
            editorAssociations[key] = undefined;
            config.update(configKey, editorAssociations, true);
        }
    } catch (error) {
        Output.debug('keepOriginDiff failed: ' + error);
    }
}
```

- [ ] **Step 2: 编译检查**

Run: `cd f:/Dev/vscode-office && npx tsc --noEmit -p tsconfig.json`
Expected: 无 error

- [ ] **Step 3: Commit**

```bash
git add src/extension.ts
git commit -m "feat: wire up GitApiHelper and Office SCM contribution"
```

---

## Task 11: package.json 配置项

**Files:**
- Modify: `package.json`

- [ ] **Step 1: 在 `contributes.configuration.properties` 中追加三项**

找到 `"vscode-office.pasterImgPath": { ... }` 之后,在该对象闭合 `}` 之前追加:

```json
,
"vscode-office.blame.depth": {
    "type": "number",
    "default": 20,
    "minimum": 1,
    "maximum": 100,
    "description": "How many commits back to scan when computing cell blame."
},
"vscode-office.scm.enableOfficePanel": {
    "type": "boolean",
    "default": true,
    "description": "Show a dedicated SCM panel for Office files (xlsx/csv/...)."
},
"vscode-office.scm.includedExtensions": {
    "type": "array",
    "default": ["xlsx", "xlsm", "xls", "csv", "ods"],
    "description": "File extensions included in the Office SCM panel."
}
```

- [ ] **Step 2: 验证 JSON 合法**

Run: `cd f:/Dev/vscode-office && node -e "JSON.parse(require('fs').readFileSync('package.json','utf8'))"`
Expected: 无输出(无报错)即合法

- [ ] **Step 3: Commit**

```bash
git add package.json
git commit -m "feat: add config for blame depth and Office SCM panel"
```

---

## Task 12: ExcelDiff.tsx 适配新协议(ref switcher + blame)

**Files:**
- Modify: `src/react/view/excel/ExcelDiff.tsx`
- Modify: `src/react/view/excel/ExcelDiff.less`

- [ ] **Step 1: 读取当前 ExcelDiff.tsx 完整内容,理解现有 useEffect/handler 结构**

Run: 由实施者使用 Read 工具读取该文件全文(参考: [src/react/view/excel/ExcelDiff.tsx](../../src/react/view/excel/ExcelDiff.tsx))

- [ ] **Step 2: 修改 init 消息处理,适配新协议**

旧逻辑:webview 收到 `openDiff` 后从 `content.currentPath / baseData / ext / encoding / baseLabel / currentLabel` 渲染。
新协议:`leftRef, rightRef, leftData, rightData, ext, encoding`。两边都是 base64 而不是 URL。

具体修改点在 `ExcelDiff.tsx` 的消息监听 `useEffect`(handler.on / handler.emit)中:

```tsx
useEffect(() => {
    handler.emit('init');
    const off = handler.on('openDiff', (payload: any) => {
        const { leftRef, rightRef, leftData, rightData, ext: fileExt, encoding: enc } = payload;
        setLeftRef(leftRef);
        setRightRef(rightRef);
        diffDataRef.current = { leftRef, rightRef, leftData, rightData, ext: fileExt };
        renderFromBase64(leftData, rightData, fileExt, enc);
    });
    const offUpdate = handler.on('updateSide', (payload: any) => {
        const { side, ref, data } = payload;
        if (side === 'left') { setLeftRef(ref); diffDataRef.current!.leftData = data; }
        else { setRightRef(ref); diffDataRef.current!.rightData = data; }
        const dd = diffDataRef.current!;
        renderFromBase64(dd.leftData, dd.rightData, dd.ext, encoding);
    });
    const offBlame = handler.on('blameResult', (payload: any) => {
        showBlamePopover(payload);
    });
    return () => { off(); offUpdate(); offBlame(); };
}, []);
```

新增 helpers(仿照原有 `renderDiff` 但 base64→buffer 在两边)+ state:

```tsx
const [leftRef, setLeftRef] = useState<any>(null);
const [rightRef, setRightRef] = useState<any>(null);
const [blamePopover, setBlamePopover] = useState<{ sheet:string; row:number; col:number; entry?:any; error?:string } | null>(null);

function refLabel(r: any): string {
    if (!r) return '';
    switch (r.kind) {
        case 'working': case 'svn-working': return 'Working';
        case 'staged': return 'Staged';
        case 'head': return 'HEAD';
        case 'commit': return r.hash.substring(0, 7);
        case 'stash': return `stash@{${r.index}}`;
        case 'file': return (r.uri?.path || '').split('/').pop() || 'file';
        case 'svn-revision': return `r${r.revision}`;
        default: return '?';
    }
}

const renderFromBase64 = useCallback((leftData: string, rightData: string, fileExt: string, enc: string) => {
    setLoading(true);
    setTimeout(() => {
        const b2buf = (b64: string) => {
            const bin = atob(b64);
            const u8 = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
            return u8.buffer;
        };
        const leftBuffer = b2buf(leftData);
        const rightBuffer = b2buf(rightData);
        currentBufferRef.current = rightBuffer;
        const baseExcel = loadSheets(leftBuffer, fileExt, enc);
        const currentExcel = loadSheets(rightBuffer, fileExt, enc);
        workbookRef.current = currentExcel.workbook || null;
        const diffResult = computeDiff(baseExcel, currentExcel);
        setStats(diffResult.stats);
        sheetsRef.current = { left: diffResult.leftSheets, right: diffResult.rightSheets };
        const rows = diffResult.leftSheets[0]?.changeRows || [];
        changeRowsRef.current = rows;
        setTotalChanges(rows.length);
        setChangeIndex(-1);
        renderSide(document.getElementById('diff-left')!, diffResult.leftSheets, 'left');
        renderSide(document.getElementById('diff-right')!, diffResult.rightSheets, 'right', true);
        setupScrollSync(document.getElementById('diff-left')!, document.getElementById('diff-right')!);
        setLoading(false);
    }, 16);
}, []);
```

在 toolbar JSX 中,旧的 `baseLabel/currentLabel` span 替换成两个可点击的 button:

```tsx
<button className="ref-switch" onClick={() => handler.emit('pickRef', { side: 'left' })}>
    {refLabel(leftRef)} ▼
</button>
<span className="ref-arrow">↔</span>
<button className="ref-switch" onClick={() => handler.emit('pickRef', { side: 'right' })}>
    {refLabel(rightRef)} ▼
</button>
```

cell 右键 → blame 触发(在 spreadsheet right side 的 oncellcontextmenu 钩子,或在 toolbar 加 "Blame selected cell" 按钮):

```tsx
const onBlameClick = useCallback(() => {
    const sel = rightRef?.current?.getSelector?.();
    if (!sel) return;
    const { ri, ci, sheet } = sel;
    handler.emit('requestBlame', { sheet: sheet || 'Sheet1', row: ri, col: ci });
}, []);
```

blame popover 显示组件:

```tsx
{blamePopover && (
    <div className="blame-popover" onClick={() => setBlamePopover(null)}>
        {blamePopover.error
            ? <span>{blamePopover.error}</span>
            : <span>{blamePopover.entry.hash.substring(0,7)} · {blamePopover.entry.author} · {blamePopover.entry.date.substring(0,10)} — {blamePopover.entry.message}</span>}
    </div>
)}
```

`showBlamePopover` 函数:`setBlamePopover({ sheet:msg.sheet, row:msg.row, col:msg.col, entry: msg.entry, error: msg.error })`。

- [ ] **Step 3: 添加 ExcelDiff.less 样式**

在文件末尾追加:

```less
.ref-switch {
    background: var(--vscode-button-secondaryBackground, #3a3d41);
    color: var(--vscode-button-secondaryForeground, #fff);
    border: 1px solid var(--vscode-button-border, transparent);
    padding: 2px 8px;
    margin: 0 4px;
    cursor: pointer;
    border-radius: 3px;
    font-family: var(--vscode-editor-font-family, monospace);
    font-size: 12px;
    &:hover { background: var(--vscode-button-secondaryHoverBackground, #45494e); }
}
.ref-arrow {
    color: var(--vscode-foreground);
    margin: 0 4px;
}
.blame-popover {
    position: fixed;
    bottom: 12px;
    right: 12px;
    max-width: 60vw;
    padding: 8px 12px;
    background: var(--vscode-editorHoverWidget-background);
    color: var(--vscode-editorHoverWidget-foreground);
    border: 1px solid var(--vscode-editorHoverWidget-border);
    border-radius: 4px;
    z-index: 100;
    cursor: pointer;
    font-size: 12px;
}
```

- [ ] **Step 4: 构建并检查**

Run: `cd f:/Dev/vscode-office && npm run build`
Expected: 构建成功,无 TS 错误

- [ ] **Step 5: Commit**

```bash
git add src/react/view/excel/ExcelDiff.tsx src/react/view/excel/ExcelDiff.less
git commit -m "feat(diff-ui): ref switcher dropdown and cell blame popover"
```

---

## Task 13: 端到端手测(Extension Development Host)

**Files:** 无(仅手测验证)

- [ ] **Step 1: 启动 Extension Development Host**

Run(在 vscode 中按 F5,或命令行):
```bash
cd f:/Dev/vscode-office
# 在 vscode 打开本仓库,F5 启动 [Extension Development Host]
```

- [ ] **Step 2: 验证 diffWithVCS (HEAD ↔ Working)**

在 EDH 中打开本仓库:
1. 在 `TestData/Client_Text.xlsx` 上右键 → "Compare with VCS Version"
2. 期望:打开 diff 面板,左侧显示 HEAD 内容,右侧显示 Working 内容
3. 修改 xlsx 中一个 cell,保存,触发 diff,确认变更行显示

- [ ] **Step 3: 验证任意两端切换**

在 diff 面板中:
1. 点击左侧 `HEAD ▼` button → 出现 QuickPick(Working Tree / History / Stash / External 分组)
2. 选择 `History` 下任意一个 commit
3. 期望:左侧内容更新到该 commit 的内容,面板 title 同步更新
4. 点击右侧 button → 选 Browse file → 选择 `TestData/Leechdom_Dream.csv`
5. 期望:右侧切到该外部文件

- [ ] **Step 4: 验证 SCM 面板**

1. 在源代码管理视图(Ctrl+Shift+G)中,确认出现 "Office Files (vscode-office)" 面板
2. 修改 TestData 下 xlsx 文件
3. 期望:Office Files 面板自动出现该文件 entry
4. 点击 entry → 打开 diff 面板

- [ ] **Step 5: 验证 Cell Blame**

1. 提交 `TestData/Client_Text.xlsx` 当前修改到一个临时 commit
2. 再修改另一个 cell,打开 diff
3. 点击右侧 toolbar "Blame" 按钮(或右键单元格)
4. 期望:右下角 popover 显示该 cell 最后一次修改的 commit hash + author + date + message

- [ ] **Step 6: 验证 SVN(可选,需 svn 客户端)**

跳过 if 无 svn 工作副本。否则对一个 csv 文件触发 diffWithVCS → 应回到 svn revision 内容。

- [ ] **Step 7: 修复任何 manual 测试中发现的问题**

若有问题,根据具体错误调整对应 Task 文件并 commit fix。

- [ ] **Step 8: 最终 commit(若有 fix)**

```bash
git add <changed files>
git commit -m "fix(vcs): adjustments from end-to-end testing"
```

---

## Self-Review 结果

- ✅ Spec section 4.1 模块划分 → Tasks 2-8 一对一覆盖
- ✅ Section 4.2 Ref 模型 → Task 2
- ✅ Section 4.4 Git API → Task 3
- ✅ Section 4.5/4.6 Resolver+CLI → Tasks 4-5
- ✅ Section 4.7 Picker → Task 6
- ✅ Section 4.8 Blame → Task 7
- ✅ Section 4.9 SCM → Task 8
- ✅ Section 5 webview 协议 → Task 12
- ✅ Section 6 配置 → Task 11
- ✅ Section 7 命令 → 已存在,Task 10 重新绑定
- ✅ Section 8 验证 → Task 13
- ✅ 无 TBD/placeholder,所有代码完整
- ✅ 命名一致(GitApiHelper / VcsResolver / CliFallback / BlameProvider / OfficeScmContribution)
