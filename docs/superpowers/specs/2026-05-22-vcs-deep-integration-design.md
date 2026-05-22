# VCS Deep Integration Design (Subsystem 1 of 5)

> 目标全景:将本插件优化为高可用的 csv/xlsx 编辑、diff 工具,深度融合 vscode 的 git,并支持 TortoiseSVN/TortoiseGit 的 diff、编辑冲突,以及 vscode git 的冲突编辑工具。
>
> 本 spec 仅覆盖 **子系统 1 — VCS 深度集成**。其它子系统(Tortoise 集成、三方合并冲突编辑器、Diff 工具增强、CSV/XLSX 编辑器健壮性)将各自独立成 spec。

## 1. Context

`src/provider/excelDiffProvider.ts` 当前以 `child_process.execFile` 直接调用 `git`/`svn` 命令,带来:

- 不识别 vscode workspace 中的多仓库(用 `dirname(filePath)` 作为 cwd 容易错)
- 无法响应 vscode 内置 git 扩展的事件(branch 切换、checkout 后 cache 不刷新)
- 不能在 SCM 面板里把 xlsx/csv 变更聚合给用户
- HEAD↔Working 之外的对比组合无法表达(staged / stash / 任意 commit/commit)

本子系统重构 VCS 访问层,接入 `vscode.git` Extension API(混合方案 B:状态/资源走 API,内容读取走 CLI fallback),并把"任意 Ref 对任意 Ref" 作为一等模型贯穿 webview 与扩展端。

## 2. Goals

1. 接入内置 Git Extension API,所有 git 状态/仓库定位走 API
2. 支持**任意两端**对比:Working / Staged / HEAD / 任意 commit / stash / 外部文件;SVN 支持 Working / 任意 revision
3. 独立 SCM 面板 `Office Files`,只列出 xlsx/csv/xlsm/xls/ods 变更
4. Diff 面板顶部支持**动态切换两端**,不需要重开面板
5. **单元格级 blame**:右键单元格按需查询并显示最近修改 commit

## 3. Non-Goals(交给后续子系统)

- TortoiseSVN/TortoiseGit 调用(子系统 2)
- 三方合并、冲突标记解析、vscode mergeEditor 联动(子系统 3)
- Diff 同步滚动、变更导航、UI 与原生对齐(子系统 4)
- 编辑器内部行为(子系统 5)

## 4. Architecture

### 4.1 模块划分

```
src/provider/vcs/
├── gitApi.ts          GitApiHelper 单例:lazy init + repoFor(uri) + watch
├── cliFallback.ts     execFile 包装:gitShow / gitBlame / gitStashList / svnCat / svnLog
├── vcsResolver.ts     resolveBuffer(uri, ref): Promise<Buffer> 统一入口
├── revisionPicker.ts  showRevisionPicker(uri): Promise<Ref | undefined>
├── blameProvider.ts   getCellBlame(uri, sheet, row, col): Promise<BlameEntry | null>
├── scmContribution.ts 独立 SourceControl 'office-excel-diff'
└── types.ts           Ref / BlameEntry / RepoInfo
```

`provider/excelDiffProvider.ts` 收缩为:接收命令 → 通过 `revisionPicker` / `vcsResolver` 解析 → 启动 webview → 转发消息。所有 `getGitVersion`/`getSvnVersion`/`getGitLog`/`getSvnLog`/`getGitFileAtRevision`/`getSvnFileAtRevision` 删除。

### 4.2 统一 Ref 模型

```typescript
// src/provider/vcs/types.ts
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
  hash: string;        // 'abc1234'
  author: string;
  email: string;
  date: string;        // ISO
  message: string;     // commit subject
}

export interface RepoInfo {
  kind: 'git' | 'svn' | null;
  root: string;                // 绝对路径
  relPath: string;             // 相对 root 的 posix path
  api?: { repo: any /* Repository from vscode.git */ };
}
```

### 4.3 数据流

```
Command office.excel.diffWithVCS(uri)
  ExcelDiffProvider.diffWithVCS(uri)
    info = await VcsResolver.detect(uri)             // git API 或 svn CLI 探测
    refA = info.kind === 'svn' ? { kind:'svn-working' } : { kind:'head' }
    refB = info.kind === 'svn' ? ??? : { kind:'working' }   // svn 必须由用户选,见 4.6
    [bufA, bufB] = await Promise.all([
       VcsResolver.resolveBuffer(uri, refA),
       VcsResolver.resolveBuffer(uri, refB),
    ])
    openDiffPanel(uri, bufA, bufB, { refA, refB, availableRefs: await listAvailableRefs(uri) })

Webview side switch:
  user clicks left ref dropdown
  → postMessage { type:'pickRef', side:'left' }
    → ext: ref = await RevisionPicker.show(uri)
    → buf = await VcsResolver.resolveBuffer(uri, ref)
    → webview.postMessage { type:'updateSide', side:'left', data: buf.toString('base64'), ref }
```

### 4.4 Git API 接入(`gitApi.ts`)

```typescript
import * as vscode from 'vscode';
import type { GitExtension, API, Repository } from '../../types/vscode-git';

export class GitApiHelper {
  private static _instance: GitApiHelper;
  private api: API | undefined;
  private initialized = false;

  static get instance(): GitApiHelper {
    return this._instance ??= new GitApiHelper();
  }

  async ensureInit(): Promise<API | undefined> {
    if (this.initialized) return this.api;
    this.initialized = true;
    const ext = vscode.extensions.getExtension<GitExtension>('vscode.git');
    if (!ext) return undefined;
    if (!ext.isActive) await ext.activate();
    this.api = ext.exports.getAPI(1);
    return this.api;
  }

  repoFor(uri: vscode.Uri): Repository | undefined {
    if (!this.api) return undefined;
    let best: Repository | undefined;
    let bestLen = -1;
    for (const r of this.api.repositories) {
      const root = r.rootUri.fsPath;
      if (uri.fsPath.startsWith(root + require('path').sep) || uri.fsPath === root) {
        if (root.length > bestLen) { best = r; bestLen = root.length; }
      }
    }
    return best;
  }

  repos(): Repository[] { return this.api?.repositories ?? []; }

  onDidOpenRepository(cb: (r: Repository) => void): vscode.Disposable {
    return this.api?.onDidOpenRepository(cb) ?? new vscode.Disposable(() => {});
  }
}
```

类型来源:从 vscode 仓库 `extensions/git/src/api/git.d.ts` 拷贝到本仓库 `src/types/vscode-git.d.ts`。

### 4.5 VcsResolver(`vcsResolver.ts`)

```typescript
export class VcsResolver {
  static async detect(uri: vscode.Uri): Promise<RepoInfo> {
    await GitApiHelper.instance.ensureInit();
    const repo = GitApiHelper.instance.repoFor(uri);
    if (repo) {
      const root = repo.rootUri.fsPath;
      const relPath = path.relative(root, uri.fsPath).replace(/\\/g, '/');
      return { kind: 'git', root, relPath, api: { repo } };
    }
    // svn fallback:exec `svn info` 探测
    try {
      await execFileP('svn', ['info', uri.fsPath]);
      return { kind: 'svn', root: path.dirname(uri.fsPath), relPath: path.basename(uri.fsPath) };
    } catch { /* fall through */ }
    return { kind: null, root: path.dirname(uri.fsPath), relPath: path.basename(uri.fsPath) };
  }

  static async resolveBuffer(uri: vscode.Uri, ref: Ref): Promise<Buffer> {
    switch (ref.kind) {
      case 'working':       return Buffer.from(await vscode.workspace.fs.readFile(uri));
      case 'file':          return Buffer.from(await vscode.workspace.fs.readFile(ref.uri));
      case 'staged':        return CliFallback.gitShow(uri, ':');           // git show :path
      case 'head':          return CliFallback.gitShow(uri, 'HEAD');
      case 'commit':        return CliFallback.gitShow(uri, ref.hash);
      case 'stash':         return CliFallback.gitShow(uri, `stash@{${ref.index}}`);
      case 'svn-working':   return Buffer.from(await vscode.workspace.fs.readFile(uri));
      case 'svn-revision':  return CliFallback.svnCat(uri, ref.revision);
    }
  }

  static async listAvailableRefs(uri: vscode.Uri): Promise<Ref[]> {
    const info = await this.detect(uri);
    if (info.kind === 'git') {
      const commits = await CliFallback.gitLog(uri, 20);
      const stashes = await CliFallback.gitStashList(info.root).catch(() => []);
      return [
        { kind: 'working' },
        { kind: 'staged' },
        { kind: 'head' },
        ...commits.map(c => ({ kind: 'commit', hash: c.hash, message: c.message, date: c.date } as Ref)),
        ...stashes.map((s, i) => ({ kind: 'stash', index: i, message: s.message } as Ref)),
      ];
    }
    if (info.kind === 'svn') {
      const revs = await CliFallback.svnLog(uri, 20).catch(() => []);
      return [
        { kind: 'svn-working' },
        ...revs.map(r => ({ kind: 'svn-revision', revision: r.hash, message: r.message, date: r.date } as Ref)),
      ];
    }
    return [{ kind: 'working' }];
  }
}
```

### 4.6 CliFallback(`cliFallback.ts`)

```typescript
export class CliFallback {
  /** ref 形如 'HEAD' | ':' (staged) | 'abc1234' | 'stash@{0}' */
  static async gitShow(uri: vscode.Uri, ref: string): Promise<Buffer> {
    const info = await VcsResolver.detect(uri);
    if (info.kind !== 'git') throw new Error('Not a git repo');
    const arg = ref === ':' ? `:${info.relPath}` : `${ref}:${info.relPath}`;
    return execFileBufferP('git', ['show', arg], { cwd: info.root });
  }
  static async gitLog(uri: vscode.Uri, n: number): Promise<{hash;message;date}[]> { ... }
  static async gitStashList(root: string): Promise<{message:string}[]> { ... }
  static async gitBlame(uri: vscode.Uri, line: number): Promise<BlameEntry | null> { ... }
  static async svnCat(uri: vscode.Uri, revision: string): Promise<Buffer> { ... }
  static async svnLog(uri: vscode.Uri, n: number): Promise<{hash;message;date}[]> { ... }
  static async svnInfo(uri: vscode.Uri): Promise<{ revision: string } | null> { ... }
}
```

所有命令的 cwd 一律使用 `RepoInfo.root`(从 Git API 拿到),修复当前实现中 `dirname(filePath)` 在子目录可能错位的问题。`maxBuffer: 50 * 1024 * 1024` 保持。

### 4.7 RevisionPicker(`revisionPicker.ts`)

```typescript
export async function showRevisionPicker(uri: vscode.Uri): Promise<Ref | undefined> {
  const refs = await VcsResolver.listAvailableRefs(uri);
  const items: (vscode.QuickPickItem & { ref: Ref })[] = [];

  const wt = refs.filter(r => r.kind === 'working' || r.kind === 'staged' || r.kind === 'svn-working');
  if (wt.length) {
    items.push({ label: 'Working Tree', kind: vscode.QuickPickItemKind.Separator } as any);
    for (const r of wt) items.push({ label: refLabel(r), description: refDesc(r), ref: r });
  }

  const commits = refs.filter(r => r.kind === 'commit' || r.kind === 'svn-revision');
  if (commits.length) {
    items.push({ label: 'History', kind: vscode.QuickPickItemKind.Separator } as any);
    for (const r of commits) items.push({
      label: r.kind === 'commit' ? r.hash.substring(0, 7) : `r${(r as any).revision}`,
      description: r.message,
      detail: r.date,
      ref: r,
    });
  }

  const stashes = refs.filter(r => r.kind === 'stash');
  if (stashes.length) {
    items.push({ label: 'Stash', kind: vscode.QuickPickItemKind.Separator } as any);
    for (const r of stashes) items.push({
      label: `stash@{${(r as any).index}}`,
      description: (r as any).message,
      ref: r,
    });
  }

  items.push({ label: 'External', kind: vscode.QuickPickItemKind.Separator } as any);
  items.push({ label: '$(file) Browse file…', ref: { kind: 'file', uri: vscode.Uri.parse('placeholder:') } });

  const sel = await vscode.window.showQuickPick(items, {
    placeHolder: 'Select revision to compare',
    matchOnDescription: true,
  });
  if (!sel) return undefined;

  if (sel.ref.kind === 'file') {
    const picked = await vscode.window.showOpenDialog({
      canSelectFiles: true, canSelectMany: false,
      filters: { 'Excel/CSV': ['xlsx','xls','xlsm','csv','ods'] }
    });
    if (!picked?.[0]) return undefined;
    return { kind: 'file', uri: picked[0] };
  }
  return sel.ref;
}
```

### 4.8 BlameProvider(`blameProvider.ts`)

二进制 xlsx 不能直接 `git blame`,采用**历史回放**算法:

```typescript
export class BlameProvider {
  private cache = new LruCache<string /* hash:path */, Map<string /* sheet */, string[][]>>(20);

  async getCellBlame(uri: vscode.Uri, sheet: string, row: number, col: number): Promise<BlameEntry | null> {
    const info = await VcsResolver.detect(uri);
    if (info.kind !== 'git') return null;
    const currentBuf = await VcsResolver.resolveBuffer(uri, { kind: 'working' });
    const currentVal = this.cellOf(this.parse(currentBuf, uri), sheet, row, col);

    const commits = (await CliFallback.gitLog(uri, 20));
    // 从最新到最旧找第一个值与 working 不同的 commit;再上一个就是设置成 current 的那次提交
    let lastSameCommit: { hash; message; date } | null = null;
    for (const c of commits) {
      const cacheKey = `${c.hash}:${info.relPath}`;
      let parsed = this.cache.get(cacheKey);
      if (!parsed) {
        const buf = await VcsResolver.resolveBuffer(uri, { kind: 'commit', hash: c.hash });
        parsed = this.parse(buf, uri);
        this.cache.set(cacheKey, parsed);
      }
      const val = this.cellOf(parsed, sheet, row, col);
      if (val !== currentVal) {
        // 当前 commit 与 working 已不同 → 设值的是上一个 commit (lastSameCommit)
        // 若 lastSameCommit 为空,表示当前值就是 working 里新加的,blame 不到
        return lastSameCommit ? await this.toEntry(lastSameCommit) : null;
      }
      lastSameCommit = c;
    }
    // 走完历史都相同
    return lastSameCommit ? await this.toEntry(lastSameCommit) : null;
  }

  private parse(buf: Buffer, uri: vscode.Uri): Map<string, string[][]> { /* XLSX.read or CSV parse */ }
  private cellOf(parsed: Map<string,string[][]>, sheet: string, row: number, col: number): string { ... }
  private async toEntry(c: {hash;message;date}): Promise<BlameEntry> { /* git show -s --format ... */ }
}
```

- 历史上限 20 commits(配置项 `vscode-office.blameDepth`)
- LRU 缓存 20 个 `(commit, file)` 解析结果
- 命中失败时返回 null,webview 显示 "Older than N commits, blame unavailable"

### 4.9 SCM 面板(`scmContribution.ts`)

```typescript
export class OfficeScmContribution {
  private scs = new Map<string /* repoRoot */, vscode.SourceControl>();
  private groups = new Map<string, vscode.SourceControlResourceGroup>();
  private disposables: vscode.Disposable[] = [];

  async activate(ctx: vscode.ExtensionContext): Promise<void> {
    const git = GitApiHelper.instance;
    const api = await git.ensureInit();
    if (!api) return;
    for (const repo of api.repositories) this.attach(repo);
    this.disposables.push(api.onDidOpenRepository(r => this.attach(r)));
    this.disposables.push(api.onDidCloseRepository(r => this.detach(r)));
    ctx.subscriptions.push({ dispose: () => this.disposables.forEach(d => d.dispose()) });
  }

  private attach(repo: any) {
    const root = repo.rootUri.fsPath;
    const sc = vscode.scm.createSourceControl('office-excel-diff', `Office Files (${path.basename(root)})`, repo.rootUri);
    const group = sc.createResourceGroup('changes', 'Changes');
    sc.quickDiffProvider = { provideOriginalResource: (uri) => this.quickDiffOriginal(uri) };
    this.scs.set(root, sc); this.groups.set(root, group);
    this.disposables.push(sc);
    this.disposables.push(repo.state.onDidChange(() => this.refresh(repo)));
    this.refresh(repo);
  }

  private refresh(repo: any) {
    const group = this.groups.get(repo.rootUri.fsPath);
    if (!group) return;
    const all = [...repo.state.workingTreeChanges, ...repo.state.indexChanges];
    const office = all.filter(c => /\.(xlsx|xlsm|xls|csv|ods)$/i.test(c.uri.fsPath));
    group.resourceStates = office.map(c => ({
      resourceUri: c.uri,
      command: { command: 'office.excel.diffWithVCS', title: 'Compare', arguments: [c.uri] },
      decorations: { tooltip: this.statusText(c.status), strikeThrough: c.status === 6 /* DELETED */ }
    }));
  }

  private quickDiffOriginal(uri: vscode.Uri): vscode.Uri | undefined {
    if (!/\.csv$/i.test(uri.fsPath)) return undefined;  // xlsx 二进制不提供 gutter
    return uri.with({ scheme: 'git', path: uri.path, query: JSON.stringify({ path: uri.fsPath, ref: '' }) });
  }
}
```

注:Office 自建 SCM 与原生 Git SCM 并存。原生面板仍能看到 xlsx 的变更标记;Office 面板提供"专属入口 + 一键 diff"。

## 5. Webview 协议变更(`ExcelDiff.tsx`)

新协议:

```typescript
// ext → webview
type ExtMsg =
  | { type: 'openDiff'; content: {
      leftRef: Ref; rightRef: Ref;
      leftData: string /* base64 */; rightData: string;
      ext: string; encoding: string;
      availableRefs: Ref[];
    }}
  | { type: 'updateSide'; side: 'left'|'right'; ref: Ref; data: string /* base64 */ }
  | { type: 'blameResult'; sheet: string; row: number; col: number; entry: BlameEntry | null }
  | { type: 'saveDone' };

// webview → ext
type WvMsg =
  | { type: 'init' }
  | { type: 'pickRef'; side: 'left'|'right' }
  | { type: 'requestBlame'; sheet: string; row: number; col: number }
  | { type: 'save'; content: any };
```

UI 改动:
- 顶部增加左右两个 ref 下拉 `[refLabel ▼] ↔ [refLabel ▼]`,点击发 `pickRef`
- 单元格右键菜单增加 "Show blame for this cell" → 发 `requestBlame`
- 收到 `blameResult` 在该 cell 上 popover 显示 `{shortHash} {author} · {relativeTime} — {message}`

## 6. 配置项

```jsonc
{
  "vscode-office.blame.depth": { "type": "number", "default": 20, "minimum": 1, "maximum": 100 },
  "vscode-office.scm.enableOfficePanel": { "type": "boolean", "default": true },
  "vscode-office.scm.includedExtensions": { "type": "array", "default": ["xlsx","xlsm","xls","csv","ods"] }
}
```

## 7. 命令贡献(`package.json`)

无新增 user-facing 命令;现有三个保留:
- `office.excel.diffWithVCS`(HEAD ↔ Working 默认入口)
- `office.excel.diffWithFile`
- `office.excel.diffWithRevision`(改为内部统一走 RevisionPicker)

激活事件追加 `onCommand:office.excel.*` 已由 `onStartupFinished` 覆盖。

## 8. 验证方案

| 功能 | 测试步骤 | 期望 |
|---|---|---|
| Git API 接入 | 多 root workspace,从 repo B 子目录的 xlsx 触发 diffWithVCS | cwd 正确,显示 repo B HEAD 内容 |
| 任意两端切换 | 打开 diff,左下拉选 stash@{0},右下拉选 abc1234 | 面板更新到对应内容,不重开 |
| Staged 对比 | `git add file.xlsx`,选 staged ↔ working | 显示已暂存的 vs 当前编辑 |
| SVN revision | 在 svn 工作副本对 csv 选 r123 ↔ working | 正确显示 |
| SCM 面板 | 修改 TestData/Client_Text.xlsx | Office Files 面板出现该 entry,点击进入 diff |
| Cell blame | 修改 Client_Text.xlsx 一个 cell,提交,再修改同 cell,右键 blame | popover 显示上一次提交的 hash/author/date |
| Blame 深度超限 | 一个 cell 30 个 commit 都没变,右键 blame | 显示 "older than 20 commits" |

## 9. 依赖

无新增 npm 依赖。

## 10. 受影响文件

新增:
- `src/types/vscode-git.d.ts`
- `src/provider/vcs/gitApi.ts`
- `src/provider/vcs/cliFallback.ts`
- `src/provider/vcs/vcsResolver.ts`
- `src/provider/vcs/revisionPicker.ts`
- `src/provider/vcs/blameProvider.ts`
- `src/provider/vcs/scmContribution.ts`
- `src/provider/vcs/types.ts`

修改:
- `src/extension.ts`(async activate,init GitApi + ScmContribution)
- `src/provider/excelDiffProvider.ts`(瘦身,委托 vcsResolver/revisionPicker/blameProvider)
- `src/react/view/excel/ExcelDiff.tsx`(顶部 ref switcher + blame UI + 新消息协议)
- `src/react/view/excel/ExcelDiff.less`(ref switcher 样式)
- `package.json`(三项新配置)

## 11. 风险

- vscode.git API 版本依赖:`getAPI(1)` 在 VSCode 1.64+ 稳定,与本插件 engines 要求一致
- Cell blame 算法对超大文件可能慢(20 次 XLSX.read);LRU 缓解,且仅按需触发
- 多 SourceControl 在某些 vscode 版本里 UI 不直观;通过配置项 `enableOfficePanel` 允许关闭
