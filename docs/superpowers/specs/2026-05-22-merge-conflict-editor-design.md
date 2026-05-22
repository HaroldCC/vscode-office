# 3-Way Merge / Conflict Editor Design (Subsystem 3 of 5)

> 本 spec 仅覆盖 **子系统 3 — xlsx/csv 三方合并冲突编辑器**。

## 1. Context

git merge / rebase / cherry-pick 遇到 xlsx/csv 冲突时,工作区会出现 conflicted 状态。对文本文件 vscode 内置 mergeEditor(BASE/LOCAL/REMOTE/RESULT)能直接介入,但 xlsx 是二进制 → vscode 默认会拒绝合并,csv 虽然是文本但 git 会插入 `<<<<<<<` 标记导致 sheet 无法正确解析。

本子系统提供 **表格感知的三方合并编辑器**:从 git 的 :1: / :2: / :3: 取出 BASE/OURS/THEIRS,在自建 webview 里以 3 列(或 2+1)视图呈现冲突单元格,用户逐 cell 选择并写回 working file。

## 2. Goals

1. **冲突检测**:活动文件出现 git mergeChanges 中 → 自动建议打开 Office Merge Editor
2. **三方加载**:`git show :1:path` / `:2:path` / `:3:path` 拉出 BASE/OURS/THEIRS
3. **合并 UI**:
   - 三列表格视图(BASE | OURS ← Result → THEIRS)
   - 自动合并:base/ours/theirs 中两端相同的 cell 自动选定值
   - 冲突 cell:高亮黄,提供按钮 "Take Ours" / "Take Theirs" / "Take Both"(拼接)/ "Edit Manually"
4. **CSV 中 git 标记的清理**:打开 csv 时若检测到 `<<<<<<<` 标记,提示用户使用 Office Merge Editor 替代直接编辑
5. **完成后**:写回 working file,并执行 `git add <path>` 标记冲突已解决
6. **命令**:`office.merge.openConflictEditor` 触发,SCM 面板 merge changes 区域条目点击直接打开

## 3. Non-Goals

- 公式级别合并(`SUM(A1:A10)` 与 `SUM(A1:A11)` 视为字符串差异,不重新计算)
- 样式合并(颜色、字体改动不参与冲突计算)
- 跨 sheet 的结构性合并(增删 sheet 提示用户手动处理)

## 4. Architecture

### 4.1 新增模块

```
src/provider/merge/
├── mergeDetector.ts       # 检测 mergeChanges + 自动提示
├── mergeLoader.ts         # git show :1: :2: :3: 加载三方
├── cellMerge.ts           # 三方逐 cell 合并算法 + 冲突分类
└── mergeEditorProvider.ts # webview 启动 + 写回 + git add

src/react/view/excel/
├── ExcelMerge.tsx         # 三方合并 UI(三列 + Result)
└── ExcelMerge.less        # 样式
```

### 4.2 三方合并算法(`cellMerge.ts`)

输入:`baseSheet`, `oursSheet`, `theirsSheet`(都是 `string[][]`)

输出:
```typescript
interface MergedSheet {
    name: string;
    /** result[r][c] — 当前选定值;初始填充 auto-resolve 后的结果 */
    result: string[][];
    /** 每个 cell 的状态 */
    status: CellStatus[][];
    conflicts: ConflictPosition[];   // 待用户处理的位置
}

type CellStatus =
    | 'unchanged'         // base = ours = theirs
    | 'auto-ours'         // theirs = base ≠ ours → 自动取 ours
    | 'auto-theirs'       // ours = base ≠ theirs → 自动取 theirs
    | 'auto-same'         // ours = theirs ≠ base → 两边改成一样,自动取
    | 'conflict'          // 三个都不一样 / 一端删除一端改
    | 'resolved-ours'     // 用户选了 ours
    | 'resolved-theirs'   // 用户选了 theirs
    | 'resolved-both'     // 用户选了拼接
    | 'resolved-manual';  // 用户手动编辑
```

算法:对每个 cell 跑 `classify(base, ours, theirs)`,得到 status 与初始 result 值。冲突位置入 `conflicts` 数组(给 UI 导航用)。

对于行数不一致:用最大行数,缺失行视为空值(同样进入分类)。

### 4.3 MergeLoader

```typescript
class MergeLoader {
    static async load(uri: vscode.Uri): Promise<{ base: Buffer; ours: Buffer; theirs: Buffer }> {
        const info = await CliFallback.gitInfo(uri);
        if (!info) throw new Error('Not in a git repository');
        const [base, ours, theirs] = await Promise.all([
            CliFallback.gitShow(uri, ':1').catch(() => Buffer.alloc(0)),
            CliFallback.gitShow(uri, ':2'),
            CliFallback.gitShow(uri, ':3'),
        ]);
        return { base, ours, theirs };
    }
}
```

> 注:`git show :1:path` 在不存在 BASE(添加冲突)时会失败,返回空 buffer 表示该端不存在。

需要在 `CliFallback.gitShow` 支持 `:1` `:2` `:3` 这种 stage 写法(目前已经支持 `:` 作为 staged,扩展到 `:N`):

```typescript
const spec = /^:(\d+)$/.test(ref) ? `${ref}:${info.relPath}` : `${ref}:${info.relPath}`;
// stage 0 (`:`) ⇒ `:path`;stage 1 (`:1`) ⇒ `:1:path`
```

实际上 `git show :1:path` 这个语法 git 原生支持,不需要特殊改动,只要 `ref === ':1'` 时拼成 `:1:path` 即可。

### 4.4 MergeDetector

监听 vscode.git `repo.state.onDidChange`,在 `mergeChanges` 出现 xlsx/csv 时:

```typescript
// 弹一次性的通知(不重复打扰)
const choice = await vscode.window.showWarningMessage(
    `${count} office files have merge conflicts. Open Office Merge Editor?`,
    'Open',
    'Later',
);
if (choice === 'Open') {
    for (const c of officeConflicts) {
        vscode.commands.executeCommand('office.merge.openConflictEditor', c.uri);
    }
}
```

### 4.5 MergeEditorProvider

类似 ExcelDiffProvider 的结构,但 webview 路由 `'excel-merge'`,UI 是 3 列 + 顶部 Result 概览:

```
┌──────────────────────────────────────────────────────────────────┐
│ Merge: data.xlsx                              [Mark resolved]    │
│ Conflicts: 3 unresolved, 12 auto-resolved                        │
├──────────────────────┬──────────────────────┬───────────────────┤
│ BASE                 │ OURS (current branch)│ THEIRS (incoming) │
│ (read-only)          │ (read-only, click   │ (read-only, click │
│                      │  to take this)      │  to take this)    │
├──────────────────────┴──────────────────────┴───────────────────┤
│ RESULT (editable, what will be written)                          │
└──────────────────────────────────────────────────────────────────┘
```

### 4.6 写回与 git add

完成后:
1. xlsx:用现有 `excel_writer.export_xlsx` 把 result 表格写成 buffer
2. csv:按 result 序列化,使用选定 encoding
3. `vscode.workspace.fs.writeFile(uri, buffer)`
4. 调用 `Repository.add([uri])`(Git API)或 `execFile('git', ['add', '--', relPath])`

### 4.7 与子系统 1 的关系

- 复用 `VcsResolver.detect` / `CliFallback.gitShow`
- 复用 `excel_reader.loadSheets` 解析 buffer
- 复用 `excel_writer.export_xlsx` 写回

## 5. 命令贡献

| 命令 | 触发 |
|---|---|
| `office.merge.openConflictEditor(uri?)` | mergeChanges 通知按钮、手动触发、SCM merge changes 点击 |
| `office.merge.markResolved(uri?)` | 编辑器内顶部按钮(写回 + git add) |

## 6. 配置项

```jsonc
{
  "vscode-office.merge.autoPromptOnConflict": { "type": "boolean", "default": true, "description": "Auto-prompt to open Office Merge Editor when xlsx/csv conflicts appear." }
}
```

## 7. 验证

| 功能 | 测试 |
|---|---|
| 制造冲突 | 在 TestData 中复制 csv,在两个 branch 各改不同 cell,merge,出现冲突 |
| 自动提示 | 冲突出现后,通知出现 |
| 三方加载 | Merge Editor 打开,BASE/OURS/THEIRS 三列正确显示 |
| 自动合并 | 仅一方改的 cell 自动绿色 |
| 冲突 cell | 三方都不同的 cell 黄色,点击 Take Ours/Theirs 切换 |
| 写回 + git add | 标记 resolved 后 file 写回,git status 显示该文件已 staged |

## 8. 风险

- 大文件(>10MB)三份解析内存压力 → 限制:仅前两个 sheet 全量,其余按需
- 公式被当字符串差异 → 用户教育:在 README 说明
- 行数大幅变化(插入/删除行)→ 行级别就分类为冲突,UI 显示插入/删除标记
