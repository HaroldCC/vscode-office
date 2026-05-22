# Tortoise Integration Design (Subsystem 2 of 5)

> 目标全景:将本插件优化为高可用的 csv/xlsx 编辑、diff 工具,深度融合 vscode 的 git,并支持 TortoiseSVN/TortoiseGit 的 diff、编辑冲突,以及 vscode git 的冲突编辑工具。
>
> 本 spec 仅覆盖 **子系统 2 — Tortoise(SVN/Git) 工具集成**。

## 1. Context

TortoiseSVN 与 TortoiseGit 是 Windows 上广泛使用的图形 VCS 客户端,提供 shell 集成与外部 diff 工具协议。用户的两个诉求:

1. **被调起**:在 Tortoise* 中右键 xlsx/csv → "Diff with previous version" → 本插件能作为外部工具被启动并对比文件
2. **主动调起**:vscode 中右键 xlsx/csv → 调起 Tortoise* 的 "Show Log / Diff / Blame" 窗口(查看分支图、提交说明等富 UI)

本子系统建立在 [子系统 1](2026-05-22-vcs-deep-integration-design.md) 的 VcsResolver 基础上。

## 2. Goals

1. 提供 CLI 入口 `office-diff` 二进制(其实是 vscode 命令转 URI),让 Tortoise* 把它配置为外部 diff 工具
2. 实现外部启动协议:`code --command office.excel.diffWithBuffers --base <baseFile> --mine <mineFile> --title <text>`
3. 自动检测 TortoiseSVN/TortoiseGit 安装路径(注册表 + 默认安装目录)
4. 提供命令:
   - `office.tortoise.showLog` — 调起 TortoiseGit/SVN 的 Log Dialog
   - `office.tortoise.diff` — 调起 Tortoise 的原生 diff(用户偏好 Tortoise UI 时)
   - `office.tortoise.blame` — 调起 Tortoise 的 Blame
5. 在编辑器右键菜单中按 OS+VCS 类型有条件地显露这些命令(`when` clause)

## 3. Non-Goals

- 跨平台 Tortoise(macOS / Linux 没有 Tortoise,本子系统在非 Windows 上整体不激活)
- 替换 Tortoise 自身的提交对话框(我们只调用它们,不复制)

## 4. Architecture

### 4.1 新增模块

```
src/provider/tortoise/
├── tortoiseLocator.ts     # 探测 TortoiseSVN/TortoiseGit 安装路径
├── tortoiseProc.ts        # 启动 TortoiseProc.exe / TortoiseGitProc.exe 的封装
└── externalDiffCommand.ts # 处理被外部工具调起场景(office.excel.diffWithBuffers)
```

### 4.2 TortoiseLocator

按优先级:
1. 用户配置 `vscode-office.tortoise.svnPath` / `tortoiseGitPath` 显式指定
2. 环境变量 `PATH` 中的 `TortoiseProc.exe` / `TortoiseGitProc.exe`
3. Windows 注册表 `HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN`(`ProcPath` 键)和 `HKLM\SOFTWARE\TortoiseGit`(`ProcPath`)
4. 默认安装路径 `C:\Program Files\TortoiseSVN\bin\TortoiseProc.exe` / `C:\Program Files\TortoiseGit\bin\TortoiseGitProc.exe`

注册表读取通过 `reg query` 命令(纯文本)避免引入 native 依赖。

### 4.3 TortoiseProc 命令封装

TortoiseSVN: `TortoiseProc.exe /command:<cmd> /path:"<file>"`
TortoiseGit: `TortoiseGitProc.exe /command:<cmd> /path:"<file>"`

支持的 `<cmd>`:`diff` / `log` / `blame` / `revert` / `commit` / `update`

调用方式:`child_process.spawn`,不阻塞 vscode。

### 4.4 外部 Diff 协议(被调起场景)

Tortoise 的外部 diff 工具配置格式:

```
TortoiseSVN: Settings → Diff Viewer → External:
  C:\Path\To\Code.exe --new-window --command office.excel.diffWithFiles --base "%base" --mine "%mine" --base-name "%bname" --mine-name "%yname"
```

但 vscode 不接受任意 `--command + 参数`,所以采用环境变量传递 + 启动命令:

```bash
# 包装脚本 office-diff.cmd(随插件分发):
set OFFICE_DIFF_BASE=%1
set OFFICE_DIFF_MINE=%2
set OFFICE_DIFF_BASE_NAME=%3
set OFFICE_DIFF_MINE_NAME=%4
code --reuse-window <BASE_FILE>
```

Vscode 启动时,我们的扩展激活,**检查环境变量 `OFFICE_DIFF_BASE`**:存在 → 自动以 base/mine 启动 diff 面板。

不同方案对比:

| 方案 | 复杂度 | 用户配置 |
|---|---|---|
| A. URI handler(`vscode://rjwang.vscode-office-enhanced/diff?base=...&mine=...`) | 低 | 在 Tortoise 中配置 cmd 调 `start vscode://...` |
| B. 包装脚本 + 环境变量 | 中 | 配置脚本路径作为外部 diff |
| C. CLI 子命令 | 高 | 需要 vscode 支持自定义 startup CLI |

**采用 A(URI handler)**,vscode 已经内建 URI handler 注册机制,跨平台一致。

### 4.5 URI Handler

注册 `vscode.window.registerUriHandler`,解析 `vscode://rjwang.vscode-office-enhanced/diff?base=...&mine=...&baseName=...&mineName=...`:

```typescript
class ExternalDiffUriHandler implements vscode.UriHandler {
  handleUri(uri: vscode.Uri) {
    const params = new URLSearchParams(uri.query);
    const base = params.get('base');
    const mine = params.get('mine');
    if (!base || !mine) return;
    excelDiffProvider.diffWithExternalFiles(
      vscode.Uri.file(base),
      vscode.Uri.file(mine),
      params.get('baseName') || 'BASE',
      params.get('mineName') || 'MINE',
    );
  }
}
```

`ExcelDiffProvider.diffWithExternalFiles(baseUri, mineUri, baseLabel, mineLabel)` 是 `openDiff` 的薄封装:左侧 Ref `{ kind: 'file', uri: baseUri, label: baseLabel }`,右侧 Ref `{ kind: 'file', uri: mineUri, label: mineLabel }`。

用户在 Tortoise 设置中配置(平台命令格式):

**TortoiseSVN(Settings → Diff Viewer → External):**
```
cmd.exe /c "start vscode://rjwang.vscode-office-enhanced/diff?base=%base&mine=%mine&baseName=%bname&mineName=%yname"
```

**TortoiseGit(Settings → Diff Viewer → External diff tool for .xlsx):**
```
cmd.exe /c "start vscode://rjwang.vscode-office-enhanced/diff?base=%base&mine=%mine"
```

文档化在 README,通过 `office.tortoise.configureExternal` 命令一键将上述行写入剪贴板 + 打开 Tortoise 设置页提示。

### 4.6 主动调用 Tortoise(右键 → Tortoise Log/Diff/Blame)

新增命令:
- `office.tortoise.svnLog` / `office.tortoise.gitLog`
- `office.tortoise.svnDiff` / `office.tortoise.gitDiff`
- `office.tortoise.svnBlame` / `office.tortoise.gitBlame`

每个命令:
1. `VcsResolver.detect(uri)` → 判断 svn/git
2. 选对应 TortoiseLocator 的 exe path
3. spawn TortoiseProc.exe /command:<cmd> /path:"<filePath>"

UI 入口(`package.json` 的 `menus.editor/title` 与 `explorer/context`):
```jsonc
{
  "command": "office.tortoise.svnLog",
  "when": "isWindows && resourceExtname =~ /\\.(xlsx|xls|csv|xlsm|ods)/i",
  "group": "3_compare@4"
}
// ... 其他类似
```

`when` 中 `isWindows` 通过 `context.isWindows` 上下文键(activation 时根据 `process.platform === 'win32'` 设置)。

### 4.7 与子系统 1 的关系

- 复用 `VcsResolver.detect` 判断仓库类型
- 复用 `ExcelDiffProvider.openDiff`(将其 `openDiff` 改为 public 或导出一个 wrapper)
- 不依赖 `BlameProvider`,Tortoise blame 是 Tortoise 自己的 UI

## 5. 配置项

```jsonc
{
  "vscode-office.tortoise.svnPath": { "type": "string", "description": "Path to TortoiseProc.exe (auto-detected if empty)" },
  "vscode-office.tortoise.gitPath": { "type": "string", "description": "Path to TortoiseGitProc.exe (auto-detected if empty)" },
  "vscode-office.tortoise.enabled": { "type": "boolean", "default": true, "description": "Show Tortoise commands in context menus (Windows only)" }
}
```

## 6. 命令贡献(package.json)

```jsonc
{
  "command": "office.tortoise.svnLog",       "title": "TortoiseSVN: Show Log", "category": "Office Viewer"
},
{ "command": "office.tortoise.gitLog",       "title": "TortoiseGit: Show Log" },
{ "command": "office.tortoise.svnDiff",      "title": "TortoiseSVN: Diff with HEAD" },
{ "command": "office.tortoise.gitDiff",      "title": "TortoiseGit: Diff with HEAD" },
{ "command": "office.tortoise.svnBlame",     "title": "TortoiseSVN: Blame" },
{ "command": "office.tortoise.gitBlame",     "title": "TortoiseGit: Blame" },
{ "command": "office.tortoise.configureExternal", "title": "Office: Configure as Tortoise External Diff Tool" }
```

## 7. 验证方案

| 功能 | 测试步骤 | 期望 |
|---|---|---|
| Tortoise 路径自动探测 | Windows 上安装 TortoiseGit,启动 EDH,执行命令 `Office: Configure as Tortoise External Diff Tool` | 自动找到路径,弹出对话框确认 |
| 主动调起 TortoiseGit Log | git 仓库下 xlsx 右键 → "TortoiseGit: Show Log" | TortoiseGitProc 打开 log 对话框 |
| 主动调起 TortoiseSVN Diff | svn 工作副本下 xlsx 右键 → "TortoiseSVN: Diff with HEAD" | TortoiseProc 打开 diff 对话框 |
| URI handler 被调起 | 在浏览器 `vscode://rjwang.vscode-office-enhanced/diff?base=...&mine=...` | vscode 启动并打开 diff 面板 |
| 非 Windows 不显示菜单 | 在 macOS/Linux 上(无法测) | when clause 屏蔽 |

## 8. 受影响文件

新增:
- `src/provider/tortoise/tortoiseLocator.ts`
- `src/provider/tortoise/tortoiseProc.ts`
- `src/provider/tortoise/externalDiffCommand.ts`
- `src/provider/tortoise/index.ts`(barrel + activate)

修改:
- `src/extension.ts`(注册 URI handler + Tortoise 命令 + 设置 isWindows context key)
- `src/provider/excelDiffProvider.ts`(暴露 `diffWithExternalFiles(baseUri, mineUri, baseLabel, mineLabel)`)
- `package.json`(3 项配置 + 7 个命令 + menu 贡献)
- `README.md`(说明 Tortoise 集成方式)— 跳过文档,只口头提示

## 9. 风险

- TortoiseProc CLI 在不同 Tortoise 版本下参数兼容性:仅使用最稳定的 `/command:` + `/path:`,这两者从早期版本起一致
- 注册表读取在某些 Windows 上需要 admin:fall back 到默认路径
- URI handler 跨进程交互:vscode 已有现成机制,无需额外测试
