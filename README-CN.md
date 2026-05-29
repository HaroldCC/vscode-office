# Office Suite Viewer

直接在 VS Code 中预览 Word、Excel、PDF、Markdown 等文件。支持 Excel/CSV 版本对比、单元格级 Blame、三方合并编辑器、TortoiseSVN/Git 集成、图片内嵌的独立 HTML 导出，以及 Mermaid v11 图表渲染。

## Office Viewer（Fork）

> Fork 自 [cweijan/vscode-office](https://github.com/cweijan/vscode-office)，原由 [RJ.Wang](https://github.com/rjwang1982) 维护，现由 [haroldcc](https://github.com/haroldcc) 维护。

本扩展在原项目基础上进行了 VCS 集成、Diff/Merge 工具链以及整体健壮性方面的大量增强，同时继承了 RJ.Wang 的既有改进。

## 本 Fork 的新功能

### VCS 深度集成（新增）
- **单元格级 Blame** — 查看 xlsx/csv 文件中每个单元格的最后修改者，基于历史回放和 LRU 缓存实现高性能。
- **Office SCM 面板** — 独立的 SCM 面板展示工作区中 xlsx/csv 文件变更。
- **版本对比** — 通过 QuickPick 版本选择器，将 Excel/CSV 文件与任意 Git/SVN 版本进行对比。
- **可配置的 Blame 深度** — `vscode-office.blame.depth` 控制 Blame 扫描的提交深度。

### Diff 增强（新增）
- **仅显示变更行** — 切换过滤器，仅显示有差异的行而非整个工作表。
- **HTML Diff 导出** — 将 Diff 结果导出为独立 HTML 文件以便分享。
- **固定表头** — 在大表格中滚动时列标题始终可见。
- **单元格级差异高亮** — 精确高亮有变化的单元格，而非整行着色。
- **版本切换下拉框** — 在 Diff 面板工具栏中快速切换对比版本。
- **单元格 Blame 弹窗** — 在 Diff 中悬停/点击单元格即可查看其最后提交信息。

### 三方合并编辑器（新增）
- 可视化三方合并编辑器（theirs → base ← ours），解决 xlsx/csv 合并冲突。
- 文件打开时自动检测冲突（`vscode-office.merge.autoPromptOnConflict`）。

### TortoiseSVN/TortoiseGit 集成（新增）
- 右键菜单命令：Show Log、Diff、Blame（仅 Windows）。
- 可将 Office Suite Viewer 配置为 Tortoise 的外部 Diff 工具。

### 健壮性改进（新增）
- **CSV 分隔符自动检测** — 自动识别逗号、分号、制表符、竖线等分隔符。
- **CSV 冲突检测** — 当 CSV 数据与检测到的格式冲突时主动发出 `saveError` 警告。
- **编码自动检测** — 状态栏编码选择器，无需重新加载即可实时切换编码。
- **延迟加载工作表** — 大文件（>10MB）XLSX 仅按需加载工作表，实现秒开。

### 编辑器体验（继承并增强）
- **独立 HTML 导出** — 本地图片自动转换为 Base64 内嵌，导出文件完全独立可分享。
- **Mermaid v11 支持** — 从 v8.8.0 升级到 v11.14.0，本地加载，离线更稳定。
- **可配置编辑模式** — WYSIWYG、即时渲染、分屏视图（`vscode-office.editorMode`）。
- **更清爽的渲染效果** — 内容铺满编辑器可用宽度，左对齐布局。
- **更小的安装包** — 移除内置 Icon Theme 和 Java Decompiler（约 4.4 MB），优化至约 5.25 MB。
- **One Dark Modern 主题** — 内置编辑器主题。

## 介绍

[English](README.md) | 简体中文

本扩展支持在 VS Code 中预览以下文件类型：

- Excel：`.xls`、`.xlsx`、`.csv`
- Word：`.docx`
- SVG：`.svg`
- PDF：`.pdf`
- 字体：`.ttf`、`.otf`、`.woff`、`.woff2`
- Markdown：`.md`
- HTTP 请求：`.http`
- Windows 注册表文件：`.reg`
- 压缩文件：`.zip`、`.jar`、`.vsix`、`.rar`

## Markdown

本扩展会使用 Vditor 替换默认的 Markdown 编辑器。请注意，Vditor 已不再积极维护。

如果你想使用原生 VS Code Markdown 编辑器，请在 `settings.json` 中加入以下配置：

```json
{
    "workbench.editorAssociations": {
        "*.md": "default",
        "*.markdown": "default"
    }
}
```

在编辑器中右键即可将 Markdown 导出为 PDF、DOCX 或 HTML。导出 PDF 依赖 Chromium，可通过 `vscode-office.chromiumPath` 配置其可执行文件路径。

导出 HTML 时，所有本地图片都会自动转换为 Base64 并嵌入文件中，因此导出的文件是完全独立的，分享时无需再附带图片资源。

![Markdown 编辑器截图](images/screenshot.png)

快捷键基于 [Vditor 快捷键](shortcut.md)，并额外提供以下命令：

- 列表上移一行：`Ctrl+Alt+I` / `⌘ ^ I`
- 列表下移一行：`Ctrl+Alt+J` / `⌘ ^ J`
- 在 VS Code 中编辑：`Ctrl+Alt+E` / `⌘ ^ E`

提示：

- 可通过 Ctrl/Cmd + 鼠标滚轮调整编辑器大小
- 可通过 Ctrl/Meta + 点击或双击打开超链接

## HTML

HTML 编辑器支持实时预览。按 `Ctrl+Shift+V` 即可打开实时视图。

## 致谢

本项目的诞生离不开以下作者的工作：

- **[cweijan](https://github.com/cweijan)** — 原始项目 [vscode-office](https://github.com/cweijan/vscode-office) 的作者，本 Fork 基于其工作构建。同时维护了一个专为该扩展适配的 [Vditor 定制版本](https://github.com/vscode-ext-studio/vditor)。
- **[RJ.Wang](https://github.com/rjwang1982)** — [vscode-office-enhanced](https://github.com/rjwang1982/vscode-office) 的维护者，贡献了 Mermaid v11 支持、自包含 HTML 导出、安装包体积优化以及多项编辑器 UX 改进，这些工作本 Fork 均继承自他。
- **[Vanessa219（Liyuan Li）](https://github.com/Vanessa219)** — [Vditor](https://github.com/Vanessa219/vditor) 的作者，本扩展 Markdown 所见即所得编辑能力的核心引擎。该项目由 [B3log](https://b3log.org) 开源社区开发，采用 MIT 协议。

## Credits

- PDF rendering: [mozilla/pdf.js/](https://github.com/mozilla/pdf.js/)
- DOCX rendering: [VolodymyrBaydalka/docxjs](https://github.com/VolodymyrBaydalka/docxjs)
- XLSX rendering:
  - [SheetJS/sheetjs](https://github.com/SheetJS/sheetjs): XLSX parsing
  - [myliang/x-spreadsheet](https://github.com/myliang/x-spreadsheet): XLSX rendering
- HTTP: [Rest Client](https://github.com/Huachao/vscode-restclient)
- Markdown: [Vanessa219/vditor](https://github.com/Vanessa219/vditor)
- Mermaid diagrams: [mermaid-js/mermaid](https://github.com/mermaid-js/mermaid)