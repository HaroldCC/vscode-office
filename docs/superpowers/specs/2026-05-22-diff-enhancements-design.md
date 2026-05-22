# Diff Tool Enhancements Design (Subsystem 4 of 5)

> 仅覆盖 **子系统 4 — diff 工具的 UX/性能强化**。

## 1. Context

子系统 1 给 diff 添加了 ref switcher 与 blame。本子系统聚焦 UX 缺陷与性能。

## 2. Goals

1. **Sticky 标题行**:diff 表格滚动时表头行(第 0 行)保持可见
2. **Inline diff(cell 内文本级)**:修改 cell 上 hover 显示具体字符级差异(类似 textual diff 红绿染色)
3. **过滤模式**:toolbar 增加 "Only show changes" 切换,隐藏 unchanged 行
4. **跨 sheet 变更汇总**:面板顶部显示所有 sheet 累计的 +/-/~ 统计
5. **导出 diff**:右上角增加 "Export diff…" → 导出 HTML 报告(单文件,可邮件分享)
6. **大文件性能**:>5MB 时 diff 计算放进 Web Worker 避免 UI 卡顿

## 3. Non-Goals

- 替换 x-spreadsheet(改库成本远超价值)
- 排序/筛选(交给编辑器子系统)

## 4. 实现要点

### 4.1 Sticky Header

x-spreadsheet 已有 `freeze` 选项。在 `renderSide` 调用时设置 `freeze: [1, 0]`(冻结第一行)。如果数据有多个 sheet,逐 sheet 设置。

### 4.2 Inline diff

在 `excel_diff.ts` 的 cell-level diff 中,当 base ≠ current 且二者都是字符串时,运行简单 LCS 字符级 diff,产出 segments `[{type:'eq'|'add'|'del', text:string}, ...]`,存在 cell 的 `comment` 字段中。webview 渲染 cell hover 时显示一个 popover,segments 按颜色染色。

### 4.3 Only Show Changes

新增 toolbar 切换。开启时:`excel_diff.ts` 输出的 `leftSheets[i].rows / rightSheets[i].rows` 过滤,只保留 `changeRows` 中的索引。视觉上行号显示原始 row(用 fixed col header 或子标题)。

### 4.4 跨 sheet 统计

`computeDiff()` 已返回每个 sheet 的 stats。聚合所有 sheet 的 added/deleted/modified/unchanged 总和。Toolbar 显示新的"Total" 与"Per sheet"切换。

### 4.5 导出 HTML

新增 `excel_diff_html.ts`:遍历 leftSheets/rightSheets,生成自包含 HTML:

```html
<table class="diff">
  <tr class="row-added"><td class="left empty"></td><td class="right">new cell</td></tr>
  <tr class="row-modified"><td class="left">old</td><td class="right">new</td></tr>
  ...
</table>
<style>...</style>
```

通过 `vscode.workspace.fs.writeFile` 写到用户选择的路径。

### 4.6 Web Worker

`excel_diff_worker.ts`:Vite 支持 `new Worker(new URL('./diff.worker.ts', import.meta.url))`。Webview 在 buffer > 5MB 时把 buffer + sheetData 传给 worker,worker 跑 computeDiff,返回结果。

## 5. 受影响文件

新增:
- `src/react/view/excel/excel_diff_html.ts`
- `src/react/view/excel/diff.worker.ts`

修改:
- `src/react/view/excel/excel_diff.ts`(inline diff segments)
- `src/react/view/excel/ExcelDiff.tsx`(toolbar 增加 "Only changes"/"Export" 按钮 + sticky freeze)
- `src/react/view/excel/ExcelDiff.less`(过滤模式/导出按钮样式)

## 6. 验证

- Sticky:打开大表,垂直滚动,确认 header 不动
- Inline diff:hover 改动 cell,popover 显示字符级染色
- Only changes:切换后只显示变更行,变更率高时显著加速渲染
- Export:导出 HTML 文件,用浏览器打开能正常查看
- Worker:>5MB 文件加载时主线程不阻塞(spinner 仍能动)
