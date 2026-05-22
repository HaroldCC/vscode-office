# Excel Enhancement Design: Encoding Detection, Cell Diff, Version Diff, Lazy Loading

## Context

vscode-office 已有完善的 Excel/CSV 查看、编辑和 diff 功能。本次增强四个方面：
1. CSV 自动编码检测 — 减少用户手动选择编码的频率
2. Cell 级 diff 高亮 — 精确到单元格的变更可视化
3. 任意版本 diff — 支持选择 git commit / svn revision 做对比
4. 大文件懒加载 — 超过 10MB 的 XLSX 分 sheet 按需加载

---

## 1. CSV 自动编码检测

### 目标
首次打开 CSV 文件时自动检测编码，无需用户手动选择。

### 实现

**库选型：** `jschardet`（纯 JS，浏览器兼容，~30KB gzipped）

**执行位置：** Webview 端（`excel_reader.ts`）

**逻辑流程：**
```
打开 CSV 文件:
  1. 检查 globalState 是否有该文件的持久化编码
     → 有: 使用持久化编码（用户之前手动选择过）
     → 无: 继续步骤 2
  2. 取 buffer 前 4096 字节
  3. jschardet.detect(sample) → { encoding, confidence }
  4. confidence >= 0.7: 使用检测到的编码
     confidence < 0.7: fallback 到 UTF-8
  5. 将检测结果通知 Extension 端更新 StatusBar 显示
```

**修改文件：**
- `src/react/view/excel/excel_reader.ts` — 新增 `detectEncoding(buffer)` 函数
- `src/react/view/excel/Excel.tsx` — 在 `renderExcel()` 中首次打开时调用检测
- `package.json` — 添加 `jschardet` 依赖

**接口设计：**
```typescript
// excel_reader.ts
export function detectEncoding(buffer: ArrayBuffer): { encoding: string; confidence: number } {
  const sample = new Uint8Array(buffer.slice(0, 4096));
  const result = jschardet.detect(Buffer.from(sample));
  return {
    encoding: result.confidence >= 0.7 ? normalizeEncoding(result.encoding) : 'utf-8',
    confidence: result.confidence
  };
}
```

**编码名称标准化：** jschardet 返回的编码名（如 "windows-1252"、"GB2312"）需映射到 TextDecoder 接受的标准名。新增一个映射表。

---

## 2. Cell 级 Diff 高亮

### 目标
对比视图中，修改行内只高亮变化的具体 cell，而非整行染色。

### 当前状态
`excel_diff.ts` 的 `diffSheet()` 函数已在 match 行中逐 cell 比较文本，但最终只给整行分配一个 `modified` 样式。x-spreadsheet 的数据结构支持 per-cell 的 `style` 索引。

### 实现

**样式定义扩展：**
```typescript
// 现有 styles 数组
const styles = [
  { bgcolor: '#d4edda' },  // 0: added row (green)
  { bgcolor: '#f8d7da' },  // 1: deleted row (red)
  { bgcolor: '#fff3cd' },  // 2: modified row (yellow) - 不再使用于整行
  { bgcolor: '#f5f5f5' },  // 3: empty placeholder (gray)
  { bgcolor: '#fff3cd' },  // 4: modified cell (yellow) - 新增，用于单个 cell
  { bgcolor: '#ffffff' },  // 5: unchanged cell in modified row (white)
];
```

**逻辑变更（`excel_diff.ts`）：**
```
对于 match 行:
  逐 cell 比较:
    → cell 内容相同: style = 5 (white/unchanged)
    → cell 内容不同: style = 4 (yellow/modified)
  如果该行有任何 cell 不同: 标记为 modified 行（用于统计和导航）
  如果所有 cell 相同: 标记为 unchanged 行
```

**修改文件：**
- `src/react/view/excel/excel_diff.ts` — 修改 `diffSheet()` 中 match 行的样式分配逻辑
- `src/react/view/excel/ExcelDiff.less` — 无需改动（样式通过 x-spreadsheet 内联）

**向后兼容：** 新增/删除行保持整行着色不变。仅 match 行的表现变化。

---

## 3. 任意版本 Diff

### 目标
支持用户从最近的 git commit / svn revision 列表中选择一个版本做对比。

### 实现

**新增命令：** `office.excel.diffWithRevision`

**Extension 端流程：**
```
1. 判断 VCS 类型（先尝试 git，失败则 svn）
2. 获取版本列表:
   Git: git log --oneline --follow -20 -- <file>
   SVN: svn log -l 20 --xml <file>  (解析 XML 获取 revision + message + date)
3. 展示 QuickPick:
   Git: "abc1234 - Fix data format (2 days ago)"
   SVN: "r1234 - Update report (2026-05-20)"
4. 用户选择后获取该版本文件:
   Git: git show <hash>:<relative-path>
   SVN: svn cat -r <revision> <file>
5. 调用现有 openDiffPanel(currentUri, baseBuffer, title)
```

**QuickPick Item 格式：**
```typescript
interface RevisionItem extends vscode.QuickPickItem {
  label: string;       // "abc1234" 或 "r1234"
  description: string; // commit message (截断到 60 字符)
  detail: string;      // 相对时间 "2 days ago" 或日期
  hash: string;        // git hash 或 svn revision number
}
```

**修改文件：**
- `src/provider/excelDiffProvider.ts` — 新增 `diffWithRevision(uri)` 方法，包含 `getGitLog()` 和 `getSvnLog()` 辅助函数
- `src/extension.ts` — 注册新命令
- `package.json` — 添加 command 定义和 context menu 贡献

**Context Menu：** 在 Excel 文件的编辑器标题栏右键菜单中添加 "Compare with Revision..."。

---

## 4. 大文件懒加载

### 目标
超过 10MB 的 XLSX 文件只先加载第一个 sheet，其他 sheet 按用户点击按需加载。

### 实现

**阈值：** 10MB（`10 * 1024 * 1024` bytes）

**XLSX 库能力：** `XLSX.read()` 支持 `sheets` 选项指定只解析哪些 sheet（按名称或索引）。可以先用 `{ bookSheets: true }` 只读 sheet 名称列表。

**新增 API（`excel_reader.ts`）：**
```typescript
// 只读取 sheet 名称列表（不解析数据）
export function loadSheetNames(buffer: ArrayBuffer, ext: string): string[] {
  const wb = XLSX.read(buffer, { bookSheets: true });
  return wb.SheetNames;
}

// 按需加载单个 sheet
export function loadSingleSheet(
  buffer: ArrayBuffer, ext: string, sheetName: string, encoding?: string
): SheetInfo {
  const wb = XLSX.read(buffer, { sheets: [sheetName], cellFormula: true, cellNF: true });
  // 解析该 sheet 并返回 SheetInfo
}
```

**Excel.tsx 逻辑变更：**
```
renderExcel(path, ext, encoding):
  1. fetch buffer
  2. if buffer.byteLength > 10MB:
       → loadSheetNames(buffer) → 获得 sheet 名称列表
       → loadSingleSheet(buffer, ext, sheetNames[0], encoding) → 解析第一个 sheet
       → UI: sheet tabs 全部显示，未加载的 tab 带 "..." 标记
       → 缓存 buffer 引用（用于后续按需加载）
     else:
       → loadSheets(buffer, ext, encoding) → 全量加载（现有逻辑不变）
  3. 用户切换 sheet tab:
     → if sheet 未加载: loadSingleSheet() → 更新数据 → 渲染
     → if sheet 已加载: 直接切换（现有逻辑）
```

**状态管理：**
```typescript
// Excel.tsx 新增状态
const bufferRef = useRef<ArrayBuffer | null>(null);     // 缓存原始 buffer
const loadedSheetsRef = useRef<Set<string>>(new Set()); // 已加载的 sheet 名称
const isLazyMode = useRef(false);                       // 是否懒加载模式
```

**修改文件：**
- `src/react/view/excel/excel_reader.ts` — 新增 `loadSheetNames()` 和 `loadSingleSheet()`
- `src/react/view/excel/Excel.tsx` — 添加懒加载逻辑和 sheet 切换按需加载

**UX 细节：**
- 未加载的 sheet tab 显示名称但带加载指示器
- 切换到未加载 sheet 时显示 loading spinner
- 加载完成后 spinner 消失，自动渲染

---

## 5. 验证方案

| 功能 | 测试方法 |
|------|----------|
| 编码检测 | 打开一个 GBK 编码的 CSV，确认自动检测为 GBK 而非乱码 |
| Cell 级 diff | 修改 XLSX 中几个 cell，触发 diff，确认只有修改的 cell 高亮黄色 |
| 版本 diff | 对已提交的 XLSX 做修改，执行 "Compare with Revision"，确认列表正常、对比正确 |
| 懒加载 | 打开一个 >10MB 的多 sheet XLSX，确认首次只加载第一个 sheet，切换 tab 按需加载 |

---

## 6. 依赖变更

| 包名 | 版本 | 用途 |
|------|------|------|
| `jschardet` | `^3.1.x` | CSV 编码自动检测 |

其他功能无需新增依赖，均基于现有的 xlsx 库和 VSCode API。
