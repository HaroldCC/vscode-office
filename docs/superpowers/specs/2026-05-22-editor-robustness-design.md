# CSV/XLSX Editor Robustness Design (Subsystem 5 of 5)

> 仅覆盖 **子系统 5 — 编辑器自身的健壮性与体验**。

## 1. Context

现有 Excel.tsx 实现了 viewer + 编辑 + lazy load + 搜索。已知不足:
1. CSV 中含 git 冲突标记(`<<<<<<<`/`=======`/`>>>>>>>`)时直接打开会破坏表格结构
2. 编辑后切换 sheet tab 不提示未保存
3. 行高列宽用户手动调整后 reload 丢失(无持久化)
4. 没有 undo/redo 状态指示(x-spreadsheet 内置 undo/redo 但 UI 不显式)
5. 保存失败(磁盘满/只读)时,编辑器内仍显示 saved 状态

## 2. Goals

1. **冲突标记检测**:打开 csv 时若发现 `<<<<<<<` 标记,弹出"打开 3-way merge 编辑器?"
2. **未保存提示**:dirty 状态强化 — title bar 显示 `●`,切换 sheet 时 dirty 状态不丢失
3. **保存失败回退**:保存失败时不清 dirty 标志,显示错误
4. **行高/列宽持久化**:每个文件的尺寸调整存到 workspaceState(per file path)
5. **CSV 分隔符自动检测**:已检测编码,扩展为也检测分隔符(`,` `;` `\t`)
6. **大文件提示**:>10MB 文件打开时 toast 说明已启用 lazy mode

## 3. 实现要点

### 3.1 冲突标记检测(Excel.tsx)

```typescript
// 在 renderExcel 中,csv 内容首字符为 `<<<<<<<` 或包含 `\n<<<<<<< ` 时,
const text = new TextDecoder(encoding).decode(buffer);
if (/^|\n<{7} /m.test(text) && /\n={7}\n/.test(text) && /\n>{7} /m.test(text)) {
    const choice = await new Promise<string>(resolve => {
        // antd Modal.confirm
    });
    if (choice === 'open-merge') {
        handler.emit('openMergeEditor');
        return;
    }
    // 否则继续解析,但 warn 用户
}
```

Extension 端新增消息 `openMergeEditor` → 调起 `office.merge.openConflictEditor` 命令。

### 3.2 未保存提示

x-spreadsheet 已派发 `change` 事件。捕获:`spreadsheetRef.current.on('cell-edited', () => setDirty(true))`。

dirty 状态在 sheet 切换时**不重置**,仅在 save 成功后重置。

```typescript
const [dirty, setDirty] = useState(false);
...
useEffect(() => {
    handler.emit('setTitle', { dirty });
}, [dirty]);
```

Extension 端监听 `setTitle`,设置 panel title 前缀 `● `。

### 3.3 保存失败回退

修改 `Excel.tsx` 的 saveDone handler:

```typescript
.on('saveDone', () => { setDirty(false); message.success(...); })
.on('saveError', (err) => { message.error(`Save failed: ${err}`); /* dirty 保持 true */ })
```

`OfficeViewerProvider` 的 save handler 在 catch 中改为发 `saveError`:

```typescript
} catch (err) {
    panel.webview.postMessage({ type: 'saveError', error: err.message });
}
```

### 3.4 行高/列宽持久化

x-spreadsheet 没有"调整尺寸"事件 hook。但 sheet data 序列化时 `rows[r].height` / `cols[c].width` 保留。每次 save 时把这些抽出存 workspaceState:

```typescript
ctx.workspaceState.update(`office.size.${uri.toString()}`, {
    rows: ss.data.rows._ /* {0: {height: 30}, ...} */,
    cols: ss.data.cols._
});
```

重新打开时,在 loadSheets 后注入。

由于现有保存路径已经写入文件本身(xlsx 保留 cols/rows),这一项仅对 csv 有意义(csv 没有元数据)。**Scope reduce**:仅 csv 走 workspaceState。

### 3.5 CSV 分隔符检测

`excel_reader.ts` 的 `loadSheets` 在 CSV 分支加上:

```typescript
const delimiter = detectDelimiter(text); // 在 first 50 行 take majority of (',', ';', '\t')
wb = XLSX.read(text, { type: "string", raw: true, FS: delimiter });
```

实现:统计前 50 行各分隔符出现次数,取最多。

### 3.6 大文件 toast

`Excel.tsx` 中 `shouldLazyLoad` 已经存在;在进入 lazy mode 分支时:

```typescript
message.info({ duration: 3, content: `Lazy mode enabled (${(size/1024/1024).toFixed(1)}MB). Sheets load on demand.` });
```

## 4. 受影响文件

修改:
- `src/react/view/excel/Excel.tsx`(冲突检测 + dirty + saveError + 大文件 toast)
- `src/react/view/excel/excel_reader.ts`(分隔符检测函数)
- `src/provider/officeViewerProvider.ts`(saveError 消息 + setTitle)
- `src/extension.ts`(命令 office.merge.openConflictEditor 已存在,无需改)

## 5. 验证

- 制造 csv 冲突 → 打开 → 弹出提示
- 修改后切换 sheet → 还是 dirty
- 保存到只读文件 → 错误 toast,dirty 不变
- 打开 `;` 分隔的 csv → 正确解析为列
- 打开 >10MB xlsx → 进入 lazy mode,toast 提示

## 6. 风险

- CSV 内容里恰好有 `<<<<<<<` 的非冲突文本会误报 → 三段标记都匹配才提示,误报率低
- 分隔符检测在很短 csv 上不稳 → 仅当前 50 行内某分隔符占比 > 1.5 倍其他时才切换,否则 fallback `,`
