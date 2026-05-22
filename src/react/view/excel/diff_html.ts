import { DiffSheetData } from './excel_diff';

/**
 * 生成自包含 HTML 报告。leftSheets / rightSheets 必须长度一致且对齐。
 */
export function exportDiffHtml(
    leftSheets: DiffSheetData[],
    rightSheets: DiffSheetData[],
    stats: { added: number; deleted: number; modified: number; unchanged: number },
    title: string = 'Diff Report',
): string {
    const styles = `
        body { font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Helvetica, Arial, sans-serif; margin: 16px; color: #222; }
        h1 { font-size: 18px; margin: 0 0 8px; }
        .stats { font-family: monospace; margin-bottom: 12px; }
        .stat-added { color: #28a745; margin-right: 12px; }
        .stat-deleted { color: #dc3545; margin-right: 12px; }
        .stat-modified { color: #d4a017; margin-right: 12px; }
        .sheet-title { font-weight: 600; margin: 16px 0 6px; font-size: 13px; }
        table.diff { border-collapse: collapse; font-size: 12px; width: 100%; table-layout: fixed; }
        table.diff th, table.diff td { border: 1px solid #ddd; padding: 4px 6px; vertical-align: top; word-break: break-word; }
        table.diff th { background: #f5f5f5; font-weight: 600; text-align: left; }
        table.diff td.idx { background: #fafafa; color: #999; text-align: right; width: 40px; }
        table.diff td.added { background: #d4edda; }
        table.diff td.deleted { background: #f8d7da; }
        table.diff td.modified { background: #fff3cd; }
    `;

    const renderSheetPair = (left: DiffSheetData, right: DiffSheetData): string => {
        const rows: string[] = [];
        const maxRows = Math.max(left.maxRows, right.maxRows);
        const maxCols = Math.max(left.maxCols, right.maxCols);
        for (let r = 0; r < maxRows; r++) {
            const lRow = left.rows[r]?.cells || {};
            const rRow = right.rows[r]?.cells || {};
            const lCells: string[] = [];
            const rCells: string[] = [];
            for (let c = 0; c < maxCols; c++) {
                const lv = lRow[c]?.text ?? '';
                const rv = rRow[c]?.text ?? '';
                const lStyle = lRow[c]?.style;
                const rStyle = rRow[c]?.style;
                const lCls = lStyle != null ? cellStyleClass(left.styles[lStyle]) : '';
                const rCls = rStyle != null ? cellStyleClass(right.styles[rStyle]) : '';
                lCells.push(`<td class="${lCls}">${escapeHtml(lv)}</td>`);
                rCells.push(`<td class="${rCls}">${escapeHtml(rv)}</td>`);
            }
            rows.push(`<tr><td class="idx">${r + 1}</td>${lCells.join('')}<td class="idx">|</td>${rCells.join('')}</tr>`);
        }
        return `<table class="diff"><tbody>${rows.join('')}</tbody></table>`;
    };

    const sheetSections = leftSheets.map((l, i) => {
        const r = rightSheets[i];
        if (!r) return '';
        return `<div class="sheet-title">Sheet: ${escapeHtml(l.name)}</div>${renderSheetPair(l, r)}`;
    }).join('');

    return `<!doctype html>
<html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>${styles}</style></head>
<body>
<h1>${escapeHtml(title)}</h1>
<div class="stats">
  <span class="stat-added">+${stats.added}</span>
  <span class="stat-deleted">-${stats.deleted}</span>
  <span class="stat-modified">~${stats.modified}</span>
  <span>${stats.unchanged} unchanged</span>
</div>
${sheetSections}
</body></html>`;
}

function escapeHtml(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function cellStyleClass(style: any): string {
    if (!style?.bgcolor) return '';
    const bg = style.bgcolor.toLowerCase();
    if (bg === '#d4edda') return 'added';
    if (bg === '#f8d7da') return 'deleted';
    if (bg === '#fff3cd') return 'modified';
    return '';
}
