import * as vscode from 'vscode';
import * as XLSX from 'xlsx';
import * as iconv from 'iconv-lite';
import { VcsResolver } from './vcsResolver';
import { CliFallback } from './cliFallback';
import { BlameEntry } from './types';

type SheetMap = Map<string, string[][]>;

class LruCache<K, V> {
    private map = new Map<K, V>();
    constructor(private capacity: number) { }
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
        try { return buf.toString('utf-8'); }
        catch { return iconv.decode(buf, 'windows-1252'); }
    }

    private parseCsvLine(line: string): string[] {
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
