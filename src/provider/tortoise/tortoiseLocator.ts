import * as fs from 'fs';
import * as path from 'path';
import { execFile } from 'child_process';
import * as vscode from 'vscode';

export type TortoiseKind = 'svn' | 'git';

const DEFAULT_PATHS: Record<TortoiseKind, string[]> = {
    svn: [
        'C:\\Program Files\\TortoiseSVN\\bin\\TortoiseProc.exe',
        'C:\\Program Files (x86)\\TortoiseSVN\\bin\\TortoiseProc.exe',
    ],
    git: [
        'C:\\Program Files\\TortoiseGit\\bin\\TortoiseGitProc.exe',
        'C:\\Program Files (x86)\\TortoiseGit\\bin\\TortoiseGitProc.exe',
    ],
};

const REG_KEYS: Record<TortoiseKind, string> = {
    svn: 'HKLM\\SOFTWARE\\TortoiseSVN',
    git: 'HKLM\\SOFTWARE\\TortoiseGit',
};

function regQuery(key: string, value: string): Promise<string | null> {
    return new Promise(resolve => {
        execFile('reg', ['query', key, '/v', value], { encoding: 'utf8' }, (err, stdout) => {
            if (err) return resolve(null);
            const m = stdout.match(/REG_\w+\s+(.+)/);
            resolve(m ? m[1].trim() : null);
        });
    });
}

export class TortoiseLocator {
    private static cache = new Map<TortoiseKind, string | null>();

    static async locate(kind: TortoiseKind): Promise<string | null> {
        if (process.platform !== 'win32') return null;
        if (this.cache.has(kind)) return this.cache.get(kind)!;

        // 1. user config
        const cfgKey = kind === 'svn' ? 'tortoise.svnPath' : 'tortoise.gitPath';
        const configured = vscode.workspace.getConfiguration('vscode-office').get<string>(cfgKey);
        if (configured && fs.existsSync(configured)) {
            this.cache.set(kind, configured);
            return configured;
        }

        // 2. registry: TortoiseSVN uses "ProcPath", TortoiseGit uses "ProcPath" too
        const fromReg = await regQuery(REG_KEYS[kind], 'ProcPath');
        if (fromReg && fs.existsSync(fromReg)) {
            this.cache.set(kind, fromReg);
            return fromReg;
        }

        // 3. installation directory key
        const dirReg = await regQuery(REG_KEYS[kind], 'Directory');
        if (dirReg) {
            const exe = kind === 'svn' ? 'TortoiseProc.exe' : 'TortoiseGitProc.exe';
            const candidate = path.join(dirReg, 'bin', exe);
            if (fs.existsSync(candidate)) {
                this.cache.set(kind, candidate);
                return candidate;
            }
        }

        // 4. defaults
        for (const p of DEFAULT_PATHS[kind]) {
            if (fs.existsSync(p)) {
                this.cache.set(kind, p);
                return p;
            }
        }

        this.cache.set(kind, null);
        return null;
    }

    static clearCache(): void {
        this.cache.clear();
    }
}
