import * as vscode from 'vscode';
import { CliFallback } from '../vcs/cliFallback';

export interface MergeBuffers {
    base: Buffer;     // 可能为空(Add/Add 冲突)
    ours: Buffer;
    theirs: Buffer;
}

export class MergeLoader {
    static async load(uri: vscode.Uri): Promise<MergeBuffers> {
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
