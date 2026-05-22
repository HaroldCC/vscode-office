import { spawn } from 'child_process';
import * as vscode from 'vscode';
import { TortoiseLocator, TortoiseKind } from './tortoiseLocator';

export type TortoiseCommand =
    | 'diff' | 'log' | 'blame' | 'commit' | 'revert' | 'update' | 'showcompare';

export class TortoiseProc {
    static async run(kind: TortoiseKind, command: TortoiseCommand, filePath: string): Promise<void> {
        const exe = await TortoiseLocator.locate(kind);
        if (!exe) {
            const name = kind === 'svn' ? 'TortoiseSVN' : 'TortoiseGit';
            const action = await vscode.window.showErrorMessage(
                `${name} not found. Configure path in settings?`,
                'Open Settings'
            );
            if (action === 'Open Settings') {
                vscode.commands.executeCommand('workbench.action.openSettings', `vscode-office.tortoise`);
            }
            return;
        }

        const args = [`/command:${command}`, `/path:${filePath}`];
        try {
            const child = spawn(exe, args, { detached: true, stdio: 'ignore' });
            child.unref();
        } catch (err: any) {
            vscode.window.showErrorMessage(`Failed to launch Tortoise: ${err.message}`);
        }
    }
}
