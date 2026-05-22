import * as vscode from 'vscode';
import { ExcelDiffProvider } from '../excelDiffProvider';

/**
 * URI handler: vscode://rjwang.vscode-office-enhanced/diff
 *   ?base=<absolute path>&mine=<absolute path>&baseName=<label>&mineName=<label>
 *
 * 用于被 TortoiseSVN/TortoiseGit 作为外部 diff 工具调起。
 */
export class ExternalDiffUriHandler implements vscode.UriHandler {

    constructor(private excelDiffProvider: ExcelDiffProvider) { }

    handleUri(uri: vscode.Uri): void {
        if (uri.path !== '/diff' && uri.path !== 'diff') return;
        const params = new URLSearchParams(uri.query);
        const base = params.get('base');
        const mine = params.get('mine');
        if (!base || !mine) {
            vscode.window.showErrorMessage('Office Diff URI missing base or mine parameter');
            return;
        }
        const baseLabel = params.get('baseName') || 'BASE';
        const mineLabel = params.get('mineName') || 'MINE';
        const baseUri = vscode.Uri.file(base);
        const mineUri = vscode.Uri.file(mine);
        this.excelDiffProvider.diffWithExternalFiles(baseUri, mineUri, baseLabel, mineLabel);
    }
}

/**
 * 命令实现:`office.tortoise.configureExternal`
 * 把可粘贴到 Tortoise 设置的外部 diff 命令写入剪贴板,并显示说明。
 */
export async function configureExternalDiff(): Promise<void> {
    const cmd = 'cmd.exe /c "start vscode://rjwang.vscode-office-enhanced/diff?base=%base&mine=%mine&baseName=%bname&mineName=%yname"';
    await vscode.env.clipboard.writeText(cmd);
    const choice = await vscode.window.showInformationMessage(
        'Tortoise external diff command copied to clipboard. Paste it into TortoiseSVN/TortoiseGit Settings → Diff Viewer → External, for *.xlsx, *.xlsm, *.xls, *.csv, *.ods file types.',
        'Open TortoiseSVN Settings',
        'Open TortoiseGit Settings',
    );
    if (choice === 'Open TortoiseSVN Settings') {
        vscode.commands.executeCommand('office.tortoise.openTortoiseSvnSettings');
    } else if (choice === 'Open TortoiseGit Settings') {
        vscode.commands.executeCommand('office.tortoise.openTortoiseGitSettings');
    }
}
