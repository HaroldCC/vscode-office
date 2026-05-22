import * as vscode from 'vscode';

const ENCODING_LIST = [
    { label: 'UTF-8', value: 'utf-8' },
    { label: 'GBK', value: 'gbk' },
    { label: 'GB2312', value: 'gb2312' },
    { label: 'GB18030', value: 'gb18030' },
    { label: 'Big5', value: 'big5' },
    { label: 'Shift_JIS', value: 'shift_jis' },
    { label: 'EUC-JP', value: 'euc-jp' },
    { label: 'EUC-KR', value: 'euc-kr' },
    { label: 'ISO-8859-1', value: 'iso-8859-1' },
    { label: 'Windows-1252', value: 'windows-1252' },
    { label: 'UTF-16 LE', value: 'utf-16le' },
    { label: 'UTF-16 BE', value: 'utf-16be' },
];

export class EncodingStatusBar {

    private statusBarItem: vscode.StatusBarItem;
    private encodingMap: Map<string, string> = new Map();
    private currentUri: string | undefined;
    private listeners: ((uri: string, encoding: string) => void)[] = [];

    constructor(private state: vscode.Memento) {
        this.statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 90);
        this.statusBarItem.command = 'office.selectEncoding';
    }

    registerCommand(): vscode.Disposable {
        return vscode.commands.registerCommand('office.selectEncoding', () => this.showEncodingPicker());
    }

    onEncodingChanged(callback: (uri: string, encoding: string) => void) {
        this.listeners.push(callback);
    }

    bind(uri: string) {
        this.currentUri = uri;
        const encoding = this.getEncoding(uri);
        this.statusBarItem.text = `$(file-code) ${this.getDisplayName(encoding)}`;
        this.statusBarItem.tooltip = 'Select File Encoding';
        this.statusBarItem.show();
    }

    hide() {
        this.currentUri = undefined;
        this.statusBarItem.hide();
    }

    getEncoding(uri: string): string {
        const saved = this.state.get<string>(`encoding_${uri}`);
        if (saved) return saved;
        return this.encodingMap.get(uri) || 'utf-8';
    }

    setEncoding(uri: string, encoding: string) {
        this.encodingMap.set(uri, encoding);
        this.state.update(`encoding_${uri}`, encoding);
        if (this.currentUri === uri) {
            this.statusBarItem.text = `$(file-code) ${this.getDisplayName(encoding)}`;
        }
    }

    setDetectedEncoding(uri: string, encoding: string) {
        this.encodingMap.set(uri, encoding);
        if (this.currentUri === uri) {
            this.statusBarItem.text = `$(file-code) ${this.getDisplayName(encoding)}`;
        }
    }

    isExplicitlySet(uri: string): boolean {
        return this.state.get<string>(`encoding_${uri}`) !== undefined;
    }

    private getDisplayName(encoding: string): string {
        const item = ENCODING_LIST.find(e => e.value === encoding);
        return item ? item.label : encoding.toUpperCase();
    }

    private async showEncodingPicker() {
        if (!this.currentUri) return;

        const currentEncoding = this.getEncoding(this.currentUri);
        const items: vscode.QuickPickItem[] = ENCODING_LIST.map(e => ({
            label: e.label,
            description: e.value === currentEncoding ? 'Current' : undefined,
        }));

        const selected = await vscode.window.showQuickPick(items, {
            placeHolder: 'Select File Encoding',
        });
        if (!selected) return;

        const encoding = ENCODING_LIST.find(e => e.label === selected.label)?.value;
        if (!encoding || encoding === currentEncoding) return;

        this.encodingMap.set(this.currentUri, encoding);
        this.state.update(`encoding_${this.currentUri}`, encoding);
        this.statusBarItem.text = `$(file-code) ${selected.label}`;

        for (const listener of this.listeners) {
            listener(this.currentUri, encoding);
        }
    }

    dispose() {
        this.statusBarItem.dispose();
    }
}
