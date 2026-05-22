import { EncodingStatusBar } from '@/common/encodingStatusBar';
import { ReactApp } from '@/common/reactApp';
import { readFileSync } from 'fs';
import { extname } from 'path';
import * as vscode from 'vscode';
import { Handler } from '../common/handler';
import { Util } from '../common/util';
import { handleImage, isImage } from './handlers/imageHandler';
import { handleZip } from './compress/zipHandler';
import { handleRar } from './compress/rarHandler';
import { handleCommonEvent } from './compress/commonHandler';
import iconv from 'iconv-lite';

/**
 * support view office files
 */
export class OfficeViewerProvider implements vscode.CustomReadonlyEditorProvider {

    private extensionPath: string;
    private activeHandlers: Map<string, Handler> = new Map();

    constructor(private context: vscode.ExtensionContext, private encodingStatusBar: EncodingStatusBar) {
        this.extensionPath = context.extensionPath;
        this.encodingStatusBar.onEncodingChanged((uri, encoding) => {
            const handler = this.activeHandlers.get(uri);
            if (handler) {
                handler.emit('changeEncoding', encoding);
            }
        });
    }

    bindCustomEditors(viewOption: { webviewOptions: vscode.WebviewPanelOptions }) {
        const viewers = ['cweijan.officeViewer', 'cweijan.imageViewer', 'cweijan.htmlViewer']
        return viewers.map(viewer => vscode.window.registerCustomEditorProvider(viewer, this, viewOption))
    }

    public openCustomDocument(uri: vscode.Uri, openContext: vscode.CustomDocumentOpenContext, token: vscode.CancellationToken): vscode.CustomDocument | Thenable<vscode.CustomDocument> {
        return { uri, dispose: (): void => { } };
    }
    public resolveCustomEditor(document: vscode.CustomDocument, webviewPanel: vscode.WebviewPanel, token: vscode.CancellationToken): void | Thenable<void> {
        const uri = document.uri;
        const uriStr = uri.toString();
        const webview = webviewPanel.webview;
        const folderPath = vscode.Uri.joinPath(uri, '..')
        webview.options = {
            enableScripts: true,
            localResourceRoots: [vscode.Uri.file(this.extensionPath), folderPath]
        }

        const handler = Handler.bind(webviewPanel, uri)
        this.activeHandlers.set(uriStr, handler);
        handleCommonEvent(uri, handler, this.encodingStatusBar)

        webviewPanel.onDidChangeViewState(e => {
            if (e.webviewPanel.active) {
                this.encodingStatusBar.bind(uriStr);
            }
        });
        webviewPanel.onDidDispose(() => {
            this.activeHandlers.delete(uriStr);
            this.encodingStatusBar.hide();
        });
        this.encodingStatusBar.bind(uriStr);

        let route: string;
        const ext = extname(uri.fsPath).toLowerCase()
        if (isImage(ext)) {
            handleImage(handler, uri, webview)
            route = 'image'
        }
        switch (ext) {
            case ".xlsx":
            case ".xlsm":
            case ".xls":
            case ".csv":
            case ".ods":
                route = 'excel';
                break;
            case ".docx":
            case ".dotx":
                route = 'word'
                break;
            case ".jar":
            case ".zip":
            case ".apk":
            case ".vsix":
                route = 'zip';
                handleZip(uri, handler);
                break;
            case ".rar":
                route = 'zip';
                handleRar(uri, handler);
                break;
            case ".ttf":
            case ".woff":
            case ".woff2":
            case ".otf":
                route = 'font';
                break;
            case ".pdf":
                webview.html = readFileSync(this.extensionPath + "/resource/pdf/viewer.html", 'utf8')
                    .replace("{{baseUrl}}", this.getBaseUrl(webview, 'pdf'))
                break;
            case ".htm":
            case ".html":
                const readHtml = () => {
                    const encoding = this.encodingStatusBar.getEncoding(uriStr);
                    const buf = readFileSync(uri.fsPath);
                    const text = encoding === 'utf-8' ? buf.toString('utf8') : iconv.decode(buf, encoding);
                    return text;
                };
                webview.html = Util.buildPath(readHtml(), webview, folderPath.fsPath);
                Util.listen(webviewPanel, uri, () => {
                    webviewPanel.webview.html = Util.buildPath(readHtml(), webviewPanel.webview, folderPath.fsPath);
                })
                handler.on('changeEncoding', () => {
                    webviewPanel.webview.html = Util.buildPath(readHtml(), webviewPanel.webview, folderPath.fsPath);
                })
                break;
            default:
                if (route) break;
                vscode.commands.executeCommand('vscode.openWith', uri, "default");
        }
        if (route) return ReactApp.view(webview, { route })
    }

    private getBaseUrl(webview: vscode.Webview, path: string) {
        const baseUrl = webview.asWebviewUri(vscode.Uri.file(`${this.extensionPath}/resource/${path}`))
            .toString().replace(/\?.+$/, '').replace('https://git', 'https://file')
        return baseUrl;
    }

}