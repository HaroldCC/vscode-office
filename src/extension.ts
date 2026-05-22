import * as vscode from 'vscode';
import { EncodingStatusBar } from './common/encodingStatusBar';
import { MarkdownEditorProvider } from './provider/markdownEditorProvider';
import { OfficeViewerProvider } from './provider/officeViewerProvider';
import { ExcelDiffProvider } from './provider/excelDiffProvider';
import { HtmlService } from './service/htmlService';
import { MarkdownService } from './service/markdownService';
import { Output } from './common/Output';
import { FileUtil } from './common/fileUtil';
import { ReactApp } from './common/reactApp';
const httpExt = require('./bundle/extension');

export function activate(context: vscode.ExtensionContext) {
	keepOriginDiff();
	activeHTTP(context)
	const viewOption = { webviewOptions: { retainContextWhenHidden: true, enableFindWidget: false } };
	const viewOptionWithFind = { webviewOptions: { retainContextWhenHidden: true, enableFindWidget: true } };
	FileUtil.init(context)
	ReactApp.init(context)
	const encodingStatusBar = new EncodingStatusBar(context.globalState);
	const markdownService = new MarkdownService(context);
	const viewerInstance = new OfficeViewerProvider(context, encodingStatusBar);
	const markdownEditorProvider = new MarkdownEditorProvider(context, encodingStatusBar)
	const excelDiffProvider = new ExcelDiffProvider(context);
	context.subscriptions.push(
		encodingStatusBar.registerCommand(),
		vscode.commands.registerCommand('office.quickOpen', () => vscode.commands.executeCommand('workbench.action.quickOpen')),
		vscode.commands.registerCommand('office.markdown.switch', (uri) => { markdownService.switchEditor(uri) }),
		vscode.commands.registerCommand('office.markdown.paste', () => { markdownService.loadClipboardImage() }),
		vscode.commands.registerCommand('office.html.preview', uri => HtmlService.previewHtml(uri, context)),
		vscode.commands.registerCommand('office.excel.diffWithVCS', (uri) => excelDiffProvider.diffWithVCS(uri)),
		vscode.commands.registerCommand('office.excel.diffWithFile', (uri) => excelDiffProvider.diffWithFile(uri)),
		vscode.commands.registerCommand('office.excel.diffWithRevision', (uri) => excelDiffProvider.diffWithRevision(uri)),
		vscode.window.registerCustomEditorProvider("cweijan.markdownViewer", markdownEditorProvider, viewOptionWithFind),
		vscode.window.registerCustomEditorProvider("cweijan.markdownViewer.optional", markdownEditorProvider, viewOptionWithFind),
		...viewerInstance.bindCustomEditors(viewOption)
	);
}

export function deactivate() { }

async function activeHTTP(context: vscode.ExtensionContext) {
	try {
		httpExt.activate(context)
	} catch (error) {
		Output.debug(error)
	}
}

/**
 * 保持 Git diff 等场景使用默认编辑器，避免被本扩展接管。
 */
function keepOriginDiff() {
	try {
		const config = vscode.workspace.getConfiguration("workbench");
		const configKey = 'editorAssociations'
		const editorAssociations = config.get(configKey)
		const key = '{git,gitlens,git-graph}:/**/*.{md,csv,svg}'
		if (editorAssociations[key]) {
			editorAssociations[key] = undefined
			config.update(configKey, editorAssociations, true)
		}
	} catch (error) {
		Output.debug('keepOriginDiff failed: ' + error)
	}
}