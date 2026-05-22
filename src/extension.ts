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
import { GitApiHelper } from './provider/vcs/gitApi';
import { OfficeScmContribution } from './provider/vcs/scmContribution';
import { TortoiseProc } from './provider/tortoise/tortoiseProc';
import { TortoiseLocator } from './provider/tortoise/tortoiseLocator';
import { ExternalDiffUriHandler, configureExternalDiff } from './provider/tortoise/externalDiffCommand';
import { MergeEditorProvider } from './provider/merge/mergeEditorProvider';
import { MergeDetector } from './provider/merge/mergeDetector';
const httpExt = require('./bundle/extension');

export async function activate(context: vscode.ExtensionContext) {
	keepOriginDiff();
	activeHTTP(context);
	const viewOption = { webviewOptions: { retainContextWhenHidden: true, enableFindWidget: false } };
	const viewOptionWithFind = { webviewOptions: { retainContextWhenHidden: true, enableFindWidget: true } };
	FileUtil.init(context);
	ReactApp.init(context);
	const encodingStatusBar = new EncodingStatusBar(context.globalState);
	const markdownService = new MarkdownService(context);
	const viewerInstance = new OfficeViewerProvider(context, encodingStatusBar);
	const markdownEditorProvider = new MarkdownEditorProvider(context, encodingStatusBar);
	const excelDiffProvider = new ExcelDiffProvider(context);
	const mergeEditorProvider = new MergeEditorProvider(context);

	GitApiHelper.instance.ensureInit().catch((err) => Output.debug('git api init: ' + err));
	const scmContribution = new OfficeScmContribution();
	scmContribution.activate(context).catch((err) => Output.debug('scm init: ' + err));
	const mergeDetector = new MergeDetector();
	mergeDetector.activate(context).catch((err) => Output.debug('merge detector init: ' + err));

	vscode.commands.executeCommand('setContext', 'office.isWindows', process.platform === 'win32');

	context.subscriptions.push(
		vscode.window.registerUriHandler(new ExternalDiffUriHandler(excelDiffProvider)),
		encodingStatusBar.registerCommand(),
		vscode.commands.registerCommand('office.quickOpen', () => vscode.commands.executeCommand('workbench.action.quickOpen')),
		vscode.commands.registerCommand('office.markdown.switch', (uri) => { markdownService.switchEditor(uri) }),
		vscode.commands.registerCommand('office.markdown.paste', () => { markdownService.loadClipboardImage() }),
		vscode.commands.registerCommand('office.html.preview', uri => HtmlService.previewHtml(uri, context)),
		vscode.commands.registerCommand('office.excel.diffWithVCS', (uri) => excelDiffProvider.diffWithVCS(uri)),
		vscode.commands.registerCommand('office.excel.diffWithFile', (uri) => excelDiffProvider.diffWithFile(uri)),
		vscode.commands.registerCommand('office.excel.diffWithRevision', (uri) => excelDiffProvider.diffWithRevision(uri)),
		vscode.commands.registerCommand('office.merge.openConflictEditor', (uri) => mergeEditorProvider.openConflictEditor(uri)),
		vscode.commands.registerCommand('office.tortoise.svnLog', (uri?: vscode.Uri) => runTortoise('svn', 'log', uri)),
		vscode.commands.registerCommand('office.tortoise.gitLog', (uri?: vscode.Uri) => runTortoise('git', 'log', uri)),
		vscode.commands.registerCommand('office.tortoise.svnDiff', (uri?: vscode.Uri) => runTortoise('svn', 'diff', uri)),
		vscode.commands.registerCommand('office.tortoise.gitDiff', (uri?: vscode.Uri) => runTortoise('git', 'diff', uri)),
		vscode.commands.registerCommand('office.tortoise.svnBlame', (uri?: vscode.Uri) => runTortoise('svn', 'blame', uri)),
		vscode.commands.registerCommand('office.tortoise.gitBlame', (uri?: vscode.Uri) => runTortoise('git', 'blame', uri)),
		vscode.commands.registerCommand('office.tortoise.configureExternal', configureExternalDiff),
		vscode.commands.registerCommand('office.tortoise.refreshLocator', () => {
			TortoiseLocator.clearCache();
			vscode.window.showInformationMessage('Tortoise locator cache cleared.');
		}),
		vscode.window.registerCustomEditorProvider("cweijan.markdownViewer", markdownEditorProvider, viewOptionWithFind),
		vscode.window.registerCustomEditorProvider("cweijan.markdownViewer.optional", markdownEditorProvider, viewOptionWithFind),
		...viewerInstance.bindCustomEditors(viewOption)
	);
}

function runTortoise(kind: 'svn' | 'git', cmd: 'log' | 'diff' | 'blame', uri?: vscode.Uri) {
	const target = uri || vscode.window.activeTextEditor?.document.uri;
	if (!target) {
		vscode.window.showErrorMessage('No file selected for Tortoise command.');
		return;
	}
	TortoiseProc.run(kind, cmd, target.fsPath);
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
