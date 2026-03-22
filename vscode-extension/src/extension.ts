import * as vscode from 'vscode';
import * as path from 'path';

/**
 * Outlook MCP VS Code Extension
 *
 * This extension provides MCP server definitions for the Outlook MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft Outlook
 * through native COM automation.
 *
 * The extension bundles self-contained executables for both the MCP server and CLI -
 * no .NET SDK or runtime installation required.
 *
 * Agent Skills are registered via the chatSkills contribution point in package.json.
 */

export async function activate(context: vscode.ExtensionContext) {
	console.log('Outlook MCP extension is now active');

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('outlook-mcp', {
			provideMcpServerDefinitions: async () => {
				// Return the MCP server definition for the Outlook migration server
				const extensionPath = context.extensionPath;
				const mcpServerPath = path.join(extensionPath, 'bin', 'PptMcp.McpServer.exe');

				return [
					new vscode.McpStdioServerDefinition(
						'outlook-mcp',
						mcpServerPath,
						[],
						{
							// Optional environment variables can be added here if needed
						}
					)
				];
			}
		})
	);

	// Show welcome message on first activation
	const hasShownWelcome = context.globalState.get<boolean>('outlookmcp.hasShownWelcome', false);
	if (!hasShownWelcome) {
		showWelcomeMessage();
		context.globalState.update('outlookmcp.hasShownWelcome', true);
	}
}

function showWelcomeMessage() {
	const message = 'Outlook MCP migration extension activated! The Outlook MCP server is now available for AI assistants.';
	const learnMore = 'Learn More';

	vscode.window.showInformationMessage(message, learnMore).then(selection => {
		if (selection === learnMore) {
			vscode.env.openExternal(vscode.Uri.parse('https://github.com/trsdn/mcp-server-outlook'));
		}
	});
}

export function deactivate() {
	console.log('Outlook MCP extension is now deactivated');
}
