import * as fs from "fs";
import * as path from "path";
import * as vscode from "vscode";

const PRESENTATION_EXTENSIONS = [".pptx", ".ppt", ".odp", ".key"];

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(
    vscode.commands.registerCommand("presentationViewer.open", async (uri?: vscode.Uri) => {
      const targetUri =
        uri ??
        vscode.window.activeTextEditor?.document.uri ??
        (await pickPresentationFile());

      if (!targetUri) {
        return;
      }

      if (!isPresentationFile(targetUri.fsPath)) {
        void vscode.window.showWarningMessage("Select a .pptx, .ppt, .odp, or .key file to open in Presentation Viewer.");
        return;
      }

      openPresentation(targetUri.fsPath, context);
    })
  );

  const openListener = vscode.workspace.onDidOpenTextDocument((doc: vscode.TextDocument) => {
    if (!isPresentationFile(doc.uri.fsPath)) {
      return;
    }

    openPresentation(doc.uri.fsPath, context);
    closeIfActive(doc);
  });

  context.subscriptions.push(openListener);
}

async function openPresentation(filePath: string, context: vscode.ExtensionContext) {
  try {
    const panel = vscode.window.createWebviewPanel(
      "presentation",
      `Presentation: ${path.basename(filePath)}`,
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        localResourceRoots: [vscode.Uri.joinPath(context.extensionUri, "src", "webview")]
      }
    );

    panel.webview.html = getWebviewContent(panel.webview, context);

    const data = fs.readFileSync(filePath);
    const base64 = data.toString("base64");

    panel.webview.postMessage({
      type: "loadFile",
      fileName: path.basename(filePath),
      size: data.byteLength,
      base64
    });
  } catch (error) {
    void vscode.window.showErrorMessage(`Unable to open presentation: ${String(error)}`);
  }
}

function isPresentationFile(filePath: string): boolean {
  const lower = filePath.toLowerCase();
  return PRESENTATION_EXTENSIONS.some((ext) => lower.endsWith(ext));
}

function closeIfActive(document: vscode.TextDocument) {
  const active = vscode.window.activeTextEditor;
  if (active && active.document.uri.toString() === document.uri.toString()) {
    void vscode.commands.executeCommand("workbench.action.closeActiveEditor");
  }
}

async function pickPresentationFile(): Promise<vscode.Uri | undefined> {
  const picked = await vscode.window.showOpenDialog({
    filters: {
      Presentations: ["pptx", "ppt", "odp", "key"] // KEY PROBABLY NOT WORKING
    },
    canSelectMany: false,
    openLabel: "Open presentation"
  });

  return picked?.[0];
}

function getWebviewContent(webview: vscode.Webview, context: vscode.ExtensionContext): string {
  const webviewRoot = vscode.Uri.joinPath(context.extensionUri, "src", "webview");
  const htmlPath = vscode.Uri.joinPath(webviewRoot, "index.html");

  const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(webviewRoot, "main.js"));
  const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(webviewRoot, "style.css"));

  const rawHtml = fs.readFileSync(htmlPath.fsPath, "utf8");

  return rawHtml
    .replace(/{{mainScript}}/g, scriptUri.toString())
    .replace(/{{styleSheet}}/g, styleUri.toString())
    .replace(/{{cspSource}}/g, webview.cspSource);
}

export function deactivate() {}
