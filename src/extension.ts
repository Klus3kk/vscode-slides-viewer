import * as fs from "fs";
import * as path from "path";
import * as vscode from "vscode";

const PRESENTATION_EXTENSIONS = [
  ".pptx",
  ".pptm",
  ".potx",
  ".potm",
  ".ppsx",
  ".ppsm",
  ".ppt",
  ".pps",
  ".pot",
  ".odp",
  ".key"
];
const VIEW_TYPE = "presentationViewer.viewer";
const channel = vscode.window.createOutputChannel("Presentation Viewer");

export function activate(context: vscode.ExtensionContext) {
  channel.appendLine("Presentation Viewer activated.");

  const provider = new PresentationViewerProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false
    })
  );

  const openWithCustom = async (uri: vscode.Uri) => {
    try {
      await vscode.commands.executeCommand("vscode.openWith", uri, VIEW_TYPE);
    } catch (error) {
      // Fall back to an ad-hoc webview if openWith fails for any reason.
      await openPresentation(uri.fsPath, context);
      void vscode.window.showWarningMessage(`Opened with fallback viewer: ${String(error)}`);
    }
  };

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
        void vscode.window.showWarningMessage(
          "Select a .pptx/.pptm/.potx/.potm/.ppsx/.ppsm, .ppt/.pps/.pot, .odp or .key file to open."
        );
        return;
      }

      await openWithCustom(targetUri);
    })
  );

  const openListener = vscode.workspace.onDidOpenTextDocument((doc: vscode.TextDocument) => {
    if (!isPresentationFile(doc.uri.fsPath)) {
      return;
    }

    void openWithCustom(doc.uri);
  });

  context.subscriptions.push(openListener);

  // Handle already-open documents when the extension activates.
  vscode.workspace.textDocuments.forEach((doc) => {
    if (isPresentationFile(doc.uri.fsPath)) {
      void openWithCustom(doc.uri);
    }
  });
}

class PresentationViewerProvider implements vscode.CustomReadonlyEditorProvider<vscode.CustomDocument> {
  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(uri: vscode.Uri): Promise<vscode.CustomDocument> {
    channel.appendLine(`openCustomDocument: ${uri.fsPath}`);
    return { uri, dispose: () => undefined };
  }

  async resolveCustomEditor(document: vscode.CustomDocument, webviewPanel: vscode.WebviewPanel): Promise<void> {
    channel.appendLine(`resolveCustomEditor: ${document.uri.fsPath}`);

    const assetRoots = [
      vscode.Uri.joinPath(this.context.extensionUri, "src", "webview"),
      vscode.Uri.joinPath(this.context.extensionUri, "node_modules", "jszip", "dist"),
      vscode.Uri.joinPath(this.context.extensionUri, "node_modules", "cfb", "dist")
    ];

    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: assetRoots
    };

    try {
      hydrateWebview(webviewPanel.webview, this.context, document.uri.fsPath);
    } catch (error) {
      channel.appendLine(`resolveCustomEditor error: ${String(error)}`);
      void vscode.window.showErrorMessage(`Unable to open presentation: ${String(error)}`);
    }
  }
}

async function openPresentation(filePath: string, context: vscode.ExtensionContext) {
  channel.appendLine(`openPresentation: ${filePath}`);

  const assetRoots = [
    vscode.Uri.joinPath(context.extensionUri, "src", "webview"),
    vscode.Uri.joinPath(context.extensionUri, "node_modules", "jszip", "dist"),
    vscode.Uri.joinPath(context.extensionUri, "node_modules", "cfb", "dist")
  ];

  const panel = vscode.window.createWebviewPanel(
    "presentation",
    `Presentation: ${path.basename(filePath)}`,
    vscode.ViewColumn.Active,
    {
      enableScripts: true,
      localResourceRoots: assetRoots
    }
  );

  hydrateWebview(panel.webview, context, filePath);
}

function hydrateWebview(webview: vscode.Webview, context: vscode.ExtensionContext, filePath: string) {
  channel.appendLine(`hydrateWebview: ${filePath}`);
  webview.html = getWebviewContent(webview, context);

  const data = fs.readFileSync(filePath);
  const payload = {
    type: "loadFile",
    fileName: path.basename(filePath),
    size: data.byteLength,
    base64: data.toString("base64")
  };

  const readyTimeout = setTimeout(() => {
    channel.appendLine("Webview did not signal ready within 2s; posting again.");
    void webview.postMessage(payload);
  }, 2000);

  const readyListener = webview.onDidReceiveMessage((msg) => {
    if (msg?.type === "ready") {
      channel.appendLine("Webview signaled ready; posting payload.");
      void webview.postMessage(payload);
      clearTimeout(readyTimeout);
    }
  });

  // Fire once in case the webview is already ready.
  channel.appendLine("Posting payload immediately.");
  void webview.postMessage(payload);
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
      Presentations: [
        "pptx",
        "pptm",
        "potx",
        "potm",
        "ppsx",
        "ppsm",
        "ppt",
        "pps",
        "pot",
        "odp",
        "key"
      ]
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
  const jszipUri = webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri, "node_modules", "jszip", "dist", "jszip.min.js"));
  const cfbUri = webview.asWebviewUri(vscode.Uri.joinPath(context.extensionUri, "node_modules", "cfb", "dist", "cfb.js"));

  const rawHtml = fs.readFileSync(htmlPath.fsPath, "utf8");

  return rawHtml
    .replace(/{{mainScript}}/g, scriptUri.toString())
    .replace(/{{styleSheet}}/g, styleUri.toString())
    .replace(/{{jszip}}/g, jszipUri.toString())
    .replace(/{{cfb}}/g, cfbUri.toString())
    .replace(/{{cspSource}}/g, webview.cspSource);
}

export function deactivate() {}
