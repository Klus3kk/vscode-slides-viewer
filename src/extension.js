const fs = require("fs");
const path = require("path");
const vscode = require("vscode");

const PRESENTATION_EXTENSIONS = [
    ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm",
    ".ppt", ".pps", ".pot",
    ".odp",
    ".key"
];

const VIEW_TYPE = "presentationViewer.viewer";
const channel = vscode.window.createOutputChannel("Presentation Viewer");

function activate(context) {
    channel.appendLine("Presentation Viewer activated.");

    const provider = new PresentationViewerProvider(context);

    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
            webviewOptions: { retainContextWhenHidden: true },
            supportsMultipleEditorsPerDocument: false
        })
    );

    // --- Command: presentationViewer.open ---
    context.subscriptions.push(
        vscode.commands.registerCommand("presentationViewer.open", async (uri) => {
            const targetUri =
                uri ||
                vscode.window.activeTextEditor?.document.uri ||
                await pickPresentationFile();

            if (!targetUri) return;

            if (!isPresentationFile(targetUri.fsPath)) {
                vscode.window.showWarningMessage(
                    "Select a .pptx/.ppt/.odp/.key presentation file."
                );
                return;
            }

            await openWithCustom(targetUri, context);
        })
    );

    // Open automatically when a presentation file is opened
    const openListener = vscode.workspace.onDidOpenTextDocument((doc) => {
        if (isPresentationFile(doc.uri.fsPath)) {
            openWithCustom(doc.uri, context);
        }
    });
    context.subscriptions.push(openListener);

    // For already-open files
    vscode.workspace.textDocuments.forEach((doc) => {
        if (isPresentationFile(doc.uri.fsPath)) {
            openWithCustom(doc.uri, context);
        }
    });
}

async function openWithCustom(uri, context) {
    try {
        await vscode.commands.executeCommand("vscode.openWith", uri, VIEW_TYPE);
    } catch (error) {
        await openPresentation(uri.fsPath, context);
        vscode.window.showWarningMessage(`Opened in fallback viewer: ${String(error)}`);
    }
}

class PresentationViewerProvider {
    constructor(context) {
        this.context = context;
    }

    async openCustomDocument(uri) {
        channel.appendLine(`openCustomDocument: ${uri.fsPath}`);
        return { uri, dispose: () => {} };
    }

    async resolveCustomEditor(document, webviewPanel) {
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
        } catch (err) {
            channel.appendLine("resolveCustomEditor error: " + String(err));
            vscode.window.showErrorMessage("Unable to open presentation: " + String(err));
        }
    }
}

async function openPresentation(filePath, context) {
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

function hydrateWebview(webview, context, filePath) {
    channel.appendLine(`hydrateWebview: ${filePath}`);
    webview.html = getWebviewContent(webview, context);

    const data = fs.readFileSync(filePath);
    const payload = {
        type: "loadFile",
        fileName: path.basename(filePath),
        size: data.byteLength,
        base64: data.toString("base64")
    };

    // If the webview didn't send "ready" quickly, send again
    const timeout = setTimeout(() => {
        channel.appendLine("Webview didn't signal ready – posting payload again.");
        webview.postMessage(payload);
    }, 1500);

    webview.onDidReceiveMessage((msg) => {
        if (msg && msg.type === "ready") {
            clearTimeout(timeout);
            webview.postMessage(payload);
        }
    });

    // Send immediately
    webview.postMessage(payload);
}

function isPresentationFile(filePath) {
    const lower = filePath.toLowerCase();
    return PRESENTATION_EXTENSIONS.some(ext => lower.endsWith(ext));
}

async function pickPresentationFile() {
    const picked = await vscode.window.showOpenDialog({
        canSelectMany: false,
        openLabel: "Open presentation",
        filters: {
            Presentations: [
                "pptx", "pptm", "potx", "potm", "ppsx", "ppsm",
                "ppt", "pps", "pot", "odp", "key"
            ]
        }
    });
    return picked?.[0];
}

function getWebviewContent(webview, context) {
    const root = vscode.Uri.joinPath(context.extensionUri, "src", "webview");
    const htmlPath = vscode.Uri.joinPath(root, "index.html");

    const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(root, "main.js"));
    const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(root, "style.css"));
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

function deactivate() {}

module.exports = {
    activate,
    deactivate
};
