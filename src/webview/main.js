function formatBytes(bytes) {
    if (bytes === 0) {
        return "0 B";
    }

    const units = ["B", "KB", "MB", "GB"];
    const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / Math.pow(1024, exponent);
    return `${value.toFixed(1)} ${units[exponent]}`;
}

window.addEventListener("message", async (event) => {
    const msg = event.data;

    if (msg?.type === "loadFile") {
        const name = document.getElementById("file-name");
        const meta = document.getElementById("file-meta");
        const base64SizeKb = (msg.base64.length / 1024).toFixed(1);

        name.textContent = msg.fileName ?? "Presentation";
        meta.innerHTML = `
            <p><strong>Size:</strong> ${formatBytes(msg.size)} (${base64SizeKb} KB base64)</p>
            <p><strong>Status:</strong> Bytes delivered to the webview</p>
        `;
    }
});
