export function decodeBase64ToUint8(base64) {
    const binary = atob(base64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
}

export function parseXml(xml) {
    try {
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        const parseError = doc.querySelector("parsererror");
        if (parseError) return null;
        return doc;
    } catch (e) {
        return null;
    }
}

export function mergeStyles(base, override) {
    return { ...base, ...override };
}

export function escapeHtml(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

export function guessMimeFromBytes(path, bytes) {
    const ext = path.split(".").pop()?.toLowerCase();
    const mimeTypes = {
        png: "image/png",
        jpg: "image/jpeg",
        jpeg: "image/jpeg",
        gif: "image/gif",
        bmp: "image/bmp",
        svg: "image/svg+xml",
        webp: "image/webp",
        avif: "image/avif",
        emf: "image/emf",
        wmf: "image/wmf"
    };
    if (ext && mimeTypes[ext]) return mimeTypes[ext];
    if (bytes.length > 4) {
        if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
        if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
        if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return "image/gif";
        if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
        if (bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46) return "image/webp";
        if (bytes[0] === 0x01 && bytes[1] === 0x00 && bytes[2] === 0x00 && bytes[3] === 0x00) return "image/emf";
        if (bytes[0] === 0xd7 && bytes[1] === 0xcd && bytes[2] === 0xc6 && bytes[3] === 0x9a) return "image/wmf";
    }
    return "image/png";
}

export function uint8ToBase64(bytes) {
    let binary = "";
    for (let i = 0; i < bytes.length; i += 1) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

export function extractTrailingNumber(str) {
    const match = str.match(/(\d+)(?!.*\d)/);
    return match ? parseInt(match[1], 10) : 0;
}

export function getImageDimensions(bytes, mime) {
    // PNG
    if (mime === "image/png" || (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)) {
        if (bytes.length >= 24) {
            const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
            const width = view.getUint32(16, false);
            const height = view.getUint32(20, false);
            if (width && height) return { cx: width, cy: height };
        }
    }
    // JPEG
    if (mime === "image/jpeg" || bytes[0] === 0xff) {
        let offset = 2;
        while (offset + 9 < bytes.length) {
            if (bytes[offset] !== 0xff) break;
            const marker = bytes[offset + 1];
            const len = (bytes[offset + 2] << 8) + bytes[offset + 3];
            if (marker >= 0xc0 && marker <= 0xc3 && offset + 7 < bytes.length) {
                const height = (bytes[offset + 5] << 8) + bytes[offset + 6];
                const width = (bytes[offset + 7] << 8) + bytes[offset + 8];
                if (width && height) return { cx: width, cy: height };
                break;
            }
            offset += 2 + len;
        }
    }
    return null;
}

export function lengthToPx(val) {
    if (!val) return 0;
    const n = parseFloat(val);
    if (isNaN(n)) return 0;
    if (val.includes("mm")) return n * 3.7795275591;
    if (val.includes("cm")) return n * 37.795275591;
    if (val.includes("in")) return n * 96;
    if (val.includes("pt")) return n * (96 / 72);
    return n;
}

export function getOdpStyle(allStyles, name) {
    if (!name) return {};
    return allStyles[name] || {};
}

export function guessVectorPlaceholderLabel(shape) {
    if (!shape) return "";
    if (shape.textData && shape.textData.paragraphs) {
        const txt = shape.textData.paragraphs
            .map(p => p.runs.map(r => r.text || "").join(" "))
            .join(" ")
            .trim()
            .toLowerCase();
        return txt || "";
    }
    return "";
}
