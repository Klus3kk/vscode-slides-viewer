const MAX_SLIDES = 10;
const VIEW_WIDTH = 960; // px width for preview; height scales from slide size
let slidesCache = [];
let currentSlide = 0;
let zoom = 1;
const vscode = acquireVsCodeApi();

window.addEventListener("DOMContentLoaded", () => {
    log("Webview ready; requesting file bytes.");
    vscode.postMessage({ type: "ready" });

    bindControls();

    setTimeout(() => {
        if (!document.body.dataset.loaded) {
            log("Waiting for file bytes…");
        }
    }, 1000);
});

function log(message) {
    const logContainer = document.getElementById("log");
    const entries = document.getElementById("log-entries");
    const ts = new Date().toLocaleTimeString();
    entries.insertAdjacentHTML("afterbegin", `<div>[${ts}] ${message}</div>`);
    logContainer.classList.remove("hidden");
}

window.addEventListener("error", (ev) => {
    log(`Runtime error: ${ev.message}`);
});

window.addEventListener("unhandledrejection", (ev) => {
    log(`Unhandled rejection: ${ev.reason}`);
});

function formatBytes(bytes) {
    if (bytes === 0) {
        return "0 B";
    }

    const units = ["B", "KB", "MB", "GB"];
    const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / Math.pow(1024, exponent);
    return `${value.toFixed(1)} ${units[exponent]}`;
}

function decodeBase64ToUint8(base64) {
    const binary = atob(base64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
}

function parseXml(xml) {
    return new DOMParser().parseFromString(xml, "application/xml");
}

async function getSlideSize(zip) {
    const raw = zip.file("ppt/presentation.xml");
    if (!raw) return { cx: 9144000, cy: 6858000 }; // default 16:9

    const text = await raw.async("text");
    const doc = parseXml(text);
    const sldSz = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "sldSz");
    const cx = sldSz?.getAttribute("cx");
    const cy = sldSz?.getAttribute("cy");
    return {
        cx: cx ? parseInt(cx, 10) : 9144000,
        cy: cy ? parseInt(cy, 10) : 6858000
    };
}

async function getSlideOrder(zip) {
    const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
    const relsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");

    if (!presentationXml) {
        log("presentation.xml missing; falling back to alphabetical slides.");
        return Object.keys(zip.files)
            .filter((name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml"))
            .sort();
    }

    const relMap = buildRelationshipMap(relsXml);
    const presDoc = parseXml(presentationXml);

    const slideIds = Array.from(presDoc.getElementsByTagName("*")).filter((el) => el.localName === "sldId");
    const ordered = slideIds
        .map((el) => el.getAttribute("r:id"))
        .map((rid) => (rid ? relMap[rid] : undefined))
        .filter((p) => p && zip.file(`ppt/${p}`))
        .map((p) => `ppt/${p}`);

    if (ordered.length === 0) {
        log("No slide order found; falling back to alphabetical slides.");
        return Object.keys(zip.files)
            .filter((name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml"))
            .sort();
    }

    return ordered;
}

function buildRelationshipMap(relsXml) {
    if (!relsXml) return {};
    const doc = parseXml(relsXml);
    const rels = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "Relationship");
    const map = {};
    for (const rel of rels) {
        const id = rel.getAttribute("Id");
        const target = rel.getAttribute("Target");
        if (id && target) {
            map[id] = target;
        }
    }
    return map;
}

function resolveMediaPath(slidePath, target) {
    if (target.startsWith("../")) {
        const base = slidePath.split("/").slice(0, -2).join("/"); // drop slideX.xml and slides
        return `${base}/${target.replace(/^\\.\\//, "").replace("^../", "")}`.replace(/\\/g, "/");
    }
    return `ppt/${target}`;
}

async function getSlideRelationships(zip, slidePath) {
    const relPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
    const relFile = zip.file(relPath);
    if (!relFile) {
        return {};
    }
    const relXml = await relFile.async("text");
    return buildRelationshipMap(relXml);
}

function extractTextsFromSlide(xml) {
    const doc = parseXml(xml);
    const paragraphs = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "p");
    const lines = [];

    for (const p of paragraphs) {
        const runs = Array.from(p.getElementsByTagName("*")).filter((el) => el.localName === "r" || el.localName === "br");
        const parts = [];
        for (const r of runs) {
            if (r.localName === "br") {
                parts.push("\n");
                continue;
            }
            const text = Array.from(r.getElementsByTagName("*"))
                .filter((el) => el.localName === "t")
                .map((t) => t.textContent || "")
                .join("");
            parts.push(text);
        }
        const line = parts.join("").trim();
        if (line) lines.push(line);
    }

    return lines;
}

function getShapeBox(shapeEl) {
    const xfrm = Array.from(shapeEl.getElementsByTagName("*")).find((el) => el.localName === "xfrm");
    if (!xfrm) return undefined;
    const off = Array.from(xfrm.children).find((el) => el.localName === "off");
    const ext = Array.from(xfrm.children).find((el) => el.localName === "ext");
    if (!off || !ext) return undefined;
    return {
        x: parseInt(off.getAttribute("x") ?? "0", 10),
        y: parseInt(off.getAttribute("y") ?? "0", 10),
        cx: parseInt(ext.getAttribute("cx") ?? "0", 10),
        cy: parseInt(ext.getAttribute("cy") ?? "0", 10)
    };
}

function emuToPx(value, scale) {
    return (value / 9144000) * (VIEW_WIDTH * (9144000 / scale.cx));
}

function escapeHtml(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

async function renderPptxSlides(base64) {
    const buffer = decodeBase64ToUint8(base64);
    const zip = await JSZip.loadAsync(buffer);

    const slideSize = await getSlideSize(zip);
    const slidePaths = (await getSlideOrder(zip)).slice(0, MAX_SLIDES);
    const slides = [];

    for (const slidePath of slidePaths) {
        const xml = await zip.file(slidePath)?.async("text");
        if (!xml) continue;

        const rels = await getSlideRelationships(zip, slidePath);
        const doc = parseXml(xml);
        const shapes = [];

        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        if (!spTree) continue;

        const nodes = Array.from(spTree.children).filter((el) => el.localName === "sp" || el.localName === "pic");

        for (const node of nodes) {
            const box = getShapeBox(node);
            if (!box) continue;

            if (node.localName === "sp") {
                const texts = extractTextsFromShape(node);
                shapes.push({
                    type: "text",
                    box,
                    text: texts.join("\n")
                });
            } else if (node.localName === "pic") {
                const embed = Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "blip")?.getAttribute("r:embed") ??
                    Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "blip")?.getAttribute("r:embed".replace("r:", "r:"));
                if (!embed) continue;
                const target = rels[embed];
                if (!target) continue;
                const mediaPath = resolveMediaPath(slidePath, target);
                const mediaFile = zip.file(mediaPath);
                if (!mediaFile) continue;
                const ext = mediaPath.split(".").pop()?.toLowerCase();
                const mime = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "bmp" ? "image/bmp" : undefined;
                if (!mime) continue;
                const dataUrl = `data:${mime};base64,${await mediaFile.async("base64")}`;
                shapes.push({
                    type: "image",
                    box,
                    src: dataUrl
                });
            }
        }

        slides.push({
            path: slidePath,
            size: slideSize,
            shapes
        });
    }

    return slides;
}

function extractTextsFromShape(shapeNode) {
    const txBody = Array.from(shapeNode.children).find((el) => el.localName === "txBody");
    if (!txBody) return [];
    const paragraphs = Array.from(txBody.getElementsByTagName("*")).filter((el) => el.localName === "p");
    const texts = [];
    for (const p of paragraphs) {
        const runs = Array.from(p.children).filter((el) => el.localName === "r" || el.localName === "br");
        const parts = [];
        for (const r of runs) {
            if (r.localName === "br") {
                parts.push("\n");
                continue;
            }
            const tNodes = Array.from(r.getElementsByTagName("*")).filter((el) => el.localName === "t");
            const tText = tNodes.map((t) => t.textContent || "").join("");
            parts.push(tText);
        }
        const line = parts.join("").replace(/\n+/g, "\n").trim();
        if (line) texts.push(line);
    }
    return texts;
}

function renderSlidesToHtml(slides) {
    return slides
        .map((slide, idx) => {
            const scale = VIEW_WIDTH / slide.size.cx;
            const heightPx = Math.round(slide.size.cy * scale);

            const shapesHtml = slide.shapes
                .map((shape) => {
                    const left = Math.round(shape.box.x * scale);
                    const top = Math.round(shape.box.y * scale);
                    const width = Math.round(shape.box.cx * scale);
                    const height = Math.round(shape.box.cy * scale);

                    if (shape.type === "text") {
                        const paragraphs = escapeHtml(shape.text).split("\n").join("<br>");
                        return `<div class="shape text-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;">${paragraphs}</div>`;
                    }

                    if (shape.type === "image") {
                        return `<img class="shape image-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;" src="${shape.src}" />`;
                    }

                    return "";
                })
                .join("");

            return `
                <article class="slide-frame" id="slide-${idx}" style="width:${VIEW_WIDTH}px;height:${heightPx}px;">
                    <div class="slide-canvas" style="width:${VIEW_WIDTH}px;height:${heightPx}px;">
                        ${shapesHtml}
                    </div>
                    <div class="slide-label">Slide ${idx + 1}</div>
                </article>
            `;
        })
        .join("");
}

window.addEventListener("message", async (event) => {
    const msg = event.data;

    if (msg?.type === "loadFile") {
        const name = document.getElementById("file-name");
        const meta = document.getElementById("file-meta");
        const metaContent = document.getElementById("file-meta-content");
        const slidesEl = document.getElementById("slides");
        const slidesContent = document.getElementById("slides-content");
        const base64SizeKb = (msg.base64.length / 1024).toFixed(1);

        try {
            name.textContent = msg.fileName ?? "Presentation";
            metaContent.innerHTML = `<p><strong>Size:</strong> ${formatBytes(msg.size)} (${base64SizeKb} KB base64)</p>`;
            meta.classList.remove("hidden");
            log(`Received file bytes (${formatBytes(msg.size)})`);
            document.body.dataset.loaded = "true";

            if (msg.fileName?.toLowerCase().endsWith(".pptx")) {
                const started = performance.now();
                const slides = await renderPptxSlides(msg.base64);
                slidesCache = slides;
                const durationMs = performance.now() - started;
                log(`Parsed PPTX in ${durationMs.toFixed(0)}ms; showing ${slides.length} slide(s).`);
                if (slides.length === 0) {
                    slidesContent.innerHTML = "<p>No slides found in this PPTX.</p>";
                    slidesEl.classList.remove("hidden");
                    return;
                }

                slidesContent.innerHTML = renderSlidesToHtml(slides);
                slidesContent.insertAdjacentHTML(
                    "afterbegin",
                    `<p class="hint">Previewing ${slides.length} slides (parsed in ${durationMs.toFixed(0)}ms).</p>`
                );
                slidesEl.classList.remove("hidden");
                currentSlide = 0;
                applyZoom();
                scrollToSlide(0, false);
            } else {
                slidesContent.innerHTML = "<p>Rendering currently supports PPTX text only.</p>";
                slidesEl.classList.remove("hidden");
                log("Non-PPTX file: only size shown.");
            }
        } catch (err) {
            metaContent.innerHTML = `<p><strong>Error:</strong> ${String(err)}</p>`;
            meta.classList.remove("hidden");
            slidesContent.innerHTML = "";
            slidesEl.classList.remove("hidden");
            log(`Error: ${String(err)}`);
        }
    }
});

function bindControls() {
    const prev = document.getElementById("prev");
    const next = document.getElementById("next");
    const zoomIn = document.getElementById("zoom-in");
    const zoomOut = document.getElementById("zoom-out");
    const zoomReset = document.getElementById("zoom-reset");
    const toggleLog = document.getElementById("toggle-log");

    prev?.addEventListener("click", () => scrollToSlide(currentSlide - 1));
    next?.addEventListener("click", () => scrollToSlide(currentSlide + 1));
    zoomIn?.addEventListener("click", () => changeZoom(0.1));
    zoomOut?.addEventListener("click", () => changeZoom(-0.1));
    zoomReset?.addEventListener("click", () => setZoom(1));
    toggleLog?.addEventListener("click", () => {
        document.getElementById("log")?.classList.toggle("hidden");
    });
}

function scrollToSlide(index, smooth = true) {
    if (!slidesCache.length) return;
    currentSlide = Math.min(Math.max(index, 0), slidesCache.length - 1);
    const el = document.getElementById(`slide-${currentSlide}`);
    if (el) {
        el.scrollIntoView({ behavior: smooth ? "smooth" : "auto", block: "start" });
    }
    updatePageInfo();
}

function updatePageInfo() {
    const info = document.getElementById("page-info");
    if (!info) return;
    if (!slidesCache.length) {
        info.textContent = "0 / 0";
    } else {
        info.textContent = `${currentSlide + 1} / ${slidesCache.length}`;
    }
}

function changeZoom(delta) {
    setZoom(zoom + delta);
}

function setZoom(value) {
    zoom = Math.min(Math.max(value, 0.5), 2);
    applyZoom();
}

function applyZoom() {
    const slidesContent = document.getElementById("slides-content");
    const zoomReset = document.getElementById("zoom-reset");
    if (slidesContent) {
        slidesContent.style.transform = `scale(${zoom})`;
        slidesContent.style.transformOrigin = "top left";
    }
    if (zoomReset) {
        zoomReset.textContent = `${Math.round(zoom * 100)}%`;
    }
    updatePageInfo();
}
