const MAX_SLIDES = 10;
const VIEW_WIDTH = 960;
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
    if (bytes === 0) return "0 B";
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
    try {
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        const parseError = doc.querySelector("parsererror");
        if (parseError) {
            log(`XML parse error: ${parseError.textContent}`);
            return null;
        }
        return doc;
    } catch (e) {
        log(`XML parsing exception: ${e.message}`);
        return null;
    }
}

async function getSlideSize(zip) {
    try {
        const raw = zip.file("ppt/presentation.xml");
        if (!raw) {
            log("No presentation.xml found, using default slide size");
            return { cx: 9144000, cy: 6858000 };
        }

        const text = await raw.async("text");
        const doc = parseXml(text);
        if (!doc) return { cx: 9144000, cy: 6858000 };

        const sldSz = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "sldSz");
        const cx = sldSz?.getAttribute("cx");
        const cy = sldSz?.getAttribute("cy");
        
        log(`Slide size: ${cx}x${cy} EMUs`);
        
        return {
            cx: cx ? parseInt(cx, 10) : 9144000,
            cy: cy ? parseInt(cy, 10) : 6858000
        };
    } catch (e) {
        log(`Error getting slide size: ${e.message}`);
        return { cx: 9144000, cy: 6858000 };
    }
}

async function getSlideOrder(zip) {
    try {
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
        if (!presDoc) {
            log("Failed to parse presentation.xml");
            return [];
        }

        const slideIds = Array.from(presDoc.getElementsByTagName("*")).filter((el) => el.localName === "sldId");
        log(`Found ${slideIds.length} slide IDs in presentation.xml`);

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

        log(`Slide order: ${ordered.join(", ")}`);
        return ordered;
    } catch (e) {
        log(`Error getting slide order: ${e.message}`);
        return [];
    }
}

function buildRelationshipMap(relsXml) {
    if (!relsXml) return {};
    const doc = parseXml(relsXml);
    if (!doc) return {};
    
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
        const base = slidePath.split("/").slice(0, -2).join("/");
        return `${base}/${target.replace(/^\.\.\//g, "")}`.replace(/\\/g, "/");
    }
    return `ppt/${target}`;
}

async function getSlideRelationships(zip, slidePath) {
    try {
        const relPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
        const relFile = zip.file(relPath);
        if (!relFile) {
            return {};
        }
        const relXml = await relFile.async("text");
        return buildRelationshipMap(relXml);
    } catch (e) {
        log(`Error getting slide relationships: ${e.message}`);
        return {};
    }
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

function escapeHtml(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function extractTextsFromShape(shapeNode) {
    try {
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
    } catch (e) {
        log(`Error extracting text from shape: ${e.message}`);
        return [];
    }
}

async function renderPptxSlides(base64) {
    const buffer = decodeBase64ToUint8(base64);
    log(`Decoded base64 to ${buffer.length} bytes`);
    
    const zip = await JSZip.loadAsync(buffer);
    log(`Loaded ZIP with ${Object.keys(zip.files).length} files`);

    const slideSize = await getSlideSize(zip);
    const slidePaths = (await getSlideOrder(zip)).slice(0, MAX_SLIDES);
    log(`Processing ${slidePaths.length} slides`);
    
    const slides = [];

    for (let i = 0; i < slidePaths.length; i++) {
        const slidePath = slidePaths[i];
        log(`Processing slide ${i + 1}: ${slidePath}`);
        
        try {
            const xml = await zip.file(slidePath)?.async("text");
            if (!xml) {
                log(`  No XML content for ${slidePath}`);
                continue;
            }

            const rels = await getSlideRelationships(zip, slidePath);
            const doc = parseXml(xml);
            if (!doc) {
                log(`  Failed to parse XML for ${slidePath}`);
                continue;
            }

            const shapes = [];
            const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
            
            if (!spTree) {
                log(`  No spTree found in ${slidePath}`);
                continue;
            }

            const nodes = Array.from(spTree.children).filter((el) => el.localName === "sp" || el.localName === "pic");
            log(`  Found ${nodes.length} shapes/pictures`);

            for (const node of nodes) {
                const box = getShapeBox(node);
                if (!box) continue;

                if (node.localName === "sp") {
                    const texts = extractTextsFromShape(node);
                    if (texts.length > 0) {
                        shapes.push({
                            type: "text",
                            box,
                            text: texts.join("\n")
                        });
                    }
                } else if (node.localName === "pic") {
                    const blipEl = Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "blip");
                    const embed = blipEl?.getAttribute("r:embed") || blipEl?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
                    
                    if (!embed) continue;
                    
                    const target = rels[embed];
                    if (!target) continue;
                    
                    const mediaPath = resolveMediaPath(slidePath, target);
                    const mediaFile = zip.file(mediaPath);
                    if (!mediaFile) continue;
                    
                    const ext = mediaPath.split(".").pop()?.toLowerCase();
                    const mime = ext === "png" ? "image/png" : 
                                ext === "jpg" || ext === "jpeg" ? "image/jpeg" : 
                                ext === "gif" ? "image/gif" : 
                                ext === "bmp" ? "image/bmp" : undefined;
                    
                    if (!mime) continue;
                    
                    const dataUrl = `data:${mime};base64,${await mediaFile.async("base64")}`;
                    shapes.push({
                        type: "image",
                        box,
                        src: dataUrl
                    });
                }
            }

            log(`  Extracted ${shapes.length} shapes from slide ${i + 1}`);
            slides.push({
                path: slidePath,
                size: slideSize,
                shapes
            });
        } catch (e) {
            log(`  Error processing slide ${i + 1}: ${e.message}`);
        }
    }

    return slides;
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
                <article class="slide-frame" id="slide-${idx}" data-slide-index="${idx}" style="width:${VIEW_WIDTH}px;height:${heightPx}px;">
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
                updateSlideVisibility();
                applyZoom();
                updatePageInfo();
            } else {
                slidesContent.innerHTML = "<p>Rendering currently supports PPTX files only.</p>";
                slidesEl.classList.remove("hidden");
                log("Non-PPTX file: only size shown.");
            }
        } catch (err) {
            metaContent.innerHTML = `<p><strong>Error:</strong> ${String(err)}</p>`;
            meta.classList.remove("hidden");
            slidesContent.innerHTML = `<p>Error loading presentation: ${String(err)}</p>`;
            slidesEl.classList.remove("hidden");
            log(`Error: ${String(err)}`);
            console.error(err);
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

    prev?.addEventListener("click", () => changeSlide(-1));
    next?.addEventListener("click", () => changeSlide(1));
    zoomIn?.addEventListener("click", () => changeZoom(0.1));
    zoomOut?.addEventListener("click", () => changeZoom(-0.1));
    zoomReset?.addEventListener("click", () => setZoom(1));
    toggleLog?.addEventListener("click", () => {
        document.getElementById("log")?.classList.toggle("hidden");
    });

    // Mouse wheel zoom
    const slidesContent = document.getElementById("slides-content");
    slidesContent?.addEventListener("wheel", (e) => {
        if (e.ctrlKey || e.metaKey) {
            e.preventDefault();
            const delta = e.deltaY > 0 ? -0.1 : 0.1;
            changeZoom(delta);
        }
    }, { passive: false });

    // Keyboard navigation
    window.addEventListener("keydown", (e) => {
        if (e.key === "ArrowLeft" || e.key === "PageUp") {
            e.preventDefault();
            changeSlide(-1);
        } else if (e.key === "ArrowRight" || e.key === "PageDown" || e.key === " ") {
            e.preventDefault();
            changeSlide(1);
        } else if (e.key === "Home") {
            e.preventDefault();
            goToSlide(0);
        } else if (e.key === "End") {
            e.preventDefault();
            goToSlide(slidesCache.length - 1);
        }
    });
}

function changeSlide(delta) {
    if (!slidesCache.length) return;
    goToSlide(currentSlide + delta);
}

function goToSlide(index) {
    if (!slidesCache.length) return;
    currentSlide = Math.min(Math.max(index, 0), slidesCache.length - 1);
    updateSlideVisibility();
    updatePageInfo();
}

function updateSlideVisibility() {
    const slides = document.querySelectorAll(".slide-frame");
    slides.forEach((slide, idx) => {
        if (idx === currentSlide) {
            slide.classList.remove("hidden");
        } else {
            slide.classList.add("hidden");
        }
    });
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
    zoom = Math.min(Math.max(value, 0.5), 3);
    applyZoom();
}

function applyZoom() {
    const slidesContent = document.getElementById("slides-content");
    const zoomReset = document.getElementById("zoom-reset");
    if (slidesContent) {
        slidesContent.style.transform = `scale(${zoom})`;
        slidesContent.style.transformOrigin = "top center";
    }
    if (zoomReset) {
        zoomReset.textContent = `${Math.round(zoom * 100)}%`;
    }
}