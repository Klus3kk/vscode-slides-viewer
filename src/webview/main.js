const MAX_SLIDES = 20;
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

window.addEventListener("error", (ev) => log(`Runtime error: ${ev.message}`));
window.addEventListener("unhandledrejection", (ev) => log(`Unhandled rejection: ${ev.reason}`));

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
        if (!raw) return { cx: 10080625, cy: 5670550 };
        const text = await raw.async("text");
        const doc = parseXml(text);
        if (!doc) return { cx: 10080625, cy: 5670550 };
        const sldSz = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "sldSz");
        const cx = sldSz?.getAttribute("cx");
        const cy = sldSz?.getAttribute("cy");
        return {
            cx: cx ? parseInt(cx, 10) : 10080625,
            cy: cy ? parseInt(cy, 10) : 5670550
        };
    } catch (e) {
        log(`Error getting slide size: ${e.message}`);
        return { cx: 10080625, cy: 5670550 };
    }
}

async function getSlideOrder(zip) {
    try {
        const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
        const relsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");
        if (!presentationXml) {
            return Object.keys(zip.files)
                .filter((name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml"))
                .sort();
        }
        const relMap = buildRelationshipMap(relsXml);
        const presDoc = parseXml(presentationXml);
        if (!presDoc) return [];
        const slideIds = Array.from(presDoc.getElementsByTagName("*")).filter((el) => el.localName === "sldId");
        const ordered = slideIds
            .map((el) => el.getAttribute("r:id"))
            .map((rid) => (rid ? relMap[rid] : undefined))
            .filter((p) => p && zip.file(`ppt/${p}`))
            .map((p) => `ppt/${p}`);
        if (ordered.length === 0) {
            return Object.keys(zip.files)
                .filter((name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml"))
                .sort();
        }
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
        if (id && target) map[id] = target;
    }
    return map;
}

async function getSlideMasterPath(zip, slidePath) {
    try {
        // Get slide layout from slide relationships
        const slideRelsPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
        const slideRelsXml = await zip.file(slideRelsPath)?.async("text");
        if (!slideRelsXml) {
            log(`  No slide rels found: ${slideRelsPath}`);
            return null;
        }
        
        const slideRels = buildRelationshipMap(slideRelsXml);
        const layoutRel = Object.entries(slideRels).find(([_, target]) => target.includes("slideLayout"));
        if (!layoutRel) {
            log(`  No layout relationship found`);
            return null;
        }
        
        // Resolve layout path (handle ../ relative paths)
        let layoutPath = layoutRel[1];
        if (layoutPath.startsWith("../")) {
            layoutPath = layoutPath.replace("../", "");
        }
        layoutPath = `ppt/${layoutPath}`;
        log(`  Layout path: ${layoutPath}`);
        
        // Get slide master from layout relationships
        const layoutRelsPath = layoutPath.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        const layoutRelsXml = await zip.file(layoutRelsPath)?.async("text");
        if (!layoutRelsXml) {
            log(`  No layout rels found: ${layoutRelsPath}`);
            return null;
        }
        
        const layoutRels = buildRelationshipMap(layoutRelsXml);
        const masterRel = Object.entries(layoutRels).find(([_, target]) => target.includes("slideMaster"));
        if (!masterRel) {
            log(`  No master relationship found`);
            return null;
        }
        
        // Resolve master path
        let masterPath = masterRel[1];
        if (masterPath.startsWith("../")) {
            masterPath = masterPath.replace("../", "");
        }
        masterPath = `ppt/${masterPath}`;
        log(`  Master path: ${masterPath}`);
        
        return masterPath;
    } catch (e) {
        log(`Error finding slide master: ${e.message}`);
        return null;
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

function getColorFromXml(element) {
    const srgbClr = Array.from(element.getElementsByTagName("*")).find((el) => el.localName === "srgbClr");
    if (srgbClr) {
        const val = srgbClr.getAttribute("val");
        if (val) return `#${val}`;
    }
    return null;
}

function getShapeFill(spPr) {
    if (!spPr) return null;
    
    const solidFill = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "solidFill");
    if (solidFill) {
        const color = getColorFromXml(solidFill);
        if (color) return { type: 'solid', color };
    }
    
    const noFill = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "noFill");
    if (noFill) return { type: 'none' };
    
    return null;
}

function extractTextFromShape(shapeNode) {
    try {
        const txBody = Array.from(shapeNode.children).find((el) => el.localName === "txBody");
        if (!txBody) return null;
        
        // Get body properties for text box padding/alignment
        const bodyPr = Array.from(txBody.children).find((el) => el.localName === "bodyPr");
        let verticalAlign = "center";
        if (bodyPr) {
            const anchor = bodyPr.getAttribute("anchor");
            if (anchor === "t") verticalAlign = "flex-start";
            else if (anchor === "b") verticalAlign = "flex-end";
            else if (anchor === "ctr") verticalAlign = "center";
        }
        
        const paragraphs = Array.from(txBody.getElementsByTagName("*")).filter((el) => el.localName === "p");
        const textData = [];
        
        for (const p of paragraphs) {
            const pPr = Array.from(p.children).find((el) => el.localName === "pPr");
            let align = "left";
            if (pPr) {
                const algnAttr = pPr.getAttribute("algn");
                if (algnAttr === "ctr") align = "center";
                else if (algnAttr === "r") align = "right";
                else if (algnAttr === "l") align = "left";
            }
            
            const runs = Array.from(p.children).filter((el) => el.localName === "r");
            const runData = [];
            
            for (const r of runs) {
                const rPr = Array.from(r.children).find((el) => el.localName === "rPr");
                const style = {};
                
                if (rPr) {
                    const sz = rPr.getAttribute("sz");
                    if (sz) style.fontSize = `${parseInt(sz) / 100}pt`;
                    
                    const b = rPr.getAttribute("b");
                    if (b === "1") style.fontWeight = "bold";
                    
                    const i = rPr.getAttribute("i");
                    if (i === "1") style.fontStyle = "italic";
                    
                    const solidFill = Array.from(rPr.getElementsByTagName("*")).find((el) => el.localName === "solidFill");
                    if (solidFill) {
                        const color = getColorFromXml(solidFill);
                        if (color) style.color = color;
                    }
                    
                    const latinFont = Array.from(rPr.getElementsByTagName("*")).find((el) => el.localName === "latin");
                    if (latinFont) {
                        const typeface = latinFont.getAttribute("typeface");
                        if (typeface) style.fontFamily = typeface;
                    }
                }
                
                const tNodes = Array.from(r.getElementsByTagName("*")).filter((el) => el.localName === "t");
                const text = tNodes.map((t) => t.textContent || "").join("");
                
                if (text) runData.push({ text, style });
            }
            
            if (runData.length > 0) {
                textData.push({ align, runs: runData });
            }
        }
        
        return textData.length > 0 ? { paragraphs: textData, verticalAlign } : null;
    } catch (e) {
        log(`Error extracting text: ${e.message}`);
        return null;
    }
}

async function parseMasterShapes(zip, masterPath) {
    try {
        if (!masterPath) return [];
        
        const masterXml = await zip.file(masterPath)?.async("text");
        if (!masterXml) {
            log(`  Master XML not found: ${masterPath}`);
            return [];
        }
        
        const doc = parseXml(masterXml);
        if (!doc) {
            log(`  Failed to parse master XML`);
            return [];
        }
        
        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        if (!spTree) {
            log(`  No spTree in master`);
            return [];
        }
        
        const shapes = [];
        const spElements = Array.from(spTree.children).filter((el) => el.localName === "sp");
        
        log(`  Found ${spElements.length} shapes in master`);
        
        for (const sp of spElements) {
            // Check if it's a placeholder - we want to skip text placeholders but keep decorative shapes
            const nvSpPr = Array.from(sp.children).find((el) => el.localName === "nvSpPr");
            let isPlaceholder = false;
            
            if (nvSpPr) {
                const nvPr = Array.from(nvSpPr.children).find((el) => el.localName === "nvPr");
                if (nvPr) {
                    const ph = Array.from(nvPr.children).find((el) => el.localName === "ph");
                    // Only skip if it's a content placeholder (title, body, etc)
                    if (ph) {
                        const phType = ph.getAttribute("type");
                        if (phType && (phType === "title" || phType === "body" || phType === "ctrTitle")) {
                            isPlaceholder = true;
                        }
                    }
                }
            }
            
            if (isPlaceholder) continue;
            
            const box = getShapeBox(sp);
            if (!box) continue;
            
            const spPr = Array.from(sp.children).find((el) => el.localName === "spPr");
            const fill = getShapeFill(spPr);
            
            // Even if there's no text, we want to render shapes with fills (decorative elements)
            const textData = extractTextFromShape(sp);
            
            // Include shape if it has a fill or text
            if (fill || textData) {
                shapes.push({
                    type: textData ? "text" : "shape",
                    box,
                    fill,
                    textData,
                    isMaster: true
                });
            }
        }
        
        log(`  Extracted ${shapes.length} shapes from master (skipped placeholders)`);
        return shapes;
    } catch (e) {
        log(`Error parsing master shapes: ${e.message}`);
        return [];
    }
}

async function getSlideRelationships(zip, slidePath) {
    try {
        const relPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
        const relFile = zip.file(relPath);
        if (!relFile) return {};
        const relXml = await relFile.async("text");
        return buildRelationshipMap(relXml);
    } catch (e) {
        log(`Error getting slide relationships: ${e.message}`);
        return {};
    }
}

function resolveMediaPath(slidePath, target) {
    if (target.startsWith("../")) {
        const base = slidePath.split("/").slice(0, -2).join("/");
        return `${base}/${target.replace(/^\.\.\//g, "")}`.replace(/\\/g, "/");
    }
    return `ppt/${target}`;
}

async function parseSlideShapes(zip, slidePath) {
    try {
        const xml = await zip.file(slidePath)?.async("text");
        if (!xml) return [];
        
        const doc = parseXml(xml);
        if (!doc) return [];
        
        const rels = await getSlideRelationships(zip, slidePath);
        
        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        if (!spTree) return [];
        
        const shapes = [];
        const spElements = Array.from(spTree.children).filter((el) => el.localName === "sp" || el.localName === "pic");
        
        for (const node of spElements) {
            const box = getShapeBox(node);
            if (!box) continue;
            
            if (node.localName === "sp") {
                const spPr = Array.from(node.children).find((el) => el.localName === "spPr");
                const fill = getShapeFill(spPr);
                const textData = extractTextFromShape(node);
                
                if (textData || fill) {
                    shapes.push({
                        type: "text",
                        box,
                        fill,
                        textData,
                        isMaster: false
                    });
                }
            } else if (node.localName === "pic") {
                // Extract image
                const blipEl = Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "blip");
                const embed = blipEl?.getAttribute("r:embed") || 
                             blipEl?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
                
                if (!embed) continue;
                
                const target = rels[embed];
                if (!target) continue;
                
                const mediaPath = resolveMediaPath(slidePath, target);
                const mediaFile = zip.file(mediaPath);
                if (!mediaFile) {
                    log(`Image not found: ${mediaPath}`);
                    continue;
                }
                
                const ext = mediaPath.split(".").pop()?.toLowerCase();
                const mimeTypes = {
                    'png': 'image/png',
                    'jpg': 'image/jpeg',
                    'jpeg': 'image/jpeg',
                    'gif': 'image/gif',
                    'bmp': 'image/bmp',
                    'svg': 'image/svg+xml'
                };
                const mime = mimeTypes[ext];
                
                if (!mime) continue;
                
                try {
                    const dataUrl = `data:${mime};base64,${await mediaFile.async("base64")}`;
                    shapes.push({
                        type: "image",
                        box,
                        src: dataUrl,
                        isMaster: false
                    });
                } catch (e) {
                    log(`Error loading image ${mediaPath}: ${e.message}`);
                }
            }
        }
        
        return shapes;
    } catch (e) {
        log(`Error parsing slide shapes: ${e.message}`);
        return [];
    }
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
    log(`Processing ${slidePaths.length} slides`);
    
    const slides = [];
    
    for (let i = 0; i < slidePaths.length; i++) {
        const slidePath = slidePaths[i];
        log(`Processing slide ${i + 1}`);
        
        try {
            const masterPath = await getSlideMasterPath(zip, slidePath);
            const masterShapes = await parseMasterShapes(zip, masterPath);
            const slideShapes = await parseSlideShapes(zip, slidePath);
            
            const allShapes = [...masterShapes, ...slideShapes];
            
            slides.push({
                path: slidePath,
                size: slideSize,
                shapes: allShapes
            });
        } catch (e) {
            log(`Error processing slide ${i + 1}: ${e.message}`);
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
                    
                    let bgStyle = '';
                    if (shape.fill) {
                        if (shape.fill.type === 'solid') {
                            bgStyle = `background: ${shape.fill.color};`;
                        } else if (shape.fill.type === 'none') {
                            bgStyle = 'background: transparent;';
                        }
                    }
                    
                    if (shape.type === "image") {
                        return `<img class="shape image-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;" src="${shape.src}" alt="" />`;
                    }
                    
                    if (shape.textData) {
                        const verticalAlign = shape.textData.verticalAlign || 'center';
                        const textHtml = shape.textData.paragraphs.map(para => {
                            const runHtml = para.runs.map(run => {
                                const styles = [];
                                if (run.style.fontSize) styles.push(`font-size: ${run.style.fontSize}`);
                                if (run.style.fontWeight) styles.push(`font-weight: ${run.style.fontWeight}`);
                                if (run.style.fontStyle) styles.push(`font-style: ${run.style.fontStyle}`);
                                if (run.style.color) styles.push(`color: ${run.style.color}`);
                                if (run.style.fontFamily) styles.push(`font-family: "${run.style.fontFamily}", sans-serif`);
                                
                                const styleAttr = styles.length > 0 ? ` style="${styles.join('; ')}"` : '';
                                return `<span${styleAttr}>${escapeHtml(run.text)}</span>`;
                            }).join('');
                            
                            const textAlign = para.align || 'left';
                            return `<p style="margin: 0; text-align: ${textAlign}; line-height: 1.2;">${runHtml}</p>`;
                        }).join('');
                        
                        return `<div class="shape text-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;${bgStyle}align-items:${verticalAlign};justify-content:${verticalAlign};">${textHtml}</div>`;
                    } else {
                        return `<div class="shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;${bgStyle}"></div>`;
                    }
                })
                .join("");
            
            return `
                <article class="slide-frame" id="slide-${idx}" style="width:${VIEW_WIDTH}px;height:${heightPx}px;">
                    <div class="slide-canvas" style="width:${VIEW_WIDTH}px;height:${heightPx}px;background:#ffffff;">
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
        
        try {
            name.textContent = msg.fileName ?? "Presentation";
            metaContent.innerHTML = `<p><strong>Size:</strong> ${formatBytes(msg.size)}</p>`;
            meta.classList.remove("hidden");
            log(`Received file (${formatBytes(msg.size)})`);
            document.body.dataset.loaded = "true";
            
            if (msg.fileName?.toLowerCase().endsWith(".pptx")) {
                const started = performance.now();
                const slides = await renderPptxSlides(msg.base64);
                slidesCache = slides;
                const durationMs = performance.now() - started;
                log(`Parsed in ${durationMs.toFixed(0)}ms; showing ${slides.length} slides`);
                
                if (slides.length === 0) {
                    slidesContent.innerHTML = "<p>No slides found.</p>";
                    slidesEl.classList.remove("hidden");
                    return;
                }
                
                slidesContent.innerHTML = renderSlidesToHtml(slides);
                slidesEl.classList.remove("hidden");
                currentSlide = 0;
                updateSlideVisibility();
                applyZoom();
                updatePageInfo();
            } else {
                slidesContent.innerHTML = "<p>Only PPTX files supported.</p>";
                slidesEl.classList.remove("hidden");
            }
        } catch (err) {
            metaContent.innerHTML = `<p><strong>Error:</strong> ${String(err)}</p>`;
            meta.classList.remove("hidden");
            slidesContent.innerHTML = `<p>Error: ${String(err)}</p>`;
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
    
    prev?.addEventListener("click", () => changeSlide(-1));
    next?.addEventListener("click", () => changeSlide(1));
    zoomIn?.addEventListener("click", () => changeZoom(0.1));
    zoomOut?.addEventListener("click", () => changeZoom(-0.1));
    zoomReset?.addEventListener("click", () => setZoom(1));
    toggleLog?.addEventListener("click", () => {
        document.getElementById("log")?.classList.toggle("hidden");
    });
    
    const slidesContent = document.getElementById("slides-content");
    slidesContent?.addEventListener("wheel", (e) => {
        if (e.ctrlKey || e.metaKey) {
            e.preventDefault();
            changeZoom(e.deltaY > 0 ? -0.1 : 0.1);
        }
    }, { passive: false });
    
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
    info.textContent = slidesCache.length ? `${currentSlide + 1} / ${slidesCache.length}` : "0 / 0";
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