const MAX_SLIDES = 20;
const VIEW_WIDTH = 960;
let slidesCache = [];
let currentSlide = 0;
let zoom = 1;
const vscode = acquireVsCodeApi();

window.addEventListener("DOMContentLoaded", () => {
    vscode.postMessage({ type: "ready" });
    bindControls();
});

function log(message) {
    // Logging UI removed; keep a safe no-op.
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
        if (parseError) return null;
        return doc;
    } catch (e) {
        return null;
    }
}

function mergeStyles(base, override) {
    return { ...base, ...override };
}

function getPlaceholderType(shapeNode) {
    const nvSpPr = Array.from(shapeNode.children).find((el) => el.localName === "nvSpPr");
    const nvPr = nvSpPr ? Array.from(nvSpPr.children).find((el) => el.localName === "nvPr") : undefined;
    const ph = nvPr ? Array.from(nvPr.children).find((el) => el.localName === "ph") : undefined;
    return ph?.getAttribute("type") || null;
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
        const slideRelsPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
        const slideRelsXml = await zip.file(slideRelsPath)?.async("text");
        if (!slideRelsXml) return null;
        
        const slideRels = buildRelationshipMap(slideRelsXml);
        const layoutRel = Object.entries(slideRels).find(([_, target]) => target.includes("slideLayout"));
        if (!layoutRel) return null;
        
        let layoutPath = layoutRel[1];
        if (layoutPath.startsWith("../")) {
            layoutPath = layoutPath.replace("../", "");
        }
        layoutPath = `ppt/${layoutPath}`;
        
        const layoutRelsPath = layoutPath.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        const layoutRelsXml = await zip.file(layoutRelsPath)?.async("text");
        if (!layoutRelsXml) return null;
        
        const layoutRels = buildRelationshipMap(layoutRelsXml);
        const masterRel = Object.entries(layoutRels).find(([_, target]) => target.includes("slideMaster"));
        if (!masterRel) return null;
        
        let masterPath = masterRel[1];
        if (masterPath.startsWith("../")) {
            masterPath = masterPath.replace("../", "");
        }
        masterPath = `ppt/${masterPath}`;
        
        return masterPath;
    } catch (e) {
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

function getShapeGeometry(spPr) {
    if (!spPr) return null;
    const prstGeom = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "prstGeom");
    return prstGeom?.getAttribute("prst") || null;
}

function parseRPrStyle(rPr) {
    const style = {};
    if (!rPr) return style;

    const sz = rPr.getAttribute("sz");
    if (sz) style.fontSize = `${parseInt(sz, 10) / 100}pt`;

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

    return style;
}

function extractTextFromShape(shapeNode) {
    try {
        const txBody = Array.from(shapeNode.children).find((el) => el.localName === "txBody");
        if (!txBody) {
            console.log("No txBody found in shape");
            return null;
        }

        const placeholderType = getPlaceholderType(shapeNode);
        console.log("Placeholder type:", placeholderType);
        
        const shapeDefault = parseRPrStyle(Array.from(txBody.querySelectorAll("defRPr"))[0]);
        
        const bodyPr = Array.from(txBody.children).find((el) => el.localName === "bodyPr");
        let verticalAlign = "center";
        if (bodyPr) {
            const anchor = bodyPr.getAttribute("anchor");
            if (anchor === "t") verticalAlign = "flex-start";
            else if (anchor === "b") verticalAlign = "flex-end";
            else if (anchor === "ctr") verticalAlign = "center";
        }
        
        const paragraphs = Array.from(txBody.getElementsByTagName("*")).filter((el) => el.localName === "p");
        console.log(`Found ${paragraphs.length} paragraphs`);
        
        const textData = [];
        
        for (const [paraIdx, p] of paragraphs.entries()) {
            const pPr = Array.from(p.children).find((el) => el.localName === "pPr");
            let align = "left";
            let bullet = null;
            let level = 0;
            let marL = 0;
            let indent = 0;
            const paraDefaults = parseRPrStyle(Array.from(pPr?.children || []).find((el) => el.localName === "defRPr"));
            
            if (pPr) {
                const algnAttr = pPr.getAttribute("algn");
                if (algnAttr === "ctr") align = "center";
                else if (algnAttr === "r") align = "right";
                else if (algnAttr === "l") align = "left";

                marL = parseInt(pPr.getAttribute("marL") || "0", 10);
                indent = parseInt(pPr.getAttribute("indent") || "0", 10);
                const lvlAttr = pPr.getAttribute("lvl");
                if (lvlAttr) level = parseInt(lvlAttr, 10) || 0;

                const buChar = Array.from(pPr.children).find((el) => el.localName === "buChar");
                if (buChar) {
                    const ch = buChar.getAttribute("char") || "■";
                    bullet = { type: "char", char: ch, level };
                } else if (Array.from(pPr.children).some((el) => el.localName === "buAutoNum")) {
                    const auto = Array.from(pPr.children).find((el) => el.localName === "buAutoNum");
                    const startAt = auto ? parseInt(auto.getAttribute("startAt") || "1", 10) : 1;
                    bullet = { type: "auto", index: startAt + paraIdx, level };
                }
            }
            
            const runs = Array.from(p.children).filter((el) => el.localName === "r");
            const runData = [];
            
            for (const r of runs) {
                const rPr = Array.from(r.children).find((el) => el.localName === "rPr");
                const style = mergeStyles(shapeDefault, mergeStyles(paraDefaults, parseRPrStyle(rPr)));
                
                // CRITICAL: Set default font sizes
                if (!style.fontSize) {
                    style.fontSize = placeholderType === "title" || placeholderType === "ctrTitle" ? "44pt" : "28pt";
                }
                if (!style.fontWeight && (placeholderType === "title" || placeholderType === "ctrTitle")) {
                    style.fontWeight = "bold";
                }
                
                const tNodes = Array.from(r.getElementsByTagName("*")).filter((el) => el.localName === "t");
                const text = tNodes.map((t) => t.textContent || "").join("");
                
                console.log("Text run:", text, "style:", style);
                
                if (text) runData.push({ text, style });
            }
            
            if (runData.length > 0) {
                textData.push({ align, runs: runData, bullet, level, marL, indent });
            }
        }
        
        console.log(`Returning ${textData.length} text paragraphs`);
        return textData.length > 0 ? { paragraphs: textData, verticalAlign } : null;
    } catch (e) {
        console.error("Error in extractTextFromShape:", e);
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
            const geom = getShapeGeometry(spPr);
            
            // Even if there's no text, we want to render shapes with fills (decorative elements)
            const textData = extractTextFromShape(sp);
            
            // Include shape if it has a fill or text
            if (fill || textData) {
                shapes.push({
                    type: textData ? "text" : "shape",
                    box,
                    fill,
                    textData,
                    geom,
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
        if (!xml) {
            console.log("No XML for slide");
            return [];
        }
        
        const doc = parseXml(xml);
        if (!doc) {
            console.log("Failed to parse XML");
            return [];
        }
        
        const rels = await getSlideRelationships(zip, slidePath);
        
        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        if (!spTree) {
            console.log("No spTree found");
            return [];
        }
        
        const shapes = [];
        const spElements = Array.from(spTree.children).filter((el) => el.localName === "sp" || el.localName === "pic");
        
        console.log(`Found ${spElements.length} elements in slide`);
        
        for (const node of spElements) {
            const box = getShapeBox(node);
            if (!box) {
                console.log("No box for element");
                continue;
            }
            
            if (node.localName === "sp") {
                const spPr = Array.from(node.children).find((el) => el.localName === "spPr");
                const fill = getShapeFill(spPr);
                const geom = getShapeGeometry(spPr);
                const textData = extractTextFromShape(node);
                
                console.log("Shape found:", {
                    hasFill: !!fill,
                    hasText: !!textData,
                    textContent: textData ? textData.paragraphs.length + " paragraphs" : "none"
                });
                
                // THIS IS THE KEY LINE - keep shapes with text OR fill
                if (textData || fill) {
                    shapes.push({
                        type: "text", // Always "text" for sp elements (they can contain text)
                        box,
                        fill,
                        textData,
                        geom,
                        isMaster: false
                    });
                }
            } else if (node.localName === "pic") {
                const blipEl = Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "blip");
                const embed = blipEl?.getAttribute("r:embed") || 
                             blipEl?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
                
                if (!embed) continue;
                
                const target = rels[embed];
                if (!target) continue;
                
                const mediaPath = resolveMediaPath(slidePath, target);
                const mediaFile = zip.file(mediaPath);
                if (!mediaFile) continue;
                
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
                    // Skip images that fail to load
                }
            }
        }
        
        console.log(`Total shapes extracted from slide: ${shapes.length}`);
        return shapes;
    } catch (e) {
        console.error("Error in parseSlideShapes:", e);
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
    const slides = [];
    
    for (let i = 0; i < slidePaths.length; i++) {
        const slidePath = slidePaths[i];
        
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
            // Skip slides that fail to parse
        }
    }
    
    return slides;
}

function parseStyleProps(styleEl) {
    const props = {};
    if (!styleEl) return props;
    const paraProps = Array.from(styleEl.children).find((el) => el.localName === "paragraph-properties");
    if (paraProps) {
        const marL = paraProps.getAttribute("fo:margin-left") || paraProps.getAttribute("margin-left");
        const indent = paraProps.getAttribute("fo:text-indent") || paraProps.getAttribute("text-indent");
        if (marL) props.marL = marL;
        if (indent) props.indent = indent;
        const align = paraProps.getAttribute("fo:text-align") || paraProps.getAttribute("text-align");
        if (align) props.align = align;
    }

    const textProps = Array.from(styleEl.children).find((el) => el.localName === "text-properties");
    if (textProps) {
        const fontSize = textProps.getAttribute("fo:font-size") || textProps.getAttribute("font-size");
        if (fontSize) props.fontSize = fontSize;
        const fontWeight = textProps.getAttribute("fo:font-weight") || textProps.getAttribute("font-weight");
        if (fontWeight) props.fontWeight = fontWeight;
        const fontStyle = textProps.getAttribute("fo:font-style") || textProps.getAttribute("font-style");
        if (fontStyle) props.fontStyle = fontStyle;
        const color = textProps.getAttribute("fo:color") || textProps.getAttribute("color");
        if (color) props.color = color;
        const fontFamily = textProps.getAttribute("style:font-name") || textProps.getAttribute("font-name");
        if (fontFamily) props.fontFamily = fontFamily;
    }

    const graphicProps = Array.from(styleEl.children).find((el) => el.localName === "graphic-properties");
    if (graphicProps) {
        const fill = graphicProps.getAttribute("draw:fill");
        const fillColor = graphicProps.getAttribute("draw:fill-color");
        if (fill && fill !== "none") props.fill = fill;
        if (fillColor) props.fillColor = fillColor;
    }

    const pageProps = Array.from(styleEl.children).find((el) => el.localName === "drawing-page-properties");
    if (pageProps) {
        const fill = pageProps.getAttribute("draw:fill");
        const fillColor = pageProps.getAttribute("draw:fill-color");
        if (fill && fill !== "none") props.fill = fill;
        if (fillColor) props.fillColor = fillColor;
    }
    return props;
}

function parseOdpStyles(xmlText) {
    const listStyles = {};
    if (!xmlText) return { styles: {}, listStyles };
    const doc = parseXml(xmlText);
    if (!doc) return { styles: {}, listStyles };

    const rawStyles = {};
    const styles = {};

    const styleNodes = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "style");
    for (const style of styleNodes) {
        const name = style.getAttribute("style:name");
        const family = style.getAttribute("style:family");
        if (!name) continue;
        if (family === "text" || family === "paragraph" || family === "graphic" || family === "presentation") {
            rawStyles[name] = {
                props: parseStyleProps(style),
                parent: style.getAttribute("style:parent-style-name")
            };
        }
    }

    const resolveStyle = (name, seen = new Set()) => {
        if (!name) return {};
        if (styles[name]) return styles[name];
        const def = rawStyles[name];
        if (!def) return {};
        if (seen.has(name)) return def.props;
        seen.add(name);
        const parentProps = def.parent ? resolveStyle(def.parent, seen) : {};
        styles[name] = mergeStyles(parentProps, def.props);
        return styles[name];
    };

    Object.keys(rawStyles).forEach((key) => resolveStyle(key));

    const listNodes = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "list-style");
    for (const list of listNodes) {
        const name = list.getAttribute("style:name");
        if (!name) continue;
        const levels = {};
        const levelNodes = Array.from(list.children).filter((el) => el.localName.startsWith("list-level-style"));
        for (const lvl of levelNodes) {
            const level = parseInt(lvl.getAttribute("text:level") || "1", 10);
            const isNumber = lvl.localName === "list-level-style-number";
            const numFormat = lvl.getAttribute("style:num-format");
            const ch = lvl.getAttribute("text:bullet-char") || (isNumber ? "" : "•");
            const llProps = Array.from(lvl.children).find((el) => el.localName === "list-level-properties");
            const spaceBefore = llProps?.getAttribute("text:space-before") || llProps?.getAttribute("space-before") || "0";
            const minLabelWidth = llProps?.getAttribute("text:min-label-width") || llProps?.getAttribute("min-label-width") || "0";
            const start = parseInt(lvl.getAttribute("text:start-value") || "1", 10);
            const prefix = lvl.getAttribute("style:num-prefix") || "";
            const suffix = lvl.getAttribute("style:num-suffix") || (isNumber ? "." : "");
            const type = !numFormat && isNumber ? "none" : isNumber ? "number" : "char";
            levels[level] = {
                char: ch || "•",
                spaceBefore,
                minLabelWidth,
                type,
                start,
                prefix,
                suffix
            };
        }
        if (Object.keys(levels).length > 0) {
            listStyles[name] = levels;
        }
    }

    return { styles, listStyles };
}

function lengthToPx(val) {
    if (!val) return 0;
    const n = parseFloat(val);
    if (isNaN(n)) return 0;
    if (val.includes("mm")) return n * 3.7795275591;
    if (val.includes("cm")) return n * 37.795275591;
    if (val.includes("in")) return n * 96;
    if (val.includes("pt")) return n * (96 / 72);
    return n;
}

function getOdpStyle(allStyles, name) {
    if (!name) return {};
    return allStyles[name] || {};
}

function guessMimeFromBytes(path, bytes) {
    const ext = path.split(".").pop()?.toLowerCase();
    const mimeTypes = {
        png: "image/png",
        jpg: "image/jpeg",
        jpeg: "image/jpeg",
        gif: "image/gif",
        bmp: "image/bmp",
        svg: "image/svg+xml",
        webp: "image/webp",
        avif: "image/avif"
    };
    if (ext && mimeTypes[ext]) return mimeTypes[ext];
    if (bytes.length > 4) {
        if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
        if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
        if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return "image/gif";
        if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
        if (bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46) return "image/webp";
    }
    return "image/png";
}

function uint8ToBase64(bytes) {
    let binary = "";
    for (let i = 0; i < bytes.length; i += 1) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

async function loadOdpImage(zip, href) {
    if (!href) return null;
    const cleanHref = href.replace(/^\.\//, "").replace(/^\/+/g, "");
    const file = zip.file(cleanHref);
    if (!file) return null;
    const bytes = await file.async("uint8array");
    const mime = guessMimeFromBytes(cleanHref, bytes);
    const base64 = uint8ToBase64(bytes);
    return `data:${mime};base64,${base64}`;
}

async function frameToShapes(frame, allStyles, allListStyles, zip, options = {}) {
    const shapes = [];
    const isMaster = options.isMaster ?? false;
    const skipPlaceholders = options.skipPlaceholders ?? false;

    if (skipPlaceholders && frame.getAttribute("presentation:placeholder") === "true") {
        return shapes;
    }

    const presentationClass = frame.getAttribute("presentation:class") || "";
    if (isMaster && ["page-number", "footer", "header", "date-time", "notes"].includes(presentationClass)) {
        return shapes;
    }

    const x = lengthToPx(frame.getAttribute("svg:x") || frame.getAttribute("x"));
    const y = lengthToPx(frame.getAttribute("svg:y") || frame.getAttribute("y"));
    const width = lengthToPx(frame.getAttribute("svg:width") || frame.getAttribute("width"));
    const height = lengthToPx(frame.getAttribute("svg:height") || frame.getAttribute("height"));

    const styleName = frame.getAttribute("presentation:style-name") || frame.getAttribute("draw:style-name");
    const gStyle = getOdpStyle(allStyles, styleName);
    const frameTextStyleName = frame.getAttribute("draw:text-style-name");
    const frameTextStyle = getOdpStyle(allStyles, frameTextStyleName);
    const baseTextStyle = mergeStyles(gStyle, frameTextStyle);

    if (gStyle.fill && gStyle.fill !== "none" && gStyle.fillColor) {
        shapes.push({
            type: "shape",
            box: { x, y, cx: width || 400, cy: height || 200 },
            fill: { type: "solid", color: gStyle.fillColor },
            geom: null,
            textData: null,
            isMaster
        });
    }

    const tableEl = Array.from(frame.getElementsByTagName("*")).find((el) => el.localName === "table");
    if (tableEl) {
        const tableData = parseTableData(tableEl);
        if (tableData) {
            shapes.push({
                type: "table",
                box: { x, y, cx: width || 400, cy: height || 200 },
                data: tableData,
                isMaster
            });
            return shapes;
        }
    }

    const objectEl = Array.from(frame.getElementsByTagName("*")).find((el) => el.localName === "object");
    const imageEl = Array.from(frame.getElementsByTagName("*")).find((el) => el.localName === "image");
    const objectHref = objectEl ? (objectEl.getAttribute("xlink:href") || objectEl.getAttribute("href")) : null;
    if (objectHref) {
        const chartData = await parseEmbeddedObjectChart(zip, objectHref);
        if (chartData) {
            shapes.push({
                type: "chart",
                box: { x, y, cx: width || 400, cy: height || 200 },
                data: chartData,
                isMaster
            });
            return shapes;
        }
    }

    const imageHref = imageEl ? (imageEl.getAttribute("xlink:href") || imageEl.getAttribute("href")) : null;
    const replacementHref = objectHref;
    const preferredImages = [imageHref, replacementHref].filter(Boolean);
    for (const href of preferredImages) {
        const src = await loadOdpImage(zip, href);
        if (src) {
            shapes.push({
                type: "image",
                box: { x, y, cx: width || 400, cy: height || 200 },
                fill: null,
                geom: null,
                src,
                isMaster
            });
            return shapes;
        }
    }

    const textBox = Array.from(frame.children).find((el) => el.localName === "text-box" || el.localName === "textbox") || frame;
    const paragraphs = [];
    const listCounters = {};

    const nextListIndex = (listStyleName, level, startAt) => {
        if (!listStyleName) return null;
        if (!listCounters[listStyleName]) listCounters[listStyleName] = {};
        if (listCounters[listStyleName][level] == null) {
            listCounters[listStyleName][level] = startAt;
        } else {
            listCounters[listStyleName][level] += 1;
        }
        return listCounters[listStyleName][level];
    };

    function collectParas(node, level = 0, listStyleName = null, listIndex = null) {
        if (node.localName === "list") {
            const styleName = node.getAttribute("text:style-name") || listStyleName;
            const lvl = level + 1;
            const lvlDef = styleName ? (allListStyles[styleName]?.[lvl] || allListStyles[styleName]?.[1]) : null;
            const startAt = lvlDef?.start || 1;

            if (styleName && level === 0 && node.getAttribute("text:continue-numbering") !== "true") {
                listCounters[styleName] = {};
            }

            const header = Array.from(node.children).find((el) => el.localName === "list-header");
            if (header) {
                const headerIdx = styleName ? nextListIndex(styleName, lvl, startAt) : null;
                Array.from(header.children).forEach((child) => collectParas(child, lvl, styleName, headerIdx));
            }

            const items = Array.from(node.children).filter((el) => el.localName === "list-item");
            for (const item of items) {
                const idx = styleName ? nextListIndex(styleName, lvl, startAt) : null;
                collectParas(item, lvl, styleName, idx);
            }
            return;
        }

        if (node.localName === "list-item") {
            const children = Array.from(node.children);
            for (const child of children) {
                collectParas(child, level, listStyleName, listIndex);
            }
            return;
        }

        if (node.localName === "p") {
            const pStyleName = node.getAttribute("text:style-name");
            const pStyle = getOdpStyle(allStyles, pStyleName);
            const spans = Array.from(node.childNodes)
                .filter((n) => n.nodeType === 3 || (n.nodeType === 1 && (n.localName === "span" || n.localName === "s")))
                .map((childNode) => {
                    const text = childNode.localName === "s"
                        ? " ".repeat(parseInt(childNode.getAttribute("text:c") || "1", 10) || 1)
                        : childNode.textContent || "";
                    const spanStyleName = childNode.nodeType === 1 ? childNode.getAttribute("text:style-name") : null;
                    const spanStyle = spanStyleName ? getOdpStyle(allStyles, spanStyleName) : {};
                    const combined = mergeStyles(baseTextStyle, mergeStyles(pStyle, spanStyle));
                    return { text, style: combined };
                })
                .filter((s) => s.text.trim().length > 0);

            const effectiveListStyle = presentationClass === "title" ? null : listStyleName;
            const levels = effectiveListStyle ? allListStyles[effectiveListStyle] || {} : {};
            const lvlDef = effectiveListStyle ? levels[level] || levels[1] || {} : {};

            let bullet = null;
            if (effectiveListStyle && lvlDef.type && lvlDef.type !== "none") {
                if (lvlDef.type === "number") {
                    const idx = listIndex ?? lvlDef.start ?? 1;
                    bullet = { type: "auto", index: idx, level: Math.max(level - 1, 0), prefix: lvlDef.prefix || "", suffix: lvlDef.suffix || "." };
                } else if (lvlDef.type === "char" && lvlDef.char) {
                    const bulletChar = lvlDef.char;
                    bullet = { type: "char", char: bulletChar, level: Math.max(level - 1, 0) };
                }
            }

            const spaceBefore = effectiveListStyle && (lvlDef.spaceBefore || levels[1]?.spaceBefore);
            const minLabelWidth = effectiveListStyle && (lvlDef.minLabelWidth || levels[1]?.minLabelWidth);
            const indentPx = lengthToPx(spaceBefore || "0") + lengthToPx(minLabelWidth || "0");

            const marL = pStyle.marL ? lengthToPx(pStyle.marL) : 0;
            const textIndent = pStyle.indent ? lengthToPx(pStyle.indent) : 0;

            const align = pStyle.align || baseTextStyle.align || "left";

            let fontSize = baseTextStyle.fontSize || pStyle.fontSize;
            const runsWithStyle = spans.map((s) => {
                if (s.style.fontSize) {
                    fontSize = s.style.fontSize;
                } else if (fontSize) {
                    s.style.fontSize = fontSize;
                }
                if (!s.style.color && (pStyle.color || baseTextStyle.color)) s.style.color = pStyle.color || baseTextStyle.color;
                if (!s.style.fontWeight && (pStyle.fontWeight || baseTextStyle.fontWeight)) s.style.fontWeight = pStyle.fontWeight || baseTextStyle.fontWeight;
                if (!s.style.fontStyle && (pStyle.fontStyle || baseTextStyle.fontStyle)) s.style.fontStyle = pStyle.fontStyle || baseTextStyle.fontStyle;
                if (!s.style.fontFamily && (pStyle.fontFamily || baseTextStyle.fontFamily)) s.style.fontFamily = pStyle.fontFamily || baseTextStyle.fontFamily;
                return s;
            });

            if (!fontSize) {
                const isTitle = presentationClass === "title";
                fontSize = isTitle ? "44pt" : "18pt";
                runsWithStyle.forEach((s) => (s.style.fontSize = s.style.fontSize || fontSize));
            }

            if (runsWithStyle.length > 0) {
                paragraphs.push({
                    align,
                    runs: runsWithStyle,
                    bullet,
                    level: Math.max(level - 1, 0),
                    marL: marL + indentPx,
                    indent: textIndent
                });
            }
        }
    }

    Array.from(textBox.children).forEach((child) => collectParas(child, 0, null, null));

    if (paragraphs.length > 0) {
        shapes.push({
            type: "text",
            box: { x, y, cx: width || 400, cy: height || 200 },
            fill: null,
            geom: null,
            textData: {
                paragraphs,
                verticalAlign: "flex-start"
            },
            isMaster
        });
    }

    return shapes;
}

async function parseOdpMasterPages(zip, stylesXml, allStyles, allListStyles) {
    const masters = {};
    if (!stylesXml) return masters;
    const doc = parseXml(stylesXml);
    if (!doc) return masters;
    const masterNodes = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "master-page");
    for (const master of masterNodes) {
        const name = master.getAttribute("style:name");
        if (!name) continue;
        const styleName = master.getAttribute("draw:style-name");
        const masterStyle = getOdpStyle(allStyles, styleName);
        const backgroundColor = masterStyle.fill && masterStyle.fillColor ? masterStyle.fillColor : null;
        const frameNodes = Array.from(master.getElementsByTagName("*")).filter((el) => el.localName === "frame");
        const shapes = [];
        for (const frame of frameNodes) {
            const frameShapes = await frameToShapes(frame, allStyles, allListStyles, zip, { isMaster: true, skipPlaceholders: true });
            shapes.push(...frameShapes);
        }
        masters[name] = {
            shapes,
            background: backgroundColor,
            pageLayout: master.getAttribute("style:page-layout-name") || null
        };
    }
    return masters;
}

function parseOdpPageLayouts(stylesXml) {
    const layouts = {};
    if (!stylesXml) return layouts;
    const doc = parseXml(stylesXml);
    if (!doc) return layouts;
    const layoutNodes = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "page-layout");
    for (const node of layoutNodes) {
        const name = node.getAttribute("style:name");
        if (!name) continue;
        const props = Array.from(node.children).find((el) => el.localName === "page-layout-properties");
        const w = props?.getAttribute("fo:page-width");
        const h = props?.getAttribute("fo:page-height");
        if (w && h) {
            layouts[name] = {
                cx: Math.round(lengthToPx(w)),
                cy: Math.round(lengthToPx(h))
            };
        }
    }
    return layouts;
}

function parseTableData(tableEl) {
    if (!tableEl) return null;
    const rows = Array.from(tableEl.getElementsByTagName("*")).filter((el) => el.localName === "table-row");
    if (!rows.length) return null;
    const data = rows.map((row) => {
        const cells = Array.from(row.children).filter((el) => el.localName === "table-cell");
        return cells.map((cell) => (cell.textContent || "").trim());
    });
    return data.length ? data : null;
}

async function parseEmbeddedObjectChart(zip, href) {
    if (!href) return null;
    const cleanHref = href.replace(/^\.\//, "").replace(/^\/+/g, "");
    const contentPath = `${cleanHref.replace(/\\+/g, "/").replace(/\/+$/, "")}/content.xml`;
    const contentFile = zip.file(contentPath);
    if (!contentFile) return null;
    const xml = await contentFile.async("text");
    const doc = parseXml(xml);
    if (!doc) return null;
    const tableEl = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "table");
    const tableData = parseTableData(tableEl);
    if (!tableData || tableData.length < 2) return null;

    const headers = tableData[0].slice(1);
    const categories = tableData.slice(1).map((row) => row[0] || "");
    const values = headers.map((_, colIdx) => tableData.slice(1).map((row) => parseFloat(row[colIdx + 1]) || 0));
    return { headers, categories, values };
}

async function renderOdpSlides(base64) {
    const buffer = decodeBase64ToUint8(base64);
    const zip = await JSZip.loadAsync(buffer);
    const contentXml = await zip.file("content.xml")?.async("text");
    if (!contentXml) return [];

    const doc = parseXml(contentXml);
    if (!doc) return [];

    const stylesXml = await zip.file("styles.xml")?.async("text");
    const { styles: globalStyles, listStyles: globalListStyles } = parseOdpStyles(stylesXml);
    const pageLayouts = parseOdpPageLayouts(stylesXml);

    const autoStylesNode = doc.querySelector("office\\:automatic-styles,automatic-styles");
    const { styles: autoStyles, listStyles: autoListStyles } = autoStylesNode
        ? parseOdpStyles(autoStylesNode.outerHTML)
        : { styles: {}, listStyles: {} };

    const allStyles = { ...globalStyles, ...autoStyles };
    const allListStyles = { ...globalListStyles, ...autoListStyles };
    const masterPages = await parseOdpMasterPages(zip, stylesXml, allStyles, allListStyles);

    const pages = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "page");
    const slides = [];

    for (const page of pages.slice(0, MAX_SLIDES)) {
        const wAttr = page.getAttribute("svg:width") || page.getAttribute("width");
        const hAttr = page.getAttribute("svg:height") || page.getAttribute("height");
        let size = null;

        if (wAttr && hAttr) {
            size = {
                cx: Math.round(lengthToPx(wAttr)),
                cy: Math.round(lengthToPx(hAttr))
            };
        }

        const frames = Array.from(page.getElementsByTagName("*")).filter((el) => el.localName === "frame");
        const shapes = [];

        const masterName = page.getAttribute("draw:master-page-name");
        const master = masterName ? masterPages[masterName] : undefined;
        if (master?.shapes?.length) {
            shapes.push(...master.shapes);
        }

        if (!size) {
            const layoutName = master?.pageLayout;
            if (layoutName && pageLayouts[layoutName]) {
                size = pageLayouts[layoutName];
            }
        }

        if (!size) {
            size = { cx: 960, cy: 540 };
        }

        for (const frame of frames) {
            const frameShapes = await frameToShapes(frame, allStyles, allListStyles, zip, { isMaster: false, skipPlaceholders: false });
            shapes.push(...frameShapes);
        }

        const pageStyleName = page.getAttribute("draw:style-name");
        const pageStyle = getOdpStyle(allStyles, pageStyleName);
        let backgroundColor = pageStyle.fillColor || null;
        if (!backgroundColor && master?.background) {
            backgroundColor = master.background;
        }

        slides.push({
            path: "",
            size,
            shapes,
            background: backgroundColor ? { color: backgroundColor } : null
        });
    }

    return slides;
}

// Attempt to render legacy PPT (binary) files via minimal CFB parsing.
async function renderPptSlides(base64) {
    try {
        const buffer = decodeBase64ToUint8(base64);
        
        // Parse CFB (Compound File Binary) format
        const cfb = CFB.read(buffer, { type: "array" });
        const pictures = extractPptPictures(cfb);
        
        // Find PowerPoint Document stream
        const pptStream = findCfbStream(cfb, "PowerPoint Document");
        if (!pptStream) {
            throw new Error("Not a valid PowerPoint file - PowerPoint Document stream not found");
        }
        
        // Ensure we have a proper Uint8Array
        const streamArray = pptStream instanceof Uint8Array ? pptStream : new Uint8Array(pptStream);
        
        // Parse the PowerPoint binary format
        const slides = parsePptStream(streamArray, pictures);

        // If we have pictures but no explicit shapes using them, set first picture as background
        if (pictures.length > 0) {
            const bg = pictures[0].dataUrl;
            slides.forEach((slide) => {
                slide.shapes.unshift({
                    type: "image",
                    box: { x: 0, y: 0, cx: slide.size.cx, cy: slide.size.cy },
                    src: bg,
                    isMaster: false
                });
            });
        }
        
        return slides;
    } catch (error) {
        console.error("Error in renderPptSlides:", error);
        throw error;
    }
}

function findCfbStream(cfb, name) {
    for (const entry of cfb.FileIndex) {
        if (entry.name === name && entry.content) {
            return entry.content;
        }
    }
    return null;
}

function extractPptPictures(cfb) {
    const pics = [];
    for (const entry of cfb.FileIndex) {
        const name = entry.name || "";
        const lower = name.toLowerCase();
        if (!entry.content) continue;
        if (!lower.includes("pictures") && !lower.endsWith(".png") && !lower.endsWith(".jpg") && !lower.endsWith(".jpeg") && !lower.endsWith(".gif") && !lower.endsWith(".bmp")) {
            continue;
        }
        const bytes = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content);
        if (lower.includes("pictures")) {
            const found = extractImagesFromBlob(bytes, name);
            pics.push(...found);
        } else {
            const mime = guessMimeFromBytes(name, bytes);
            const dataUrl = `data:${mime};base64,${uint8ToBase64(bytes)}`;
            pics.push({ name, dataUrl });
        }
    }
    return pics;
}

function extractImagesFromBlob(bytes, label = "Pictures") {
    const results = [];
    const len = bytes.length;
    const matches = [];

    const sigs = [
        { name: "jpeg", sig: [0xff, 0xd8, 0xff], end: [0xff, 0xd9], mime: "image/jpeg" },
        { name: "png", sig: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], end: [0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82], mime: "image/png" },
        { name: "gif", sig: [0x47, 0x49, 0x46, 0x38], end: [0x00, 0x3b], mime: "image/gif" }
    ];

    for (let i = 0; i < len; i++) {
        for (const sig of sigs) {
            const s = sig.sig;
            let match = true;
            for (let k = 0; k < s.length && i + k < len; k++) {
                if (bytes[i + k] !== s[k]) {
                    match = false;
                    break;
                }
            }
            if (!match) continue;
            // find end
            let endIdx = -1;
            const endSig = sig.end;
            for (let j = i + s.length; j < len - endSig.length; j++) {
                let ok = true;
                for (let k = 0; k < endSig.length; k++) {
                    if (bytes[j + k] !== endSig[k]) {
                        ok = false;
                        break;
                    }
                }
                if (ok) {
                    endIdx = j + endSig.length;
                    break;
                }
            }
            if (endIdx === -1) continue;
            matches.push({ start: i, end: endIdx, mime: sig.mime });
        }
    }

    matches.sort((a, b) => a.start - b.start);
    for (let idx = 0; idx < matches.length; idx++) {
        const { start, end, mime } = matches[idx];
        const slice = bytes.slice(start, end);
        results.push({
            name: `${label}-${idx}`,
            dataUrl: `data:${mime};base64,${uint8ToBase64(slice)}`
        });
    }
    return results;
}

function buildPptTextShapesFromList(texts, slideWidthPx, slideHeightPx) {
    if (!texts || texts.length === 0) return [];

    const shapes = [];
    const splitTitleBody = (allTexts) => {
        // Prefer explicit separation (multiple text runs)
        if (allTexts.length > 1) {
            return { title: allTexts[0], body: allTexts.slice(1).join("\n") };
        }
        const raw = allTexts[0] || "";
        // Try blank-line or newline split
        const parts = raw.split(/\r?\n\r?\n/).map((p) => p.trim()).filter(Boolean);
        if (parts.length > 1) {
            return { title: parts[0], body: parts.slice(1).join("\n\n") };
        }
        const lines = raw.split(/\r?\n+/).map((l) => l.trim()).filter(Boolean);
        if (lines.length > 1) {
            return { title: lines[0], body: lines.slice(1).join("\n") };
        }
        // Fallback: first sentence or first 10 words as title
        const sentenceMatch = raw.match(/^(.+?[.!?])\s+(.*)$/);
        if (sentenceMatch) {
            return { title: sentenceMatch[1].trim(), body: sentenceMatch[2].trim() };
        }
        const words = raw.split(/\s+/);
        const titleWords = words.slice(0, Math.min(10, words.length)).join(" ");
        const bodyWords = words.slice(Math.min(10, words.length)).join(" ");
        return { title: titleWords, body: bodyWords };
    };

    const { title: titleText, body: bodyText } = splitTitleBody(texts);

    // Title - larger font, positioned higher
    shapes.push({
        type: "text",
        box: {
            x: Math.round(slideWidthPx * 0.05),
            y: Math.round(slideHeightPx * 0.08),
            cx: Math.round(slideWidthPx * 0.9),
            cy: Math.round(slideHeightPx * 0.14)
        },
        textData: {
            paragraphs: [{
                align: "left",
                runs: [{ 
                    text: titleText, 
                    style: { 
                        fontSize: "42pt",
                        fontWeight: "bold",
                        color: "#000000" 
                    } 
                }],
                level: 0,
                marL: 0,
                indent: 0
            }],
            verticalAlign: "flex-start"
        },
        isMaster: false
    });

    // Body text - smaller font, positioned lower
    const bodyChunks = Array.isArray(bodyText) ? bodyText : [bodyText];
    const bodyTextNormalized = bodyChunks
        .map((b) => (typeof b === "string" ? b : ""))
        .join("\n");

    if (bodyTextNormalized && bodyTextNormalized.trim().length > 0) {
        const paragraphs = bodyTextNormalized
            .split(/\r?\n+/)
            .map((p) => p.trim())
            .filter(Boolean)
            .map((p) => ({
                align: "left",
                runs: [{ 
                    text: p, 
                    style: { 
                        fontSize: "16pt",
                        fontWeight: "normal",
                        color: "#000000" 
                    } 
                }],
                level: 0,
                marL: 0,
                indent: 0
            }));

        if (paragraphs.length) {
            shapes.push({
                type: "text",
                box: {
                    x: Math.round(slideWidthPx * 0.07),
                    y: Math.round(slideHeightPx * 0.20),
                    cx: Math.round(slideWidthPx * 0.86),
                    cy: Math.round(slideHeightPx * 0.6)
                },
                textData: {
                    paragraphs,
                    verticalAlign: "flex-start"
                },
                isMaster: false
            });
        }
    }

    return shapes;
}

function parsePptStream(stream, pictures = []) {
    const slides = [];
    let offset = 0;
    let slideWidth = 9144000;
    let slideHeight = 6858000;
    const textBySlide = new Map();
    const shapesBySlide = new Map();
    
    console.log("Parsing PPT, stream length:", stream.length);
    
    // Parse records
    while (offset < stream.length - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        if (recordEnd > stream.length) {
            console.warn(`Record extends beyond stream, stopping`);
            break;
        }
        
        try {
            switch (header.recType) {
                case 0x03E8: // RT_Document
                    const docResult = parsePptDocument(stream, offset, recordEnd);
                    if (docResult.width) slideWidth = docResult.width;
                    if (docResult.height) slideHeight = docResult.height;
                    break;
                
                case 0x03EE: // RT_Slide
                    try {
                        const slideWidthPx = Math.max(1, Math.round(slideWidth / 9525));
                        const slideHeightPx = Math.max(1, Math.round(slideHeight / 9525));
                        const slideShapes = parsePptSlide(stream, offset, recordEnd, slideWidthPx, slideHeightPx);
                        shapesBySlide.set(slides.length, slideShapes);
                        slides.push({
                            path: `slide${slides.length}`,
                            size: { cx: slideWidthPx, cy: slideHeightPx },
                            shapes: slideShapes
                        });
                    } catch (slideError) {
                        console.error("Error parsing slide:", slideError);
                        const slideWidthPx = Math.max(1, Math.round(slideWidth / 9525));
                        const slideHeightPx = Math.max(1, Math.round(slideHeight / 9525));
                        slides.push({
                            path: `slide${slides.length}`,
                            size: { cx: slideWidthPx, cy: slideHeightPx },
                            shapes: []
                        });
                    }
                    break;
                
                case 0x0FF0: // RT_SlideListWithText
                    const slideText = parsePptSlideListWithText(stream, offset, recordEnd);
                    if (slideText.length > 0) {
                        textBySlide.set(textBySlide.size, slideText);
                    }
                    break;
            }
        } catch (e) {
            console.error(`Error parsing record type 0x${header.recType.toString(16)}:`, e);
        }
        
        offset = recordEnd;
    }
    
    console.log(`Parsed ${slides.length} slides with shapes:`, slides.map(s => s.shapes.length));
    
    // Merge text into slides: gather text from records and any parsed text shapes,
    // then replace text shapes with a consistent title/body layout.
    slides.forEach((slide, index) => {
        const recordTexts = textBySlide.get(index) || [];
        const shapeTexts = slide.shapes
            .filter((s) => s.type === "text" && s.textData?.paragraphs?.length)
            .flatMap((s) =>
                s.textData.paragraphs
                    .map((p) => p.runs.map((r) => r.text || "").join(""))
                    .filter((t) => t && t.trim().length > 0)
            );
        const allTexts = [...recordTexts, ...shapeTexts].filter((t) => t && t.trim().length > 0);
        if (allTexts.length === 0) return;

        const slideWidthPx = slide.size?.cx || 960;
        const slideHeightPx = slide.size?.cy || 540;
        const nonTextShapes = slide.shapes.filter((s) => s.type !== "text");
        const generated = buildPptTextShapesFromList(allTexts, slideWidthPx, slideHeightPx);
        slide.shapes = [...nonTextShapes, ...generated];
    });
    
    return slides;
}

function readPptRecordHeader(stream, offset) {
    if (offset + 8 > stream.length) {
        throw new Error("Unexpected end of stream");
    }
    
    // Create ArrayBuffer from the slice we need
    const slice = stream.slice(offset, offset + 8);
    const view = new DataView(slice.buffer, slice.byteOffset, 8);
    
    const verAndInstance = view.getUint16(0, true);
    const recVer = verAndInstance & 0x0F;
    const recInstance = (verAndInstance >> 4) & 0x0FFF;
    const recType = view.getUint16(2, true);
    const recLen = view.getUint32(4, true);
    
    return { recVer, recInstance, recType, recLen };
}

function parsePptDocument(stream, offset, endOffset) {
    let width = null;
    let height = null;
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        // Look for environment container (0x03F2)
        if (header.recType === 0x03F2) {
            const envResult = parsePptEnvironment(stream, offset, recordEnd);
            if (envResult.width) width = envResult.width;
            if (envResult.height) height = envResult.height;
        }
        
        offset = recordEnd;
    }
    
    return { width, height };
}

function parsePptEnvironment(stream, offset, endOffset) {
    let width = null;
    let height = null;
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        // Slide size atom (0x03F4)
        if (header.recType === 0x03F4 && header.recLen >= 8) {
            const slice = stream.slice(offset, offset + 8);
            const view = new DataView(slice.buffer, slice.byteOffset, 8);
            width = view.getInt32(0, true);
            height = view.getInt32(4, true);
        }
        
        offset = recordEnd;
    }
    
    return { width, height };
}

function parseProgTags(stream, offset, endOffset) {
    const texts = [];
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        // Safety check
        if (recordEnd > endOffset || recordEnd > stream.length) {
            console.warn("ProgTags record extends beyond container, skipping");
            break;
        }
        
        // Look for text atoms in ProgTags - 0x0FA0 (TextCharsAtom) and 0x0FA8 (TextBytesAtom)
        if (header.recType === 0x0FA0) {
            const text = readPptTextChars(stream, offset, header.recLen);
            console.log("    Found TextCharsAtom:", text.substring(0, 50));
            if (text.trim()) texts.push(text);
        } else if (header.recType === 0x0FA8) {
            const text = readPptTextBytes(stream, offset, header.recLen);
            console.log("    Found TextBytesAtom:", text.substring(0, 50));
            if (text.trim()) texts.push(text);
        }
        // Recurse into container records (recVer & 0xF === 0xF means it's a container)
        else if (header.recVer === 0xF && header.recLen > 0) {
            console.log(`    Recursing into container 0x${header.recType.toString(16)}`);
            const nested = parseProgTags(stream, offset, recordEnd);
            texts.push(...nested);
        }
        
        offset = recordEnd;
    }
    
    console.log(`  ProgTags extracted ${texts.length} text fragments`);
    return texts;
}

function parsePptSlide(stream, offset, endOffset, slideWidthPx, slideHeightPx) {
    const shapes = [];
    const collectedTexts = [];
    
    console.log(`  Parsing slide from ${offset} to ${endOffset}`);
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        console.log(`    Slide sub-record: type=0x${header.recType.toString(16)}, len=${header.recLen}`);
        
        // PPDrawing container (0x040C)
        if (header.recType === 0x040C) {
            console.log("    Found PPDrawing");
            const drawingShapes = parsePptDrawing(stream, offset, recordEnd, slideWidthPx, slideHeightPx);
            console.log("    Drawing shapes:", drawingShapes);
            shapes.push(...drawingShapes);
        }
        // ProgTags container (0x1388) - may contain text
        else if (header.recType === 0x1388) {
            console.log("    Found ProgTags - parsing for text");
            const texts = parseProgTags(stream, offset, recordEnd);
            console.log("    ProgTags texts:", texts);
            if (texts.length > 0) {
                collectedTexts.push(...texts);
            }
        }
        // SlideListWithText container (0x0FF0) - often holds body text
        else if (header.recType === 0x0FF0) {
            console.log("    Found SlideListWithText - parsing for text");
            const texts = parsePptSlideListWithText(stream, offset, recordEnd);
            console.log("    SlideListWithText texts:", texts);
            if (texts.length > 0) {
                collectedTexts.push(...texts);
            }
        }
        // ALSO check for TextHeaderAtom (0x0F9F) and text containers
        else if (header.recType === 0x0F9F) {
            console.log("    Found TextHeaderAtom");
        }
        
        offset = recordEnd;
    }
    
    // Replace any text shapes with a consistent layout using collected text and any parsed text shapes
    const existingText = shapes
        .filter((s) => s.type === "text" && s.textData?.paragraphs?.length)
        .flatMap((s) =>
            s.textData.paragraphs
                .flatMap((p) => p.runs.map((r) => r.text || "").join(""))
                .filter((t) => t && t.trim().length > 0)
        );
    const allTexts = [...collectedTexts, ...existingText].filter((t) => t && t.trim().length > 0);
    if (allTexts.length > 0) {
        const nonText = shapes.filter((s) => s.type !== "text");
        const generated = buildPptTextShapesFromList(allTexts, slideWidthPx, slideHeightPx);
        shapes.length = 0;
        shapes.push(...nonText, ...generated);
    }
    
    console.log(`  Total shapes extracted: ${shapes.length}`);
    return shapes;
}

function parsePptDrawing(stream, offset, endOffset, slideWidth, slideHeight) {
    const shapes = [];
    
    console.log(`      Parsing drawing from ${offset} to ${endOffset}`);
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        console.log(`        Drawing record: type=0x${header.recType.toString(16)}, len=${header.recLen}`);
        
        // OfficeArt containers: 0xF002 (SpgrContainer), 0xF003 (group), 0xF004 (shape)
        if (header.recType === 0xF002 || header.recType === 0xF003 || header.recType === 0xF004) {
            console.log(`        Found shape container 0x${header.recType.toString(16)}`);
            const containerShapes = parsePptShapeContainer(stream, offset, recordEnd, slideWidth, slideHeight);
            console.log("        Container shapes:", containerShapes);
            shapes.push(...containerShapes);
        }
        // Also check for direct OfficeArt containers
        else if (header.recType === 0xF000) { // OfficeArtDggContainer
            console.log("        Found OfficeArtDggContainer");
        }
        
        offset = recordEnd;
    }
    
    console.log(`      Drawing extracted ${shapes.length} shapes`);
    return shapes;
}

function parsePptShapeProperties(stream, offset, length) {
    const props = { bounds: {} };
    const numProps = Math.floor(length / 6);
    
    for (let i = 0; i < numProps && offset + i * 6 + 6 <= stream.length; i++) {
        const propOffset = offset + i * 6;
        const slice = stream.slice(propOffset, propOffset + 6);
        const view = new DataView(slice.buffer, slice.byteOffset, 6);
        const propId = view.getUint16(0, true);
        const propValue = view.getUint32(2, true);
        
        // Shape bounds
        if (propId === 0x0004) props.bounds.left = propValue;
        else if (propId === 0x0005) props.bounds.top = propValue;
        else if (propId === 0x0006) props.bounds.right = propValue;
        else if (propId === 0x0007) props.bounds.bottom = propValue;
        // Fill color
        else if (propId === 0x0181) {
            props.fillColor = pptColorFromInt(propValue);
        }
        // Fill type
        else if (propId === 0x0180) {
            props.fillType = propValue;
        }
    }
    
    return props;
}

function parsePptShapeContainer(stream, offset, endOffset, slideWidth, slideHeight) {
    const shapes = [];
    let shapeData = {};  // Remove const - we need to reset this
    
    console.log(`        Parsing shape container from ${offset} to ${endOffset}`);
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        if (recordEnd > endOffset) break;
        
        console.log(`          Shape record: type=0x${header.recType.toString(16)}, len=${header.recLen}`);
        
        try {
            // OfficeArtFOPT - shape formatting (0xF00B)
            if (header.recType === 0xF00B) {
                const props = parsePptShapeProperties(stream, offset, header.recLen);
                Object.assign(shapeData, props);
                console.log("          Shape properties:", props);
            }
            // OfficeArtClientTextbox (0xF00D)
            else if (header.recType === 0xF00D) {
                const text = parsePptClientTextbox(stream, offset, recordEnd);
                if (text) {
                    console.log("          Found text:", text.substring(0, 50));
                    shapeData.text = text;
                }
            }
            // OfficeArtClientData (0xF011)
            else if (header.recType === 0xF011) {
                console.log("          Found ClientData (possibly image)");
            }
            // Nested shape containers (0xF002, 0xF003, 0xF004)
            else if (header.recType === 0xF002 || header.recType === 0xF003 || header.recType === 0xF004) {
                // CRITICAL FIX: Before recursing, save current shape if it has data
                if (shapeData.bounds || shapeData.text) {
                    console.log(`        Creating shape before recursion:`, {
                        hasBounds: !!shapeData.bounds,
                        hasText: !!shapeData.text,
                        text: shapeData.text ? shapeData.text.substring(0, 30) : null
                    });
                    const shape = createPptShape(shapeData, slideWidth, slideHeight);
                    if (shape) {
                        shapes.push(shape);
                    }
                    // Reset shapeData for next shape
                    shapeData = {};
                }
                
                // Now recurse into nested container
                const nested = parsePptShapeContainer(stream, offset, recordEnd, slideWidth, slideHeight);
                shapes.push(...nested);
            }
        } catch (e) {
            console.error(`Error in shape container:`, e);
        }
        
        offset = recordEnd;
    }
    
    // Create shape if we have remaining data at the end
    if (shapeData.bounds || shapeData.text) {
        console.log(`        Creating final shape:`, {
            hasBounds: !!shapeData.bounds,
            hasText: !!shapeData.text,
            text: shapeData.text ? shapeData.text.substring(0, 30) : null
        });
        const shape = createPptShape(shapeData, slideWidth, slideHeight);
        if (shape) {
            shapes.push(shape);
        }
    }
    
    return shapes;
}

function parsePptClientData(stream, offset, endOffset) {
    // This is a placeholder for now (it's pretty complex to extract images properly from PPT [so RE is needed, probably, definitely])
    return null;
}


function parsePptShapeProperties(stream, offset, length) {
    const props = { bounds: {} };
    const numProps = Math.floor(length / 6);
    
    for (let i = 0; i < numProps && offset + i * 6 + 6 <= stream.length; i++) {
        const propOffset = offset + i * 6;
        const view = new DataView(stream.buffer, propOffset, 6);
        const propId = view.getUint16(0, true);
        const propValue = view.getUint32(2, true);
        
        // Shape bounds
        if (propId === 0x0004) props.bounds.left = propValue;
        else if (propId === 0x0005) props.bounds.top = propValue;
        else if (propId === 0x0006) props.bounds.right = propValue;
        else if (propId === 0x0007) props.bounds.bottom = propValue;
        // Fill color
        else if (propId === 0x0181) {
            props.fillColor = pptColorFromInt(propValue);
        }
        // Fill type
        else if (propId === 0x0180) {
            props.fillType = propValue;
        }
    }
    
    return props;
}

function parsePptClientTextbox(stream, offset, endOffset) {
    const texts = [];
    
    console.log(`      Parsing ClientTextbox from ${offset} to ${endOffset}`);
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        if (recordEnd > endOffset) break;
        
        console.log(`        Textbox record: type=0x${header.recType.toString(16)}, len=${header.recLen}`);
        
        // TextCharsAtom (0x0FA0) - Unicode text
        if (header.recType === 0x0FA0) {
            const text = readPptTextChars(stream, offset, header.recLen);
            console.log("        Found Unicode text:", text.substring(0, 50));
            if (text.trim()) texts.push(text);
        }
        // TextBytesAtom (0x0FA8) - ANSI text
        else if (header.recType === 0x0FA8) {
            const text = readPptTextBytes(stream, offset, header.recLen);
            console.log("        Found ANSI text:", text.substring(0, 50));
            if (text.trim()) texts.push(text);
        }
        // Recurse into containers
        else if (header.recVer === 0xF && header.recLen > 0) {
            console.log(`        Recursing into textbox container 0x${header.recType.toString(16)}`);
            const nested = parsePptClientTextbox(stream, offset, recordEnd);
            if (nested) texts.push(nested);
        }
        
        offset = recordEnd;
    }
    
    const result = texts.length > 0 ? texts.join(" ") : null;
    console.log(`      ClientTextbox result: ${result ? result.substring(0, 50) : "null"}`);
    return result;
}

function parsePptSlideListWithText(stream, offset, endOffset) {
    const texts = [];
    
    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;
        
        // TextCharsAtom
        if (header.recType === 0x0FA0) {
            texts.push(readPptTextChars(stream, offset, header.recLen));
        }
        // TextBytesAtom
        else if (header.recType === 0x0FA8) {
            texts.push(readPptTextBytes(stream, offset, header.recLen));
        }
        
        offset = recordEnd;
    }
    
    return texts;
}

function readPptTextChars(stream, offset, length) {
    const chars = [];
    for (let i = 0; i < length && offset + i + 1 < stream.length; i += 2) {
        const charCode = stream[offset + i] | (stream[offset + i + 1] << 8);
        if (charCode !== 0) chars.push(String.fromCharCode(charCode));
    }
    return chars.join("");
}

function readPptTextBytes(stream, offset, length) {
    const decoder = new TextDecoder("windows-1252");
    const bytes = stream.slice(offset, offset + length);
    return decoder.decode(bytes);
}

function pptColorFromInt(colorInt) {
    const r = colorInt & 0xFF;
    const g = (colorInt >> 8) & 0xFF;
    const b = (colorInt >> 16) & 0xFF;
    return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
}

function createPptShape(data, slideWidth, slideHeight) {
    if (!data.bounds && !data.text) return null;
    
    const emusToPixels = (emus) => Math.round(emus / 9525);
    
    // If we have bounds, use them
    if (data.bounds && data.bounds.left !== undefined) {
        const { left = 0, top = 0, right = slideWidth * 9525, bottom = slideHeight * 9525 } = data.bounds;
        const shape = {
            type: data.text ? "text" : "shape",
            box: {
                x: emusToPixels(left),
                y: emusToPixels(top),
                cx: emusToPixels(right - left),
                cy: emusToPixels(bottom - top)
            },
            isMaster: false
        };
        
        if (data.fillColor && data.fillType !== 0) {
            shape.fill = { type: "solid", color: data.fillColor };
        }
        
        if (data.text) {
            // Determine font size based on box position (higher = likely title)
            const relativeY = emusToPixels(top) / slideHeight;
            const isLikelyTitle = relativeY < 0.15;  // Top 15% of slide
            
            shape.textData = {
                paragraphs: [{
                    align: "left",
                    runs: [{ 
                        text: data.text,
                        style: { 
                            fontSize: isLikelyTitle ? "40pt" : "18pt",  // Title vs body
                            fontWeight: isLikelyTitle ? "bold" : "normal",
                            color: "#000000" 
                        }
                    }],
                    level: 0,
                    marL: 0,
                    indent: 0
                }],
                verticalAlign: "flex-start"
            };
        }
        
        return shape;
    }
    
    // If we only have text (no bounds), create a default text box
    if (data.text) {
        console.log("Creating text-only shape (no bounds):", data.text.substring(0, 50));
        return {
            type: "text",
            box: {
                x: Math.round(slideWidth * 0.05),
                y: Math.round(slideHeight * 0.08),   // Higher
                cx: Math.round(slideWidth * 0.9),
                cy: Math.round(slideHeight * 0.15)    // Title-sized box
            },
            textData: {
                paragraphs: [{
                    align: "left",
                    runs: [{ 
                        text: data.text,
                        style: { 
                            fontSize: "16pt",    // Title size
                            fontWeight: "bold",
                            color: "#000000" 
                        }
                    }],
                    level: 0,
                    marL: 0,
                    indent: 0
                }],
                verticalAlign: "flex-start"
            },
            isMaster: false
        };
    }
    
    return null;
}

function renderSlidesToHtml(slides) {
    return slides
        .map((slide, idx) => {
            const scale = VIEW_WIDTH / slide.size.cx;
            const heightPx = Math.round(slide.size.cy * scale);
            const backgroundColor = slide.background?.color || "#ffffff";
            
            console.log(`Rendering slide ${idx} with ${slide.shapes.length} shapes`);
            
            const shapesHtml = slide.shapes
                .map((shape) => {
                    const left = Math.round(shape.box.x * scale);
                    const top = Math.round(shape.box.y * scale);
                    const width = Math.round(shape.box.cx * scale);
                    const height = Math.round(shape.box.cy * scale);
                    
                    console.log("Rendering shape:", {
                        type: shape.type,
                        hasTextData: !!shape.textData,
                        hasFill: !!shape.fill,
                        box: { left, top, width, height }
                    });
                    
                    let bgStyle = '';
                    if (shape.fill) {
                        if (shape.fill.type === 'solid') {
                            bgStyle = `background: ${shape.fill.color};`;
                        } else if (shape.fill.type === 'none') {
                            bgStyle = 'background: transparent;';
                        }
                    }
                    
                    const nearlySquare = Math.abs(width - height) / Math.max(width, height) < 0.15;
                    let borderRadius = 0;
                    if (shape.geom === "roundRect") {
                        borderRadius = 12;
                    } else if (shape.geom === "ellipse") {
                        borderRadius = Math.min(width, height) / 2;
                    } else if (!shape.geom && nearlySquare) {
                        borderRadius = Math.min(width, height) / 2;
                    }

                    if (shape.type === "image") {
                        return `<img class="shape image-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;border-radius:${borderRadius}px;" src="${shape.src}" alt="" />`;
                    }

                    if (shape.type === "table" && Array.isArray(shape.data)) {
                        const rowsHtml = shape.data
                            .map((row, idx) => {
                                const tag = idx === 0 ? "th" : "td";
                                const cells = row.map((cell) => `<${tag}>${escapeHtml(cell || "")}</${tag}>`).join("");
                                return `<tr>${cells}</tr>`;
                            })
                            .join("");
                        return `<div class="shape table-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;">` +
                            `<table>${rowsHtml}</table>` +
                            `</div>`;
                    }

                    if (shape.type === "chart" && shape.data) {
                        const colors = ["#2b5797", "#d24726", "#e3a21a", "#2d89ef"];
                        const max = Math.max(...shape.data.values.flat(), 0);
                        const plotPadding = { left: 40, right: 120, top: 10, bottom: 30 };
                        const plotWidth = Math.max(80, width - plotPadding.left - plotPadding.right);
                        const plotHeight = Math.max(80, height - plotPadding.top - plotPadding.bottom);
                        const categoryCount = shape.data.categories.length || 1;
                        const seriesCount = shape.data.headers.length || 1;
                        const groupWidth = plotWidth / categoryCount;
                        const barWidth = Math.max(8, Math.floor(groupWidth / seriesCount) - 6);
                        const barGap = Math.max(4, Math.floor((groupWidth - barWidth * seriesCount) / (seriesCount + 1)));

                        let barsHtml = "";
                        let labelsHtml = "";

                        shape.data.categories.forEach((cat, catIdx) => {
                            const groupStart = plotPadding.left + catIdx * groupWidth;
                            shape.data.headers.forEach((header, sIdx) => {
                                const val = shape.data.values[sIdx]?.[catIdx] ?? 0;
                                const h = max > 0 ? Math.round((val / max) * plotHeight) : 0;
                                const leftPos = groupStart + barGap * (sIdx + 1) + barWidth * sIdx;
                                const bottomPos = plotPadding.bottom;
                                barsHtml += `<div class="chart-bar" style="left:${leftPos}px;bottom:${bottomPos}px;width:${barWidth}px;height:${h}px;background:${colors[sIdx % colors.length]};" title="${escapeHtml(header)}: ${val}"></div>`;
                            });
                            const labelCenter = groupStart + groupWidth / 2;
                            labelsHtml += `<div class="chart-x-label" style="left:${labelCenter}px;bottom:${plotPadding.bottom - 18}px;">${escapeHtml(cat)}</div>`;
                        });

                        const legend = shape.data.headers
                            .map((name, idx) => `<div class="legend-item"><span class="legend-swatch" style="background:${colors[idx % colors.length]};"></span>${escapeHtml(name)}</div>`)
                            .join("");

                        return `<div class="shape chart-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;">` +
                            `<div class="chart-grid"></div>` +
                            `<div class="chart-plot" style="height:${height}px;">${barsHtml}${labelsHtml}</div>` +
                            `<div class="chart-legend">${legend}</div>` +
                            `</div>`;
                    }
                    
                    if (shape.textData) {
                        console.log(`Rendering text with ${shape.textData.paragraphs.length} paragraphs`);
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
                            const indentPx = Math.max(0, Math.round(((para.marL || 0) + (para.indent || 0)) * scale));
                            let bulletHtml = "";
                            if (para.bullet?.type === "char") {
                                const bulletSize = para.runs[0]?.style.fontSize ? `font-size:${para.runs[0].style.fontSize}` : "";
                                bulletHtml = `<span class="bullet" style="${bulletSize}">${escapeHtml(para.bullet.char)}</span>`;
                            } else if (para.bullet?.type === "auto") {
                                const bulletSize = para.runs[0]?.style.fontSize ? `font-size:${para.runs[0].style.fontSize}` : "";
                                const prefix = para.bullet.prefix || "";
                                const suffix = para.bullet.suffix || ".";
                                bulletHtml = `<span class="bullet" style="${bulletSize}">${escapeHtml(prefix)}${para.bullet.index}${escapeHtml(suffix)}</span>`;
                            }
                            return `<p class="para" style="text-align:${textAlign}; padding-left:${indentPx}px;">${bulletHtml}<span>${runHtml}</span></p>`;
                        }).join('');
                        
                        return `<div class="shape text-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;${bgStyle}align-items:${verticalAlign};justify-content:${verticalAlign};border-radius:${borderRadius}px;">${textHtml}</div>`;
                    } else {
                        return `<div class="shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;${bgStyle};border-radius:${borderRadius}px;"></div>`;
                    }
                })
                .join("");
            
            return `
                <article class="slide-frame" id="slide-${idx}" style="width:${VIEW_WIDTH}px;height:${heightPx}px;">
                    <div class="slide-canvas" style="width:${VIEW_WIDTH}px;height:${heightPx}px;background:${backgroundColor};">
                        ${shapesHtml}
                    </div>
                </article>
            `;
        })
        .join("");
}

window.addEventListener("message", async (event) => {
    const msg = event.data;
    if (msg?.type === "loadFile") {
        const name = document.getElementById("file-name");
        const slidesEl = document.getElementById("slides");
        const slidesContent = document.getElementById("slides-content");
        
        try {
            name.textContent = msg.fileName ?? "Presentation";
            document.body.dataset.loaded = "true";
            
            const lowerName = msg.fileName?.toLowerCase() || "";

            if (lowerName.endsWith(".pptx")) {
                const slides = await renderPptxSlides(msg.base64);
                slidesCache = slides;
                
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
            } else if (lowerName.endsWith(".ppt")) {
                const slides = await renderPptSlides(msg.base64);
                slidesCache = slides;

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
            } else if (lowerName.endsWith(".odp")) {
                const slides = await renderOdpSlides(msg.base64);
                slidesCache = slides;

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
                slidesContent.innerHTML = `<p>Preview for ${lowerName} not implemented.</p>`;
                slidesEl.classList.remove("hidden");
            }
        } catch (err) {
            console.error("Error loading presentation:", err);
            slidesContent.innerHTML = `<p>Error loading presentation: ${err.message}</p>`;
            slidesEl.classList.remove("hidden");
        }
    }
});

function bindControls() {
    const prev = document.getElementById("prev");
    const next = document.getElementById("next");
    const zoomIn = document.getElementById("zoom-in");
    const zoomOut = document.getElementById("zoom-out");
    const zoomReset = document.getElementById("zoom-reset");
    
    prev?.addEventListener("click", () => changeSlide(-1));
    next?.addEventListener("click", () => changeSlide(1));
    zoomIn?.addEventListener("click", () => changeZoom(0.1));
    zoomOut?.addEventListener("click", () => changeZoom(-0.1));
    zoomReset?.addEventListener("click", () => setZoom(1));
    // Log toggle removed
    
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
