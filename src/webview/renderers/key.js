// Keynote (.key) renderer – APXL-based

import {
    decodeBase64ToUint8,
    parseXml,
    guessMimeFromBytes,
    uint8ToBase64,
    extractTrailingNumber,
    getImageDimensions
} from "../utils.js";

// ---------------------------------------------------------------------------
// Basic geometry helpers
// ---------------------------------------------------------------------------

function reorderShapesForZIndex(shapes) {
    // Keep original ordering inside each group, but guarantee:
    //   background / images / vectors  → first
    //   text shapes                    → last (on top)
    const nonText = [];
    const text = [];

    for (const s of shapes) {
        if (s && s.type === "text") {
            text.push(s);
        } else {
            nonText.push(s);
        }
    }

    return [...nonText, ...text];
}


function getSlideBackground(slideNode, graphicStyleIndex) {
    // Look for explicit <sf:slide-background> or <sf:background-fill>
    const bgFill = Array.from(slideNode.getElementsByTagName("*")).find(
        el => 
            el.localName === "background-fill" || 
            el.localName === "slide-background" ||
            el.localName === "background"
    );

    if (bgFill) {
        const col = extractColorFromElement(bgFill);
        if (col) return { color: col, found: true };
    }

    // Look for <sf:style> child attached to slide-level background
    const styleEl = Array.from(slideNode.getElementsByTagName("*")).find(
        el => el.localName === "style" && el.parentElement === slideNode
    );

    if (styleEl) {
        const fill = extractFillFromNode(styleEl, graphicStyleIndex);
        if (fill && fill.type === "solid") {
            return { color: fill.color, found: true };
        }
    }

    return { color: "#ffffff", found: false };
}

function findNearestFill(node, graphicStyleIndex) {
    let current = node;
    while (current) {
        const fill = extractFillFromNode(current, graphicStyleIndex);
        if (fill && fill.type && fill.type !== "none") {
            return fill;
        }
        current = current.parentElement;
    }
    return null;
}


function getSlideSizeFromDoc(doc) {
    const sizes = Array.from(doc.getElementsByTagName("*")).filter(
        (el) => el.localName === "size"
    );

    for (const el of sizes) {
        const parent = el.parentElement;
        if (!parent) continue;
        const parentName = (parent.localName || "").toLowerCase();
        if (parentName !== "presentation" && parentName !== "document") continue;

        const wAttr =
            el.getAttribute("sfa:w") ||
            el.getAttribute("sf:w") ||
            el.getAttribute("w");
        const hAttr =
            el.getAttribute("sfa:h") ||
            el.getAttribute("sf:h") ||
            el.getAttribute("h");

        const w = parseFloat(wAttr || "");
        const h = parseFloat(hAttr || "");
        if (!isNaN(w) && !isNaN(h) && w > 0 && h > 0) {
            return { cx: w, cy: h };
        }
    }

    return { cx: 1024, cy: 768 };
}

function getNumericAttr(node, names) {
    if (!node || !node.attributes) return null;
    for (const attr of Array.from(node.attributes)) {
        const lname = attr.name.toLowerCase();
        if (names.some((n) => lname === n || lname.endsWith(":" + n))) {
            const v = parseFloat(attr.value);
            if (!Number.isNaN(v)) return v;
        }
    }
    return null;
}

function getGeometryBox(geometryEl, defaultSize) {
    if (!geometryEl) {
        return {
            x: 0,
            y: 0,
            cx: defaultSize.cx,
            cy: defaultSize.cy
        };
    }

    const sizeEl = Array.from(geometryEl.children).find(
        (c) => c.localName === "size"
    );
    const posEl = Array.from(geometryEl.children).find(
        (c) => c.localName === "position"
    );

    const w = sizeEl ? getNumericAttr(sizeEl, ["w"]) : null;
    const h = sizeEl ? getNumericAttr(sizeEl, ["h"]) : null;
    const x = posEl ? getNumericAttr(posEl, ["x"]) : null;
    const y = posEl ? getNumericAttr(posEl, ["y"]) : null;

    return {
        x: x ?? 0,
        y: y ?? 0,
        cx: w ?? defaultSize.cx,
        cy: h ?? defaultSize.cy
    };
}

function findNearestGeometry(node) {
    let current = node;
    while (current) {
        const candidate = Array.from(current.children || []).find(
            (c) => c.localName === "geometry"
        );
        if (candidate) return candidate;
        current = current.parentElement;
    }
    return null;
}

// ---------------------------------------------------------------------------
// Fill / color helpers (for shapes and background)
// ---------------------------------------------------------------------------

function extractColorFromElement(el) {
    if (!el || !el.attributes) return null;

    const colorEl =
        el.localName === "color"
            ? el
            : Array.from(el.getElementsByTagName("*")).find(
                  (c) => c.localName === "color"
              );

    if (!colorEl) return null;

    const attrs = colorEl.attributes || {};
    const r = parseFloat(attrs.getNamedItem("sfa:r")?.value || "0");
    const g = parseFloat(attrs.getNamedItem("sfa:g")?.value || "0");
    const b = parseFloat(attrs.getNamedItem("sfa:b")?.value || "0");
    const a = parseFloat(attrs.getNamedItem("sfa:a")?.value || "1");

    const to255 = (x) => Math.round(Math.max(0, Math.min(1, x)) * 255);
    return `rgba(${to255(r)}, ${to255(g)}, ${to255(b)}, ${a.toFixed(3)})`;
}

function resolveGraphicFill(styleIndex, ref) {
    if (!styleIndex || !ref) return null;
    const cleaned = ref.replace(/^.*:/, "");
    return styleIndex[ref] || styleIndex[cleaned] || null;
}

function extractFillFromNode(node, graphicStyleIndex) {
    if (!node) return null;

    let styleEl = null;
    if (node.localName === "style") styleEl = node;
    else {
        styleEl = Array.from(node.getElementsByTagName("*")).find(
            (el) => el.localName === "style"
        );
    }
    if (!styleEl) return null;

    const fillEl = Array.from(styleEl.getElementsByTagName("*")).find(
        (el) =>
            el.localName === "fill" ||
            el.localName === "background-fill" ||
            el.localName === "shape-fill"
    );

    if (fillEl) {
        const color = extractColorFromElement(fillEl);
        if (!color) {
            return { type: "none" };
        }
        return {
            type: "solid",
            color
        };
    }

    const graphicRefEl = Array.from(styleEl.getElementsByTagName("*")).find(
        (el) => el.localName === "graphic-style-ref"
    );
    const graphicRefAttr =
        styleEl.getAttribute?.("sf:graphic-style-ref") ||
        styleEl.getAttribute?.("graphic-style-ref");

    const ref =
        graphicRefAttr ||
        graphicRefEl?.getAttribute("sfa:IDREF") ||
        graphicRefEl?.getAttribute("IDREF") ||
        graphicRefEl?.getAttribute("idref");
    const resolved = resolveGraphicFill(graphicStyleIndex, ref);
    if (resolved) return resolved;

    // Some themes use a background color directly on the style
    const col = extractColorFromElement(styleEl);
    if (!col) return { type: "none" };
    return { type: "solid", color: col };
}

// ---------------------------------------------------------------------------
// Text style system (Keynote APXL)
// ---------------------------------------------------------------------------

function parseNumberFromProperty(propEl) {
    if (!propEl) return null;
    const numEl = Array.from(propEl.getElementsByTagName("*")).find(
        (c) => c.localName === "number"
    );
    if (!numEl) return null;

    const attrs = numEl.attributes || {};
    const cand =
        (attrs.getNamedItem("sfa:number") || attrs.getNamedItem("number"))?.value;
    if (!cand) return null;
    const v = parseFloat(cand);
    return Number.isNaN(v) ? null : v;
}

function parseColorFromProperty(propEl) {
    if (!propEl) return null;
    const colorEl = Array.from(propEl.getElementsByTagName("*")).find(
        (c) => c.localName === "color"
    );
    if (!colorEl) return null;

    const attrs = colorEl.attributes || {};
    const r = parseFloat(attrs.getNamedItem("sfa:r")?.value || "0");
    const g = parseFloat(attrs.getNamedItem("sfa:g")?.value || "0");
    const b = parseFloat(attrs.getNamedItem("sfa:b")?.value || "0");
    const a = parseFloat(attrs.getNamedItem("sfa:a")?.value || "1");

    const to255 = (x) => Math.round(Math.max(0, Math.min(1, x)) * 255);
    return `rgba(${to255(r)}, ${to255(g)}, ${to255(b)}, ${a.toFixed(3)})`;
}

function parseTextStylePropertyMap(propMap) {
    const style = {};
    if (!propMap) return style;

    for (const child of Array.from(propMap.children)) {
        const ln = child.localName;

        if (ln === "font-size" || ln === "fontSize") {
            const v = parseNumberFromProperty(child);
            if (v != null) style.fontSize = v; // in pt
        } else if (ln === "font-color" || ln === "fontColor") {
            const color = parseColorFromProperty(child);
            if (color) style.color = color;
        } else if (ln === "alignment") {
            const v = parseNumberFromProperty(child);
            // typical: 0 = left, 1 = right, 2 = center, 3 = justified
            if (v === 2) style.align = "center";
            else if (v === 1) style.align = "right";
            else if (v === 3) style.align = "justify";
            else style.align = "left";
        } else if (ln === "bold") {
            const v = parseNumberFromProperty(child);
            if (v && v > 0) style.bold = true;
        } else if (ln === "italic") {
            const v = parseNumberFromProperty(child);
            if (v && v > 0) style.italic = true;
        } else if (ln === "font-name" || ln === "fontName" || ln === "font") {
            const txt = (child.textContent || "").trim();
            if (txt) style.fontFamily = txt;
        }
    }

    return style;
}

function buildTextStyleIndex(doc) {
    const byId = Object.create(null);
    const byIdent = Object.create(null);

    const stylesheets = Array.from(doc.getElementsByTagName("*")).filter(
        (el) => el.localName === "stylesheet"
    );

    for (const ss of stylesheets) {
        const stylesEl = Array.from(ss.children).find(
            (c) => c.localName === "styles"
        );
        if (!stylesEl) continue;

        for (const node of Array.from(stylesEl.children)) {
            const ln = node.localName;
            if (
                ln !== "paragraph-style" &&
                ln !== "paragraphstyle" &&
                ln !== "character-style" &&
                ln !== "characterstyle"
            ) {
                continue;
            }

            const idAttr =
                node.getAttribute("sfa:ID") ||
                node.getAttribute("ID") ||
                node.getAttribute("id");
            const identAttr =
                node.getAttribute("sf:ident") ||
                node.getAttribute("ident") ||
                node.getAttribute("name");

            const propMap = Array.from(node.children).find(
                (c) => c.localName === "property-map"
            );
            const style = parseTextStylePropertyMap(propMap);

            const entry = {
                ...style,
                kind:
                    ln === "paragraph-style" || ln === "paragraphstyle"
                        ? "paragraph"
                        : "character"
            };

            if (idAttr) byId[idAttr] = entry;
            if (identAttr) byIdent[identAttr] = entry;
        }
    }

    return { byId, byIdent };
}

function resolveTextStyle(styleIndex, ref) {
    if (!ref) return {};
    const cleaned = ref.replace(/^.*:/, ""); // drop "key:" / "sf:" prefixes
    const { byId, byIdent } = styleIndex;

    return (
        byId[ref] ||
        byIdent[ref] ||
        byId[cleaned] ||
        byIdent[cleaned] || {}
    );
}

function mergeTextStyles(base, override) {
    return {
        ...base,
        ...override
    };
}

// A: convert points to px using 96/72
function styleToViewerRunStyle(style) {
    const out = {};
    if (style.fontSize != null) {
        const px = style.fontSize * (96 / 72);
        out.fontSize = `${Math.round(px)}px`;
    }
    if (style.bold) out.fontWeight = "bold";
    if (style.italic) out.fontStyle = "italic";
    if (style.color) out.color = style.color;
    if (style.fontFamily) out.fontFamily = style.fontFamily;
    return out;
}

function getStyleRefFromNode(node) {
    if (!node) return null;

    const direct =
        node.getAttribute("sf:style-ref") ||
        node.getAttribute("style-ref") ||
        node.getAttribute("sf:style") ||
        node.getAttribute("style") ||
        node.getAttribute("sf:paragraph-style-ref") ||
        node.getAttribute("sf:character-style-ref");
    if (direct) return direct;

    const refEl = Array.from(node.children).find(
        (c) =>
            c.localName === "paragraph-style-ref" ||
            c.localName === "character-style-ref"
    );
    if (!refEl) return null;

    return (
        refEl.getAttribute("sfa:IDREF") ||
        refEl.getAttribute("IDREF") ||
        refEl.getAttribute("idref")
    );
}

function getGeomHint(node) {
    const attrs = Array.from(node.attributes || []);
    for (const attr of attrs) {
        const v = (attr.value || "").toLowerCase();
        if (v.includes("oval") || v.includes("ellipse") || v.includes("circle")) {
            return "ellipse";
        }
        if (v.includes("roundrect") || v.includes("rounded")) {
            return "roundRect";
        }
    }
    return null;
}

function buildGraphicStyleIndex(doc) {
    const map = Object.create(null);
    const nodes = Array.from(doc.getElementsByTagName("*")).filter(
        (el) => el.localName === "graphic-style"
    );

    for (const node of nodes) {
        const id =
            node.getAttribute("sfa:ID") ||
            node.getAttribute("ID") ||
            node.getAttribute("id");
        const ident = node.getAttribute("sf:ident") || node.getAttribute("ident");
        const propMap =
            Array.from(node.children).find((c) => c.localName === "property-map") ||
            node;

        const fillEl = Array.from(propMap.getElementsByTagName("*")).find(
            (el) =>
                el.localName === "fill" ||
                el.localName === "background-fill" ||
                el.localName === "shape-fill"
        );

        let fill = null;
        if (fillEl) {
            const color = extractColorFromElement(fillEl);
            if (color) fill = { type: "solid", color };
        } else {
            const color = extractColorFromElement(propMap);
            if (color) fill = { type: "solid", color };
        }

        if (fill) {
            if (id) map[id] = fill;
            if (ident && !map[ident]) map[ident] = fill;
        }
    }

    return map;
}

// ---------------------------------------------------------------------------
// Slide-level helpers
// ---------------------------------------------------------------------------

function getSlideNodes(doc) {
    const all = Array.from(doc.getElementsByTagName("*"));
    const slides = [];

    for (const el of all) {
        const ln = el.localName;
        if (ln === "slide" || ln === "slide-archive" || ln === "slide-node") {
            slides.push(el);
        }
    }

    if (slides.length > 0) return slides;

    const slideLists = all.filter((el) => el.localName === "slide-list");
    for (const list of slideLists) {
        const nodes = Array.from(list.children).filter(
            (el) =>
                el.localName === "slide-node" ||
                el.localName === "slide" ||
                el.localName === "slide-archive"
        );
        if (nodes.length) slides.push(...nodes);
    }

    return slides;
}

function getMasterSlideNodes(doc) {
    const all = Array.from(doc.getElementsByTagName("*"));
    return all.filter((el) =>
        ["master-slide", "master-slide-archive", "master-slide-node", "masterslide"].includes(
            (el.localName || "").toLowerCase()
        )
    );
}

function getMasterRefId(slideNode) {
    const masterRef = Array.from(slideNode.getElementsByTagName("*")).find(
        (el) => el.localName === "master-ref" || el.localName === "master-slide-ref"
    );
    if (!masterRef) return null;
    return (
        masterRef.getAttribute("sfa:IDREF") ||
        masterRef.getAttribute("IDREF") ||
        masterRef.getAttribute("idref")
    );
}

function getNodeId(node) {
    if (!node) return null;
    return (
        node.getAttribute("sfa:ID") ||
        node.getAttribute("ID") ||
        node.getAttribute("id")
    );
}

// ---------------------------------------------------------------------------
// ZIP helpers / images
// ---------------------------------------------------------------------------

function collectZipFileNames(zip) {
    return Object.keys(zip.files || {});
}

function findAssetFile(zip, fileNames, requestedPath) {
    if (!requestedPath) return null;

    const cleaned = requestedPath.replace(/^\/+/, "");
    const lower = cleaned.toLowerCase();
    const base = lower.split("/").pop();

    let hit = fileNames.find((n) => n.toLowerCase() === lower);
    if (hit) return zip.file(hit) || null;

    const prefixes = ["data/", "Data/", "thumbs/", "Thumbs/", "QuickLook/", "quicklook/"];
    for (const prefix of prefixes) {
        const tryName = (prefix + base).replace(/\/+/g, "/");
        hit = fileNames.find((n) => n.toLowerCase() === tryName.toLowerCase());
        if (hit) return zip.file(hit) || null;
    }

    hit = fileNames.find(
        (n) =>
            n.toLowerCase().split("/").pop() === base ||
            n.toLowerCase().endsWith("/" + base)
    );
    if (hit) return zip.file(hit) || null;

    return null;
}

async function buildImageShapeFromElement(zip, fileNames, imageEl, slideSize, isMasterShape = false) {
    const geometryEl = findNearestGeometry(imageEl);
    const binaryEl = Array.from(imageEl.getElementsByTagName("*")).find(
        (c) => c.localName === "binary"
    );
    if (!binaryEl) return null;

    const dataEl = Array.from(binaryEl.getElementsByTagName("*")).find(
        (c) => c.localName === "data"
    );
    if (!dataEl) return null;

    const pathAttr =
        dataEl.getAttribute("sf:path") ||
        dataEl.getAttribute("path") ||
        dataEl.getAttribute("sf:displayname") ||
        dataEl.getAttribute("displayname");
    if (!pathAttr) return null;

    const file = findAssetFile(zip, fileNames, pathAttr);
    if (!file) return null;

    const bytes = await file.async("uint8array");
    let mime = null;
    try {
        mime = guessMimeFromBytes(pathAttr, bytes);
    } catch {
        // ignore
    }
    if (!mime) {
        const lower = pathAttr.toLowerCase();
        if (lower.endsWith(".png")) mime = "image/png";
        else if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) mime = "image/jpeg";
        else if (lower.endsWith(".tif") || lower.endsWith(".tiff")) mime = "image/tiff";
        else mime = "image/octet-stream";
    }

    const dims = getImageDimensions(bytes, mime) || slideSize;
    const boxFromGeom = getGeometryBox(geometryEl, dims);
    const dataUrl = `data:${mime};base64,${uint8ToBase64(bytes)}`;

    return {
        type: "image",
        box: boxFromGeom,
        src: dataUrl,
        mime,
        isMaster: isMasterShape
    };
}

// ---------------------------------------------------------------------------
// Text and vector shapes
// ---------------------------------------------------------------------------

function extractTextShapeFromNode(node, slideSize, styleIndex, graphicStyleIndex, isMasterShape = false) {
    const all = Array.from(node.getElementsByTagName("*"));

    let textStorageEl = null;

    for (const el of all) {
        if (!textStorageEl && el.localName === "text-storage") textStorageEl = el;
        if (textStorageEl) break;
    }

    if (!textStorageEl) return null;

    const paraEls = Array.from(textStorageEl.getElementsByTagName("*")).filter(
        (el) => el.localName === "p" || el.localName === "paragraph"
    );
    const paragraphs = paraEls.length ? paraEls : [textStorageEl];

    const geometryEl = findNearestGeometry(node);
    const box = getGeometryBox(geometryEl, slideSize);
    const paraData = [];
    const fill = findNearestFill(node, graphicStyleIndex);
    const geom = getGeomHint(node);

    for (const p of paragraphs) {
        const baseParaStyle = resolveTextStyle(
            styleIndex,
            getStyleRefFromNode(p)
        );

        const runEls = Array.from(p.getElementsByTagName("*")).filter(
            (el) => el.localName === "run" || el.localName === "r" || el.localName === "span"
        );

        const runs = [];

        if (runEls.length) {
            for (const r of runEls) {
                const txt = (r.textContent || "").replace(/[ \t]+/g, " ").replace(/\u00a0/g, " ").trim();
                if (!txt) continue;

                const charStyle = resolveTextStyle(
                    styleIndex,
                    getStyleRefFromNode(r)
                );
                const merged = mergeTextStyles(baseParaStyle, charStyle);
                runs.push({
                    text: txt,
                    style: styleToViewerRunStyle(merged)
                });
            }
        } else {
            const txt = (p.textContent || "").replace(/[ \t]+/g, " ").replace(/\u00a0/g, " ").trim();
            if (txt) {
                const merged = baseParaStyle;
                runs.push({
                    text: txt,
                    style: styleToViewerRunStyle(merged)
                });
            }
        }

        if (!runs.length) continue;

        const align = baseParaStyle.align || "left";

        paraData.push({
            align,
            runs,
            bullet: null,
            level: 0,
            marL: 0,
            indent: 0
        });
    }

    if (!paraData.length) return null;

    return {
        type: "text",
        box,
        fill,
        textData: {
            verticalAlign: "center",
            paragraphs: paraData
        },
        geom,
        isMaster: isMasterShape
    };
}

function extractVectorShapeFromNode(node, slideSize, graphicStyleIndex, isMasterShape = false) {
    const all = Array.from(node.getElementsByTagName("*"));

    let styleEl = null;

    for (const el of all) {
        if (!styleEl && el.localName === "style") styleEl = el;
        if (styleEl) break;
    }

    const fill = extractFillFromNode(styleEl || node, graphicStyleIndex);
    if (!fill || fill.type === "none") return null;

    const geometryEl = findNearestGeometry(node);
    const box = getGeometryBox(geometryEl, slideSize);

    const geom = getGeomHint(node);

    return {
        type: "shape",
        box,
        fill,
        geom,
        isMaster: isMasterShape
    };
}

// ---------------------------------------------------------------------------
// Slide walker / shape collection
// ---------------------------------------------------------------------------

async function collectShapesForSlide(
    slideNode,
    zip,
    fileNames,
    slideSize,
    styleIndex,
    graphicStyleIndex,
    markAsMaster = false
) {
    const shapes = [];
    const imagePromises = [];

    function walk(node) {
        const ln = node.localName;

        if (
            ln === "drawables" ||
            ln === "group" ||
            ln === "slide" ||
            ln === "slide-archive" ||
            ln === "slide-node" ||
            ln === "layer" ||
            ln === "layers" ||
            ln === "page"
        ) {
            for (const child of Array.from(node.children)) {
                walk(child);
            }
            return;
        }

        if (ln === "image") {
            imagePromises.push(
                buildImageShapeFromElement(zip, fileNames, node, slideSize, markAsMaster)
            );
            return;
        }

        if (ln === "text" || ln === "drawable-shape" || ln === "shape") {
            const textShape = extractTextShapeFromNode(
                node,
                slideSize,
                styleIndex,
                graphicStyleIndex,
                markAsMaster
            );
            if (textShape) {
                shapes.push(textShape);
            } else if (ln === "shape") {
                const vec = extractVectorShapeFromNode(node, slideSize, graphicStyleIndex, markAsMaster);
                if (vec) shapes.push(vec);
            }

            for (const child of Array.from(node.children)) {
                walk(child);
            }
            return;
        }

        for (const child of Array.from(node.children)) {
            walk(child);
        }
    }

    walk(slideNode);

    if (imagePromises.length) {
        const imgs = await Promise.all(imagePromises);
        for (const s of imgs) {
            if (s) shapes.push(s);
        }
    }

    return shapes;
}

// ---------------------------------------------------------------------------
// Background detection
// ---------------------------------------------------------------------------

function pickBackgroundFromShapes(shapes, slideSize) {
    const areaSlide = slideSize.cx * slideSize.cy;
    let best = null;
    let bestArea = 0;

    for (const s of shapes) {
        const isFilledText = s.type === "text" && s.fill && s.fill.type === "solid";
        if (s.type !== "shape" && s.type !== "image" && !isFilledText) continue;
        const area = (s.box?.cx || 0) * (s.box?.cy || 0);
        const coverage = area / areaSlide;
        if (coverage < 0.6) continue;
        if (area > bestArea) {
            best = s;
            bestArea = area;
        }
    }

    if (!best) return { background: { color: "#ffffff" }, shapes };

    const others = shapes.filter((s) => s !== best);
    const ordered = [best, ...others];

    let bgColor = "#ffffff";
    if (best.type === "shape" && best.fill && best.fill.type === "solid") {
        bgColor = best.fill.color;
    } else if (best.type === "text" && best.fill && best.fill.type === "solid") {
        bgColor = best.fill.color;
    }

    return {
        background: { color: bgColor },
        shapes: ordered
    };
}

// ---------------------------------------------------------------------------
// Fallback: image-only slides
// ---------------------------------------------------------------------------

async function buildFallbackSlidesFromImages(zip, maxSlides) {
    const fileNames = collectZipFileNames(zip);
    const imageNames = fileNames.filter((name) => {
        if (zip.files[name].dir) return false;
        const lower = name.toLowerCase();
        return (
            lower.endsWith(".png") ||
            lower.endsWith(".jpg") ||
            lower.endsWith(".jpeg") ||
            lower.endsWith(".tif") ||
            lower.endsWith(".tiff")
        );
    });

    if (!imageNames.length) return [];

    imageNames.sort((a, b) => {
        const al = a.toLowerCase();
        const bl = b.toLowerCase();
        const score = (name) => {
            let s = 0;
            if (name.includes("quicklook")) s -= 10;
            if (name.includes("preview")) s -= 8;
            if (name.includes("thumb")) s -= 6;
            if (name.includes("slide")) s -= 4;
            const n = extractTrailingNumber(name);
            if (!Number.isNaN(n)) s += n / 1000;
            return s;
        };
        return score(al) - score(bl);
    });

    const slides = [];
    for (const name of imageNames.slice(0, maxSlides)) {
        const file = zip.file(name);
        if (!file) continue;

        const bytes = await file.async("uint8array");
        const mime = guessMimeFromBytes(name, bytes);
        const dims = getImageDimensions(bytes, mime) || { cx: 1280, cy: 720 };
        const dataUrl = `data:${mime};base64,${uint8ToBase64(bytes)}`;

        slides.push({
            size: dims,
            background: { color: "#ffffff" },
            shapes: [
                {
                    type: "image",
                    box: { x: 0, y: 0, cx: dims.cx, cy: dims.cy },
                    src: dataUrl,
                    mime,
                    isMaster: false
                }
            ]
        });
    }

    return slides;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export async function renderKeySlides(base64, maxSlides = 20) {
    try {
        const buffer = decodeBase64ToUint8(base64);
        const zip = await JSZip.loadAsync(buffer);
        const fileNames = collectZipFileNames(zip);

        const indexFile =
            zip.file("index.apxl") ||
            zip.file("Index.apxl") ||
            zip.file("index.apxl.gz") ||
            zip.file("Index.apxl.gz");

        if (!indexFile) {
            return await buildFallbackSlidesFromImages(zip, maxSlides);
        }

        const xmlText = await indexFile.async("string");
        const doc = parseXml(xmlText);
        if (!doc) {
            return await buildFallbackSlidesFromImages(zip, maxSlides);
        }

        const slideSize = getSlideSizeFromDoc(doc);
        const slideNodes = getSlideNodes(doc);
        const masterNodes = getMasterSlideNodes(doc);
        const styleIndex = buildTextStyleIndex(doc);
        const graphicStyleIndex = buildGraphicStyleIndex(doc);

        const masterShapeMap = Object.create(null);
        const masterBackgroundMap = Object.create(null);
        const masterNodeMap = Object.create(null);

        for (const masterNode of masterNodes) {
            const mid = getNodeId(masterNode);
            if (mid) masterNodeMap[mid] = masterNode;

            const mShapes = await collectShapesForSlide(
                masterNode,
                zip,
                fileNames,
                slideSize,
                styleIndex,
                graphicStyleIndex,
                true
            );
            if (mid && mShapes.length) masterShapeMap[mid] = mShapes;

            const mBg = getSlideBackground(masterNode, graphicStyleIndex);
            if (mid) masterBackgroundMap[mid] = mBg;
        }

        const slides = [];
        for (const slideNode of slideNodes) {
            if (slides.length >= maxSlides) break;

            const masterRef = getMasterRefId(slideNode);
            const inheritedShapes = masterRef && masterShapeMap[masterRef] ? masterShapeMap[masterRef] : [];

            const slideShapes = await collectShapesForSlide(
                slideNode,
                zip,
                fileNames,
                slideSize,
                styleIndex,
                graphicStyleIndex
            );
            let combinedShapes = [...inheritedShapes, ...slideShapes];
            if (!combinedShapes.length) continue;

            let background = getSlideBackground(slideNode, graphicStyleIndex);
            if (!background?.found && masterRef && masterBackgroundMap[masterRef]) {
                background = masterBackgroundMap[masterRef];
            }

            if (!background || !background.found) {
                const picked = pickBackgroundFromShapes(combinedShapes, slideSize);
                background = picked.background;
                combinedShapes = picked.shapes;
            }

            const orderedShapes = reorderShapesForZIndex(combinedShapes);

            slides.push({
                size: slideSize,
                background,
                shapes: orderedShapes
            });
        }

        if (!slides.length) {
            return await buildFallbackSlidesFromImages(zip, maxSlides);
        }

        return slides;
    } catch (e) {
        console.error("Error rendering .key:", e);
        return [];
    }
}
