// Keynote (.key) renderer – APXL-based

import {
    decodeBase64ToUint8,
    parseXml,
    guessMimeFromBytes,
    uint8ToBase64,
    extractTrailingNumber,
    getImageDimensions
} from "../utils.js";

function reorderShapesForZIndex(shapes) {
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


async function getSlideBackground(slideNode, graphicStyleIndex, slideStyleIndex, imageBinaryIndex, zip, fileNames) {
    const candidates = Array.from(slideNode.getElementsByTagName("*")).filter(
        (el) =>
            el.localName === "background-fill" ||
            el.localName === "slide-background" ||
            el.localName === "background"
    );

    for (const el of candidates) {
        const fill = await extractFillFromNode(el, graphicStyleIndex, imageBinaryIndex, zip, fileNames);
        if (fill) {
            if (fill.type === "image") return { image: fill.src, color: "#ffffff", found: true };
            if (fill.type === "gradient") return { gradient: fill.colors, color: fill.colors?.[0] ?? "#ffffff", found: true };
            if (fill.type === "solid" && fill.color) return { color: fill.color, found: true };
        }
    }

    const styleRefAttr =
        slideNode.getAttribute?.("sf:style-ref") ||
        slideNode.getAttribute?.("style-ref");
    const styleRefEl = Array.from(slideNode.children || []).find(
        (el) => el.localName === "style-ref"
    );
    const styleRef = styleRefAttr || getIdRef(styleRefEl);

    if (styleRef && slideStyleIndex?.[styleRef]) {
        const fill = slideStyleIndex[styleRef];
        if (fill.type === "image-ref") {
            const img = await resolveImageRef(fill.path, zip, fileNames);
            if (img) return { image: img.src, color: "#ffffff", found: true };
        } else if (fill.type === "image") {
            return { image: fill.src, color: "#ffffff", found: true };
        } else if (fill.type === "gradient") {
            return { gradient: fill.colors, color: fill.colors?.[0] ?? "#ffffff", found: true };
        } else if (fill.type === "solid" && fill.color) {
            return { color: fill.color, found: true };
        }
    }

    // Look for <sf:style> child attached to slide-level background
    const styleEl = Array.from(slideNode.getElementsByTagName("*")).find(
        (el) => el.localName === "style" && el.parentElement === slideNode
    );

    if (styleEl) {
        const fill = await extractFillFromNode(styleEl, graphicStyleIndex, imageBinaryIndex, zip, fileNames);
        if (fill) {
            if (fill.type === "image") return { image: fill.src, color: "#ffffff", found: true };
            if (fill.type === "gradient") return { gradient: fill.colors, color: fill.colors?.[0] ?? "#ffffff", found: true };
            if (fill.type === "solid" && fill.color) {
                return { color: fill.color, found: true };
            }
        }
    }

    return { color: "#ffffff", found: false };
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
            (c) => c.localName === "geometry" || c.localName === "crop-geometry"
        );
        if (candidate) return candidate;
        current = current.parentElement;
    }
    return null;
}

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

async function resolveImageRef(pathAttr, zip, fileNames) {
    if (!zip || !fileNames || !pathAttr) return null;
    const file = findAssetFile(zip, fileNames, pathAttr);
    if (!file) return null;
    
    const bytes = await file.async("uint8array");
    const lower = (pathAttr || "").toLowerCase();
    let mime = guessMimeFromBytes(pathAttr, bytes);
    
    if (lower.endsWith(".svg") || (bytes[0] === 0x3c && (bytes[1] === 0x3f || bytes[1] === 0x73))) {
        mime = "image/svg+xml";
        try {
            const svgText = await file.async("text");
            const blob = new Blob([svgText], { type: "image/svg+xml" });
            const base64 = uint8ToBase64(new Uint8Array(await blob.arrayBuffer()));
            return {
                src: `data:image/svg+xml;base64,${base64}`,
                mime: "image/svg+xml"
            };
        } catch (e) {
            console.warn("Failed to load SVG:", e);
        }
    }
    
    if (lower.endsWith(".tif") || lower.endsWith(".tiff") ||
        (bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0x00) ||
        (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0x00 && bytes[3] === 0x2a)) {
        mime = "image/tiff";
    }
    
    return {
        src: `data:${mime};base64,${uint8ToBase64(bytes)}`,
        mime
    };
}

async function extractFillFromNode(node, graphicStyleIndex, imageBinaryIndex, zip, fileNames) {
    if (!node) return null;

    let styleEl = null;
    if (node.localName === "style") styleEl = node;
    else {
        styleEl = Array.from(node.getElementsByTagName("*")).find(
            (el) => el.localName === "style"
        );
    }

    let fillEl = null;
    if (
        node.localName === "fill" ||
        node.localName === "background-fill" ||
        node.localName === "shape-fill" ||
        node.localName === "graphic-fill"
    ) {
        fillEl = node;
    } else if (styleEl) {
        fillEl = Array.from(styleEl.getElementsByTagName("*")).find(
            (el) =>
                el.localName === "fill" ||
                el.localName === "background-fill" ||
                el.localName === "shape-fill" ||
                el.localName === "graphic-fill"
        );
    }

    if (!fillEl && !styleEl) return null;

    if (fillEl) {
        // Prefer image/pattern fills if a path is present.
        const pathAttr =
            fillEl.getAttribute("sf:path") ||
            fillEl.getAttribute("path") ||
            fillEl.getAttribute("sf:displayname") ||
            fillEl.getAttribute("displayname");

        const dataEl = Array.from(fillEl.getElementsByTagName("*")).find(
            (el) => el.localName === "data"
        );
        const dataPath =
            dataEl?.getAttribute("sf:path") ||
            dataEl?.getAttribute("path") ||
            dataEl?.getAttribute("sf:displayname") ||
            dataEl?.getAttribute("displayname");

        const refEl = Array.from(fillEl.getElementsByTagName("*")).find(
            (el) => el.localName === "unfiltered-ref" || el.localName === "data-ref"
        );
        const refId = getIdRef(refEl);
        const refPath = refId && imageBinaryIndex?.[refId]?.path;

        const pickedPath = dataPath || pathAttr || refPath;
        if (pickedPath) {
            const img = await resolveImageRef(pickedPath, zip, fileNames);
            if (img) {
                return {
                    type: "image",
                    src: img.src,
                    mime: img.mime
                };
            }
        } else if (refId) {
            const imgEntry = imageBinaryIndex?.[refId];
            if (imgEntry?.path) {
                const img = await resolveImageRef(imgEntry.path, zip, fileNames);
                if (img) {
                    return {
                        type: "image",
                        src: img.src,
                        mime: img.mime
                    };
                }
            }
        }

        // Simple gradient approximation: collect two colors if present.
        const colorNodes = Array.from(fillEl.getElementsByTagName("*")).filter(
            (el) => el.localName === "color"
        );
        if (colorNodes.length >= 2) {
            const colors = colorNodes.map((c) => extractColorFromElement(c)).filter(Boolean);
            if (colors.length >= 2) {
                return {
                    type: "gradient",
                    colors: colors.slice(0, 2)
                };
            }
        }

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
    if (resolved) {
        if (resolved.type === "image-ref") {
            const img = await resolveImageRef(resolved.path, zip, fileNames);
            if (img) return { type: "image", src: img.src, mime: img.mime };
        } else {
            return resolved;
        }
    }

    // Some themes use a background color directly on the style
    const col = extractColorFromElement(styleEl);
    if (!col) return { type: "none" };
    return { type: "solid", color: col };
}

function extractStrokeFromNode(node) {
    if (!node) return null;
    const strokeEl = Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "stroke");
    if (!strokeEl) return null;

    const color = extractColorFromElement(strokeEl);
    const widthAttr =
        strokeEl.getAttribute("sfa:width") ||
        strokeEl.getAttribute("width") ||
        strokeEl.getAttribute("sf:width");
    const width = widthAttr != null ? parseFloat(widthAttr) : null;

    if (!color && (width == null || Number.isNaN(width))) return null;
    return {
        color: color || "#000000",
        width: !Number.isNaN(width) && width != null ? width : 1
    };
}

function parseNumberFromProperty(propEl) {
    if (!propEl) return null;
    const numEl = Array.from(propEl.getElementsByTagName("*")).find(
        (c) => c.localName === "number" || c.localName === "decimal-number"
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
        const styleContainers = Array.from(ss.children).filter(
            (c) => c.localName === "styles" || c.localName === "anon-styles"
        );
        if (!styleContainers.length) continue;

        for (const stylesEl of styleContainers) {
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
            const parentIdent =
                node.getAttribute("sf:parent-ident") ||
                node.getAttribute("parent-ident");

            const propMap = Array.from(node.children).find(
                (c) => c.localName === "property-map"
            );
            const style = parseTextStylePropertyMap(propMap);

                const entry = {
                    ...style,
                    parentIdent: parentIdent || null,
                    kind:
                        ln === "paragraph-style" || ln === "paragraphstyle"
                            ? "paragraph"
                            : "character"
                };

                if (idAttr) byId[idAttr] = entry;
                if (identAttr) byIdent[identAttr] = entry;
            }
        }
    }

    return { byId, byIdent };
}

function resolveTextStyle(styleIndex, ref) {
    if (!ref) return {};
    const cleaned = ref.replace(/^.*:/, ""); // drop "key:" / "sf:" prefixes
    const { byId, byIdent } = styleIndex;

    const lookup = (key) => byId[key] || byIdent[key] || byId[key.replace(/^.*:/, "")] || byIdent[key.replace(/^.*:/, "")];
    let entry = lookup(ref) || lookup(cleaned);
    if (!entry) return {};

    const visited = new Set();
    const chain = [];
    while (entry && !visited.has(entry)) {
        chain.push(entry);
        visited.add(entry);
        if (!entry.parentIdent) break;
        entry = lookup(entry.parentIdent);
    }

    let merged = {};
    for (let i = chain.length - 1; i >= 0; i--) {
        merged = { ...merged, ...chain[i] };
    }
    return merged;
}

function mergeTextStyles(base, override) {
    return {
        ...base,
        ...override
    };
}

// Convert points to px using 96/72 and a gentle downscale to better match Keynote sizing in the webview
function styleToViewerRunStyle(style) {
    const out = {};
    if (style.fontSize != null) {
        const px = style.fontSize * (96 / 72) * 0.5;
        out.fontSize = `${Math.round(px)}px`;
    }
    if (style.bold) out.fontWeight = "bold";
    if (style.italic) out.fontStyle = "italic";
    out.color = style.color || "#000000";
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

function hasPlaceholderAncestor(node) {
    let current = node;
    while (current) {
        const name = (current.localName || "").toLowerCase();
        if (name.includes("placeholder")) return true;
        current = current.parentElement;
    }
    return false;
}

function buildGraphicStyleIndex(doc, imageBinaryIndex) {
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
            const pathAttr =
                fillEl.getAttribute("sf:path") ||
                fillEl.getAttribute("path") ||
                fillEl.getAttribute("sf:displayname") ||
                fillEl.getAttribute("displayname");
            const dataEl = Array.from(fillEl.getElementsByTagName("*")).find(
                (el) => el.localName === "data"
            );
            const dataPath =
                dataEl?.getAttribute("sf:path") ||
                dataEl?.getAttribute("path") ||
                dataEl?.getAttribute("sf:displayname") ||
                dataEl?.getAttribute("displayname");
            const refEl = Array.from(fillEl.getElementsByTagName("*")).find(
                (el) => el.localName === "unfiltered-ref" || el.localName === "data-ref"
            );
            const refId = getIdRef(refEl);
            const refPath = refId && imageBinaryIndex?.[refId]?.path;
            const chosenPath = dataPath || pathAttr || refPath;

            if (chosenPath) {
                fill = { type: "image-ref", path: chosenPath };
            } else {
                const colorNodes = Array.from(fillEl.getElementsByTagName("*")).filter(
                    (el) => el.localName === "color"
                );
                if (colorNodes.length >= 2) {
                    const colors = colorNodes.map((c) => extractColorFromElement(c)).filter(Boolean);
                    if (colors.length >= 2) {
                        fill = { type: "gradient", colors: colors.slice(0, 2) };
                    }
                }
                if (!fill) {
                    const color = extractColorFromElement(fillEl);
                    if (color) fill = { type: "solid", color };
                }
            }
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

function buildSlideStyleIndex(doc, imageBinaryIndex) {
    const map = Object.create(null);
    const nodes = Array.from(doc.getElementsByTagName("*")).filter(
        (el) => el.localName === "slide-style"
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
            const pathAttr =
                fillEl.getAttribute("sf:path") ||
                fillEl.getAttribute("path") ||
                fillEl.getAttribute("sf:displayname") ||
                fillEl.getAttribute("displayname");
            const dataEl = Array.from(fillEl.getElementsByTagName("*")).find(
                (el) => el.localName === "data"
            );
            const dataPath =
                dataEl?.getAttribute("sf:path") ||
                dataEl?.getAttribute("path") ||
                dataEl?.getAttribute("sf:displayname") ||
                dataEl?.getAttribute("displayname");
            const refEl = Array.from(fillEl.getElementsByTagName("*")).find(
                (el) => el.localName === "unfiltered-ref" || el.localName === "data-ref"
            );
            const refId = getIdRef(refEl);
            const refPath = refId && imageBinaryIndex?.[refId]?.path;
            const chosenPath = dataPath || pathAttr || refPath;

            if (chosenPath) {
                fill = { type: "image-ref", path: chosenPath };
            } else {
                const colorNodes = Array.from(fillEl.getElementsByTagName("*")).filter(
                    (el) => el.localName === "color"
                );
                if (colorNodes.length >= 2) {
                    const colors = colorNodes.map((c) => extractColorFromElement(c)).filter(Boolean);
                    if (colors.length >= 2) {
                        fill = { type: "gradient", colors: colors.slice(0, 2) };
                    }
                }
                if (!fill) {
                    const color = extractColorFromElement(fillEl);
                    if (color) fill = { type: "solid", color };
                }
            }
        } else {
            const color = extractColorFromElement(propMap);
            if (color) fill = { type: "solid", color };
        }

        if (fill && (id || ident)) {
            if (id) map[id] = fill;
            if (ident && !map[ident]) map[ident] = fill;
        }
    }

    return map;
}

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

function getIdRef(node) {
    if (!node) return null;
    return (
        node.getAttribute("sfa:IDREF") ||
        node.getAttribute("IDREF") ||
        node.getAttribute("idref")
    );
}

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

function buildImageBinaryIndex(doc) {
    const map = Object.create(null);
    const candidates = Array.from(doc.getElementsByTagName("*")).filter((el) => {
        const ln = (el.localName || "").toLowerCase();
        return (
            ln === "unfiltered" ||
            ln === "filtered-image" ||
            ln === "image-binary" ||
            ln === "imagebinary" ||
            ln === "binary"
        );
    });

    for (const node of candidates) {
        const id = getNodeId(node);
        const dataEl = Array.from(node.getElementsByTagName("*")).find((c) => c.localName === "data");
        if (!id || !dataEl) continue;

        const path =
            dataEl.getAttribute("sf:path") ||
            dataEl.getAttribute("path") ||
            dataEl.getAttribute("sf:displayname") ||
            dataEl.getAttribute("displayname");
        if (!path) continue;

        const sizeEl = Array.from(node.getElementsByTagName("*")).find((c) => c.localName === "size");
        const w = sizeEl ? parseFloat(sizeEl.getAttribute("sfa:w") || sizeEl.getAttribute("w") || "") : null;
        const h = sizeEl ? parseFloat(sizeEl.getAttribute("sfa:h") || sizeEl.getAttribute("h") || "") : null;
        const size =
            w != null && h != null && !Number.isNaN(w) && !Number.isNaN(h)
                ? { cx: w, cy: h }
                : null;

        const entry = { path, size };
        map[id] = entry;

        const dataId =
            dataEl.getAttribute("sfa:ID") ||
            dataEl.getAttribute("ID") ||
            dataEl.getAttribute("id");
        if (dataId) {
            map[dataId] = entry;
        }
    }

    return map;
}

async function buildImageShapeFromElement(zip, fileNames, imageBinaryIndex, imageEl, slideSize, isMasterShape = false) {
    const geometryEl = findNearestGeometry(imageEl);
    const dataEl = Array.from(imageEl.getElementsByTagName("*")).find((c) => c.localName === "data");
    const refEl = Array.from(imageEl.getElementsByTagName("*")).find(
        (c) => c.localName === "unfiltered-ref" || c.localName === "data-ref"
    );

    const pathAttr =
        dataEl?.getAttribute("sf:path") ||
        dataEl?.getAttribute("path") ||
        dataEl?.getAttribute("sf:displayname") ||
        dataEl?.getAttribute("displayname");
    const refId = getIdRef(refEl);
    const refEntry = refId ? imageBinaryIndex?.[refId] : null;
    const pickedPath = pathAttr || refEntry?.path;
    if (!pickedPath) return null;

    const file = findAssetFile(zip, fileNames, pickedPath);
    if (!file) return null;

    const bytes = await file.async("uint8array");
    let mime = null;
    try {
        mime = guessMimeFromBytes(pickedPath, bytes);
    } catch {
        // ignore
    }
    if (!mime) {
        const lower = pickedPath.toLowerCase();
        if (lower.endsWith(".png")) mime = "image/png";
        else if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) mime = "image/jpeg";
        else if (lower.endsWith(".tif") || lower.endsWith(".tiff")) mime = "image/tiff";
        else mime = "image/octet-stream";
    }

    const dims = getImageDimensions(bytes, mime) || refEntry?.size || slideSize;
    const boxFromGeom = getGeometryBox(geometryEl, dims);
    const dataUrl = `data:${mime};base64,${uint8ToBase64(bytes)}`;

    return {
        type: "image",
        box: boxFromGeom,
        src: dataUrl,
        mime,
        originalPath: pickedPath,
        isMaster: isMasterShape
    };
}

function extractCellText(cell) {
    if (!cell) return "";
    const textStorage = Array.from(cell.getElementsByTagName("*")).find(
        (c) => c.localName === "text-storage" || c.localName === "textStorage"
    );
    if (textStorage) {
        return (textStorage.textContent || "").trim();
    }
    return (cell.textContent || "").trim();
}

async function buildTableShapeFromElement(tableEl, slideSize, isMasterShape = false) {
    const geometryEl = findNearestGeometry(tableEl);
    const box = getGeometryBox(geometryEl, slideSize);

    const cellEls = Array.from(tableEl.getElementsByTagName("*")).filter(
        (el) => el.localName === "cell" || el.localName === "table-cell"
    );
    if (!cellEls.length) return null;

    let maxRow = 0;
    let maxCol = 0;
    const cells = [];

    for (const cell of cellEls) {
        const rowAttr =
            cell.getAttribute("row") ||
            cell.getAttribute("sf:row") ||
            cell.getAttribute("sfa:row");
        const colAttr =
            cell.getAttribute("column") ||
            cell.getAttribute("col") ||
            cell.getAttribute("sf:column") ||
            cell.getAttribute("sfa:column");
        const rIdx = rowAttr != null ? parseInt(rowAttr, 10) : 0;
        const cIdx = colAttr != null ? parseInt(colAttr, 10) : 0;
        maxRow = Math.max(maxRow, rIdx);
        maxCol = Math.max(maxCol, cIdx);
        const txt = extractCellText(cell);
        cells.push({ r: rIdx, c: cIdx, text: txt });
    }

    const rows = [];
    for (let r = 0; r <= maxRow; r++) {
        const row = new Array(maxCol + 1).fill("");
        rows.push(row);
    }
    for (const { r, c, text } of cells) {
        if (!rows[r]) rows[r] = [];
        rows[r][c] = text || "";
    }

    return {
        type: "table",
        box,
        data: rows,
        isMaster: isMasterShape,
        renderer: "key"
    };
}

async function extractTextShapeFromNode(node, slideSize, styleIndex, graphicStyleIndex, imageBinaryIndex, zip, fileNames, isMasterShape = false) {
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
    // Avoid painting a fill behind text; many Keynote text boxes have no visible background.
    const fill = null;
    const geom = getGeomHint(node);

    for (const p of paragraphs) {
        const baseParaStyle = resolveTextStyle(
            styleIndex,
            getStyleRefFromNode(p)
        );
        const listLevelAttr =
            p.getAttribute("sf:list-level") ||
            p.getAttribute("list-level") ||
            p.getAttribute("sfa:list-level");
        const listLevel = listLevelAttr ? parseInt(listLevelAttr, 10) : 0;

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
                const href =
                    r.getAttribute("href") ||
                    r.getAttribute("url") ||
                    r.getAttribute("sf:href") ||
                    r.getAttribute("sf:url");
                runs.push({
                    text: txt,
                    style: styleToViewerRunStyle(merged),
                    href: href || null
                });
            }
        } else {
            const txt = (p.textContent || "").replace(/[ \t]+/g, " ").replace(/\u00a0/g, " ").trim();
            if (txt) {
                const merged = baseParaStyle;
                runs.push({
                    text: txt,
                    style: styleToViewerRunStyle(merged),
                    href: null
                });
            }
        }

        if (!runs.length) continue;

        const align = baseParaStyle.align || "left";

        paraData.push({
            align,
            runs,
            bullet: listLevel > 0 ? { type: "char", char: "•" } : null,
            level: listLevel,
            marL: listLevel > 0 ? listLevel * 18 : 0,
            indent: listLevel > 0 ? listLevel * 6 : 0
        });
    }

    if (!paraData.length) return null;

    const combinedText = paraData.map((p) => p.runs.map((r) => r.text).join(" ")).join(" ").trim().toLowerCase();
    const isPlaceholderText =
        !combinedText ||
        combinedText.startsWith("body level") ||
        combinedText === "title text" ||
        combinedText === "subtitle" ||
        combinedText === "click to add title" ||
        combinedText === "click to add subtitle";

    if (isPlaceholderText || (isMasterShape && hasPlaceholderAncestor(node))) {
        return null;
    }

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

async function extractVectorShapeFromNode(node, slideSize, graphicStyleIndex, imageBinaryIndex, zip, fileNames, isMasterShape = false) {
    const all = Array.from(node.getElementsByTagName("*"));

    let styleEl = null;

    for (const el of all) {
        if (!styleEl && el.localName === "style") styleEl = el;
        if (styleEl) break;
    }

    const fill = await extractFillFromNode(styleEl || node, graphicStyleIndex, imageBinaryIndex, zip, fileNames);
    const stroke = extractStrokeFromNode(styleEl || node);
    if ((!fill || fill.type === "none") && !stroke) return null;

    const geometryEl = findNearestGeometry(node);
    const box = getGeometryBox(geometryEl, slideSize);

    const geom = getGeomHint(node);

    return {
        type: "shape",
        box,
        fill,
        stroke,
        geom,
        isMaster: isMasterShape,
        renderer: "key"
    };
}

async function collectShapesForSlide(
    slideNode,
    zip,
    fileNames,
    imageBinaryIndex,
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
                buildImageShapeFromElement(zip, fileNames, imageBinaryIndex, node, slideSize, markAsMaster)
            );
            return;
        }

        if (ln === "media" || ln === "image-media") {
            imagePromises.push(
                buildImageShapeFromElement(zip, fileNames, imageBinaryIndex, node, slideSize, markAsMaster)
            );
            // Avoid double-building the nested image-media if the parent media already covered it.
            for (const child of Array.from(node.children)) {
                if (child.localName === "image-media") continue;
                walk(child);
            }
            return;
        }

        if (ln === "table") {
            imagePromises.push(
                Promise.resolve(buildTableShapeFromElement(node, slideSize, markAsMaster))
            );
            return;
        }

        if (ln === "text" || ln === "drawable-shape" || ln === "shape") {
            if (markAsMaster && hasPlaceholderAncestor(node)) {
                return;
            }
            const textShapePromise = extractTextShapeFromNode(
                node,
                slideSize,
                styleIndex,
                graphicStyleIndex,
                imageBinaryIndex,
                zip,
                fileNames,
                markAsMaster
            );
            imagePromises.push(
                textShapePromise.then((textShape) => {
                    if (textShape) return textShape;
                    if (ln === "shape") {
                        return extractVectorShapeFromNode(
                            node,
                            slideSize,
                            graphicStyleIndex,
                            imageBinaryIndex,
                            zip,
                            fileNames,
                            markAsMaster
                        );
                    }
                    return null;
                })
            );
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

function pickBackgroundFromShapes(shapes, slideSize) {
    const areaSlide = slideSize.cx * slideSize.cy;
    let best = null;
    let bestArea = 0;

    for (const s of shapes) {
        const isFilledText = s.type === "text" && s.fill && s.fill.type === "solid";
        const isImageFill = s.type === "shape" && s.fill && s.fill.type === "image";
        if (s.type !== "shape" && s.type !== "image" && !isFilledText && !isImageFill) continue;
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

    let background = { color: "#ffffff" };
    if (best.type === "shape" && best.fill) {
        if (best.fill.type === "solid") {
            background = { color: best.fill.color };
        } else if (best.fill.type === "image") {
            background = { color: "#ffffff", image: best.fill.src };
        } else if (best.fill.type === "gradient") {
            background = { color: best.fill.colors?.[0] ?? "#ffffff", gradient: best.fill.colors };
        }
    } else if (best.type === "text" && best.fill && best.fill.type === "solid") {
        background = { color: best.fill.color };
    } else if (best.type === "image" && best.src) {
        background = { color: "#ffffff", image: best.src };
    }

    return {
        background,
        shapes: ordered
    };
}

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
    const list = Number.isFinite(maxSlides) ? imageNames.slice(0, maxSlides) : imageNames;
    for (const name of list) {
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

export async function renderKeySlides(base64, maxSlides = Infinity) {
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
        const imageBinaryIndex = buildImageBinaryIndex(doc);
        const slideStyleIndex = buildSlideStyleIndex(doc, imageBinaryIndex);
        const styleIndex = buildTextStyleIndex(doc);
        const graphicStyleIndex = buildGraphicStyleIndex(doc, imageBinaryIndex);

        const masterShapeMap = Object.create(null);
        const masterBackgroundMap = Object.create(null);
        const masterNodeMap = Object.create(null);

        function shapeTextContent(shape) {
            if (!shape || !shape.textData) return "";
            return shape.textData.paragraphs
                .map((p) => p.runs.map((r) => r.text || "").join(" "))
                .join(" ")
                .toLowerCase();
        }

        for (const masterNode of masterNodes) {
            const mid = getNodeId(masterNode);
            if (mid) masterNodeMap[mid] = masterNode;

            const mShapes = await collectShapesForSlide(
                masterNode,
                zip,
                fileNames,
                imageBinaryIndex,
                slideSize,
                styleIndex,
                graphicStyleIndex,
                true
            );
            if (mid && mShapes.length) masterShapeMap[mid] = mShapes;

            const mBg = await getSlideBackground(masterNode, graphicStyleIndex, slideStyleIndex, imageBinaryIndex, zip, fileNames);
            if (mid) masterBackgroundMap[mid] = mBg;
        }

        const slides = [];
        for (let sIdx = 0; sIdx < slideNodes.length; sIdx++) {
            const slideNode = slideNodes[sIdx];
            if (Number.isFinite(maxSlides) && slides.length >= maxSlides) break;

            const masterRef = getMasterRefId(slideNode);
            const inheritedShapes = masterRef && masterShapeMap[masterRef] ? masterShapeMap[masterRef] : [];

            const slideShapes = await collectShapesForSlide(
                slideNode,
                zip,
                fileNames,
                imageBinaryIndex,
                slideSize,
                styleIndex,
                graphicStyleIndex
            );
            let combinedShapes = [...inheritedShapes, ...slideShapes];
            if (sIdx === 0) {
                combinedShapes = combinedShapes.filter((sh) => {
                    const txt = shapeTextContent(sh);
                    if (!txt) return true;
                    return !txt.includes("keynotetemplate");
                });
            }
            if (!combinedShapes.length) continue;

            let background = await getSlideBackground(slideNode, graphicStyleIndex, slideStyleIndex, imageBinaryIndex, zip, fileNames);
            if (!background?.found && masterRef && masterBackgroundMap[masterRef]) {
                background = masterBackgroundMap[masterRef];
            }

            if (!background || !background.found) {
                const picked = pickBackgroundFromShapes(combinedShapes, slideSize);
                background = picked.background;
                combinedShapes = picked.shapes;
            } else if (!background.image && !background.gradient) {
                const picked = pickBackgroundFromShapes(combinedShapes, slideSize);
                if (picked.background.image || picked.background.gradient) {
                    background = picked.background;
                    combinedShapes = picked.shapes;
                }
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
