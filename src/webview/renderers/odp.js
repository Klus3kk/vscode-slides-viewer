// DONE
import { parseXml, mergeStyles, lengthToPx, getOdpStyle, guessMimeFromBytes, uint8ToBase64, decodeBase64ToUint8 } from "../utils.js";

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

async function frameToShapes(frame, allStyles, allListStyles, zip, options = {}) {
    const shapes = [];
    const isMaster = options.isMaster ?? false;
    const isPlaceholder = frame.getAttribute("presentation:placeholder") === "true";

    if (isMaster && isPlaceholder) {
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
            background: backgroundColor ? { color: backgroundColor } : null,
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

export async function renderOdpSlides(base64, maxSlides = Infinity) {
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

    const list = Number.isFinite(maxSlides) ? pages.slice(0, maxSlides) : pages;
    for (const page of list) {
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
        // if (master?.shapes?.length) { // this one line costed me 2h of my life :))
        //     shapes.push(...master.shapes);
        // }

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
        if (!backgroundColor && master?.background?.color) {
            backgroundColor = master.background.color;
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
