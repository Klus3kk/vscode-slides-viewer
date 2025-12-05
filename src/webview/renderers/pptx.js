// DONE
import { decodeBase64ToUint8, parseXml, mergeStyles } from "../utils.js";

function getPlaceholderType(shapeNode) {
    const nvSpPr = Array.from(shapeNode.children).find((el) => el.localName === "nvSpPr");
    const nvPr = nvSpPr ? Array.from(nvSpPr.children).find((el) => el.localName === "nvPr") : undefined;
    const ph = nvPr ? Array.from(nvPr.children).find((el) => el.localName === "ph") : undefined;
    return ph?.getAttribute("type") || null;
}

function parseRPrStyle(rPr) {
    const style = {};
    if (!rPr) return style;

    const sz = rPr.getAttribute("sz");
    if (sz) style.fontSize = `${parseInt(sz, 10) / 125}pt`;

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

function getColorFromXml(element) {
    const srgbClr = Array.from(element.getElementsByTagName("*")).find((el) => el.localName === "srgbClr");
    if (srgbClr) {
        const val = srgbClr.getAttribute("val");
        if (val) return `#${val}`;
    }
    return null;
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

function extractTextFromShape(shapeNode) {
    try {
        const txBody = Array.from(shapeNode.children).find((el) => el.localName === "txBody");
        if (!txBody) {
            return null;
        }

        const placeholderType = getPlaceholderType(shapeNode);
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

                if (!style.fontSize) {
                    style.fontSize = placeholderType === "title" || placeholderType === "ctrTitle" ? "44pt" : "28pt";
                }
                if (!style.fontWeight && (placeholderType === "title" || placeholderType === "ctrTitle")) {
                    style.fontWeight = "bold";
                }

                const tNodes = Array.from(r.getElementsByTagName("*")).filter((el) => el.localName === "t");
                const text = tNodes.map((t) => t.textContent || "").join("");

                if (text) runData.push({ text, style });
            }

            if (runData.length > 0) {
                textData.push({ align, runs: runData, bullet, level, marL, indent });
            }
        }

        return textData.length > 0 ? { paragraphs: textData, verticalAlign } : null;
    } catch (e) {
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
        return { cx: 10080625, cy: 5670550 };
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
        return [];
    }
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

async function parseMasterShapes(zip, masterPath) {
    try {
        if (!masterPath) return [];

        const masterXml = await zip.file(masterPath)?.async("text");
        if (!masterXml) {
            return [];
        }

        const doc = parseXml(masterXml);
        if (!doc) {
            return [];
        }

        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        if (!spTree) {
            return [];
        }

        const shapes = [];
        const spElements = Array.from(spTree.children).filter((el) => el.localName === "sp");

        for (const sp of spElements) {
            const nvSpPr = Array.from(sp.children).find((el) => el.localName === "nvSpPr");
            let isPlaceholder = false;

            if (nvSpPr) {
                const nvPr = Array.from(nvSpPr.children).find((el) => el.localName === "nvPr");
                if (nvPr) {
                    const ph = Array.from(nvPr.children).find((el) => el.localName === "ph");
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
            const textData = extractTextFromShape(sp);

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

        return shapes;
    } catch (e) {
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
            return [];
        }

        const doc = parseXml(xml);
        if (!doc) {
            return [];
        }

        const rels = await getSlideRelationships(zip, slidePath);

        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        if (!spTree) {
            return [];
        }

        const shapes = [];
        const spElements = Array.from(spTree.children).filter((el) => el.localName === "sp" || el.localName === "pic");

        for (const node of spElements) {
            const box = getShapeBox(node);
            if (!box) {
                continue;
            }

            if (node.localName === "sp") {
                const spPr = Array.from(node.children).find((el) => el.localName === "spPr");
                const fill = getShapeFill(spPr);
                const geom = getShapeGeometry(spPr);
                const textData = extractTextFromShape(node);

                if (textData || fill) {
                    shapes.push({
                        type: "text",
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
                }
            }
        }

        return shapes;
    } catch (e) {
        return [];
    }
}

export async function renderPptxSlides(base64, maxSlides = 20) {
    const buffer = decodeBase64ToUint8(base64);
    const zip = await JSZip.loadAsync(buffer);

    const slideSize = await getSlideSize(zip);
    const slidePaths = (await getSlideOrder(zip)).slice(0, maxSlides);
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
        }
    }

    return slides;
}
