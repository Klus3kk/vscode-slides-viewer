// DONE
import { decodeBase64ToUint8, parseXml, mergeStyles } from "../utils.js";

const EMUS_PER_PT = 12700;
const PX_PER_PT = 96 / 72;

const DEFAULT_THEME_COLORS = {
    dk1: "#000000",
    lt1: "#ffffff",
    dk2: "#1f497d",
    lt2: "#e5e5e5",
    accent1: "#4f81bd",
    accent2: "#c0504d",
    accent3: "#9bbb59",
    accent4: "#8064a2",
    accent5: "#4bacc6",
    accent6: "#f79646",
    hlink: "#0000ff",
    folHlink: "#800080",
    bg1: "#ffffff",
    bg2: "#000000",
    tx1: "#000000",
    tx2: "#ffffff"
};

async function getThemeColors(zip) {
    try {
        const themePath = Object.keys(zip.files)
            .filter((n) => n.startsWith("ppt/theme/theme") && n.endsWith(".xml"))
            .sort()[0];
        if (!themePath) return DEFAULT_THEME_COLORS;
        const xml = await zip.file(themePath)?.async("text");
        if (!xml) return DEFAULT_THEME_COLORS;
        const doc = parseXml(xml);
        if (!doc) return DEFAULT_THEME_COLORS;
        const clrScheme = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "clrScheme");
        if (!clrScheme) return DEFAULT_THEME_COLORS;
        const colors = { ...DEFAULT_THEME_COLORS };
        Array.from(clrScheme.children).forEach((child) => {
            const name = child.localName;
            const val = getColorFromXml(child, DEFAULT_THEME_COLORS);
            if (name && val) colors[name] = val;
        });
        return colors;
    } catch (e) {
        return DEFAULT_THEME_COLORS;
    }
}

function getPlaceholderType(shapeNode) {
    const nvSpPr = Array.from(shapeNode.children).find((el) => el.localName === "nvSpPr");
    const nvPr = nvSpPr ? Array.from(nvSpPr.children).find((el) => el.localName === "nvPr") : undefined;
    const ph = nvPr ? Array.from(nvPr.children).find((el) => el.localName === "ph") : undefined;
    return ph?.getAttribute("type") || null;
}

function parseRPrStyle(rPr, themeColors = DEFAULT_THEME_COLORS) {
    const style = {};
    if (!rPr) return style;

    const sz = rPr.getAttribute("sz");
    if (sz) style.fontSize = `${parseInt(sz, 10) / 125}pt`;

    const b = rPr.getAttribute("b");
    if (b === "1") style.fontWeight = "bold";

    const i = rPr.getAttribute("i");
    if (i === "1") style.fontStyle = "italic";

    const u = rPr.getAttribute("u");
    const strike = rPr.getAttribute("strike");
    if (u && u !== "none" && u !== "0") {
        style.textDecoration = "underline";
    }
    if (strike && strike !== "noStrike") {
        style.textDecoration = style.textDecoration ? `${style.textDecoration} line-through` : "line-through";
    }

    const spc = rPr.getAttribute("spc");
    if (spc) {
        const spcPx = (parseInt(spc, 10) / 100) * PX_PER_PT;
        if (!isNaN(spcPx)) {
            style.letterSpacing = `${spcPx.toFixed(2)}px`;
        }
    }

    const solidFill = Array.from(rPr.getElementsByTagName("*")).find((el) => el.localName === "solidFill");
    if (solidFill) {
        const color = getColorFromXml(solidFill, themeColors);
        if (color) style.color = color;
    }

    const latinFont = Array.from(rPr.getElementsByTagName("*")).find((el) => el.localName === "latin");
    if (latinFont) {
        const typeface = latinFont.getAttribute("typeface");
        if (typeface) style.fontFamily = typeface;
    }

    return style;
}

function applyTintShade(hex, tint, shade) {
    const pctTint = tint != null ? Math.min(Math.max(tint / 100000, 0), 1) : null;
    const pctShade = shade != null ? Math.min(Math.max(shade / 100000, 0), 1) : null;
    const toRgb = (h) => {
        const n = parseInt(h.replace("#", ""), 16);
        return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
    };
    const fromRgb = (r, g, b) => `#${[r, g, b].map((c) => c.toString(16).padStart(2, "0")).join("")}`;
    let [r, g, b] = toRgb(hex);
    if (pctTint != null) {
        r = Math.round(r + (255 - r) * pctTint);
        g = Math.round(g + (255 - g) * pctTint);
        b = Math.round(b + (255 - b) * pctTint);
    }
    if (pctShade != null) {
        r = Math.round(r * (1 - pctShade));
        g = Math.round(g * (1 - pctShade));
        b = Math.round(b * (1 - pctShade));
    }
    return fromRgb(r, g, b);
}

function applyLumModOff(hex, lumMod, lumOff) {
    if (lumMod == null && lumOff == null) return hex;
    const toRgb = (h) => {
        const n = parseInt(h.replace("#", ""), 16);
        return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
    };
    const fromRgb = (r, g, b) => `#${[r, g, b].map((c) => c.toString(16).padStart(2, "0")).join("")}`;
    const mod = lumMod != null ? lumMod / 100000 : 1;
    const off = lumOff != null ? lumOff / 100000 : 0;
    const [r, g, b] = toRgb(hex);
    const adj = (c) => Math.min(255, Math.max(0, Math.round(c * mod + 255 * off)));
    return fromRgb(adj(r), adj(g), adj(b));
}

function getColorFromXml(element, themeColors = DEFAULT_THEME_COLORS) {
    const srgbClr = Array.from(element.getElementsByTagName("*")).find((el) => el.localName === "srgbClr");
    if (srgbClr) {
        const val = srgbClr.getAttribute("val");
        let base = val ? `#${val}` : null;
        if (base) {
            const tintEl = Array.from(srgbClr.children).find((el) => el.localName === "tint");
            const shadeEl = Array.from(srgbClr.children).find((el) => el.localName === "shade");
            const lumModEl = Array.from(srgbClr.children).find((el) => el.localName === "lumMod");
            const lumOffEl = Array.from(srgbClr.children).find((el) => el.localName === "lumOff");
            const alphaEl = Array.from(srgbClr.children).find((el) => el.localName === "alpha");
            base = applyTintShade(
                base,
                tintEl ? parseInt(tintEl.getAttribute("val") || "0", 10) : null,
                shadeEl ? parseInt(shadeEl.getAttribute("val") || "0", 10) : null
            );
            base = applyLumModOff(
                base,
                lumModEl ? parseInt(lumModEl.getAttribute("val") || "100000", 10) : null,
                lumOffEl ? parseInt(lumOffEl.getAttribute("val") || "0", 10) : null
            );
            const alpha = alphaEl ? parseInt(alphaEl.getAttribute("val") || "100000", 10) : 100000;
            if (alpha < 100000) {
                const pct = Math.max(0, Math.min(alpha / 100000, 1));
                const n = parseInt(base.replace("#", ""), 16);
                const r = (n >> 16) & 255;
                const g = (n >> 8) & 255;
                const b = n & 255;
                return `rgba(${r}, ${g}, ${b}, ${pct.toFixed(3)})`;
            }
            return base;
        }
    }
    const schemeClr = Array.from(element.getElementsByTagName("*")).find((el) => el.localName === "schemeClr");
    if (schemeClr) {
        const val = schemeClr.getAttribute("val");
        let mapped = val;
        if (val === "bg1") mapped = "lt1";
        else if (val === "bg2") mapped = "dk1";
        else if (val === "tx1") mapped = "dk1";
        else if (val === "tx2") mapped = "lt1";
        let base = mapped && themeColors[mapped] ? themeColors[mapped] : null;
        if (base) {
            const tintEl = Array.from(schemeClr.children).find((el) => el.localName === "tint");
            const shadeEl = Array.from(schemeClr.children).find((el) => el.localName === "shade");
            const lumModEl = Array.from(schemeClr.children).find((el) => el.localName === "lumMod");
            const lumOffEl = Array.from(schemeClr.children).find((el) => el.localName === "lumOff");
            const alphaEl = Array.from(schemeClr.children).find((el) => el.localName === "alpha");
            const tint = tintEl ? parseInt(tintEl.getAttribute("val") || "0", 10) : null;
            const shade = shadeEl ? parseInt(shadeEl.getAttribute("val") || "0", 10) : null;
            base = applyTintShade(base, tint, shade);
            base = applyLumModOff(
                base,
                lumModEl ? parseInt(lumModEl.getAttribute("val") || "100000", 10) : null,
                lumOffEl ? parseInt(lumOffEl.getAttribute("val") || "0", 10) : null
            );
            const alpha = alphaEl ? parseInt(alphaEl.getAttribute("val") || "100000", 10) : 100000;
            if (alpha < 100000) {
                const pct = Math.max(0, Math.min(alpha / 100000, 1));
                const n = parseInt(base.replace("#", ""), 16);
                const r = (n >> 16) & 255;
                const g = (n >> 8) & 255;
                const b = n & 255;
                return `rgba(${r}, ${g}, ${b}, ${pct.toFixed(3)})`;
            }
            return base;
        }
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

function getFrameBox(frameEl) {
    const xfrm = Array.from(frameEl.children).find((el) => el.localName === "xfrm");
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

function getShapeFill(spPr, themeColors = DEFAULT_THEME_COLORS) {
    if (!spPr) return null;

    const solidFill = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "solidFill");
    if (solidFill) {
        const color = getColorFromXml(solidFill, themeColors);
        if (color) return { type: "solid", color };
    }

    const gradFill = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "gradFill");
    if (gradFill) {
        const stops = Array.from(gradFill.getElementsByTagName("*")).filter((el) => el.localName === "gs");
        const parsed = stops
            .map((s) => {
                const c = getColorFromXml(s, themeColors);
                const pos = parseInt(s.getAttribute("pos") || "0", 10);
                return c ? { color: c, pos: isNaN(pos) ? null : pos } : null;
            })
            .filter(Boolean);
        if (parsed.length) return { type: "gradient", stops: parsed };
    }

    const noFill = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "noFill");
    if (noFill) return { type: "none" };

    return null;
}

function getShapeStroke(spPr, themeColors = DEFAULT_THEME_COLORS) {
    if (!spPr) return null;
    const ln = Array.from(spPr.children).find((el) => el.localName === "ln");
    if (!ln) return null;
    const hasNoFill = Array.from(ln.children).some((el) => el.localName === "noFill");
    if (hasNoFill) return null;
    const w = ln.getAttribute("w");
    const widthPt = w ? parseInt(w, 10) / EMUS_PER_PT : 0;
    const widthPx = widthPt > 0 ? Math.max(1, Math.round(widthPt * PX_PER_PT)) : 0;
    const solidFill = Array.from(ln.getElementsByTagName("*")).find((el) => el.localName === "solidFill");
    const color = solidFill ? getColorFromXml(solidFill, themeColors) : null;
    if (!color && !widthPx) return null;
    return { width: widthPx || 1, color: color || "#000" };
}

function getShapeGeometry(spPr) {
    if (!spPr) return null;
    const prstGeom = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "prstGeom");
    return prstGeom?.getAttribute("prst") || null;
}

function getCustomGeometryPath(spPr) {
    if (!spPr) return null;
    const custGeom = Array.from(spPr.getElementsByTagName("*")).find((el) => el.localName === "custGeom");
    if (!custGeom) return null;
    const pathLst = Array.from(custGeom.children).find((el) => el.localName === "pathLst");
    if (!pathLst) return null;
    const path = Array.from(pathLst.children).find((el) => el.localName === "path");
    if (!path) return null;
    let d = "";
    const push = (s) => {
        if (d) d += " ";
        d += s;
    };
    Array.from(path.children).forEach((child) => {
        if (child.localName === "moveTo") {
            const pt = Array.from(child.children).find((el) => el.localName === "pt");
            if (pt) push(`M ${pt.getAttribute("x") || 0} ${pt.getAttribute("y") || 0}`);
        } else if (child.localName === "lnTo") {
            const pt = Array.from(child.children).find((el) => el.localName === "pt");
            if (pt) push(`L ${pt.getAttribute("x") || 0} ${pt.getAttribute("y") || 0}`);
        } else if (child.localName === "arcTo") {
            const wR = child.getAttribute("wR") || 0;
            const hR = child.getAttribute("hR") || 0;
            push(`A ${wR} ${hR} 0 0 1 ${child.getAttribute("stAng") || 0} ${child.getAttribute("swAng") || 0}`);
        } else if (child.localName === "close") {
            push("Z");
        }
    });
    return d || null;
}

function extractTextFromShape(shapeNode, rels, themeColors = DEFAULT_THEME_COLORS) {
    try {
        const txBody = Array.from(shapeNode.children).find((el) => el.localName === "txBody");
        if (!txBody) {
            return null;
        }

    const placeholderType = getPlaceholderType(shapeNode);
    const shapeDefault = parseRPrStyle(Array.from(txBody.querySelectorAll("defRPr"))[0]);

    const bodyPr = Array.from(txBody.children).find((el) => el.localName === "bodyPr");
    let verticalAlign = placeholderType === "ctrTitle" || placeholderType === "title" ? "center" : "flex-start";
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
            let lineHeight = null;
            let spaceBefore = null;
            let spaceAfter = null;
            const paraDefaults = parseRPrStyle(Array.from(pPr?.children || []).find((el) => el.localName === "defRPr"), themeColors);
            const fallbackFontSize =
                paraDefaults.fontSize ||
                shapeDefault.fontSize ||
                (placeholderType === "title" || placeholderType === "ctrTitle" ? "32pt" : "18pt");

            if (pPr) {
                const algnAttr = pPr.getAttribute("algn");
                if (algnAttr === "ctr") align = "center";
                else if (algnAttr === "r") align = "right";
                else if (algnAttr === "l") align = "left";

                marL = parseInt(pPr.getAttribute("marL") || "0", 10);
                indent = parseInt(pPr.getAttribute("indent") || "0", 10);
                const lvlAttr = pPr.getAttribute("lvl");
                if (lvlAttr) level = parseInt(lvlAttr, 10) || 0;

                const lnSpc = Array.from(pPr.children).find((el) => el.localName === "lnSpc");
                const spcPct = Array.from(lnSpc?.children || []).find((el) => el.localName === "spcPct");
                const spcPts = Array.from(lnSpc?.children || []).find((el) => el.localName === "spcPts");
                if (spcPct) {
                    const val = parseInt(spcPct.getAttribute("val") || "0", 10);
                    if (!isNaN(val) && val > 0) lineHeight = (val / 100000).toFixed(2);
                } else if (spcPts) {
                    const val = parseInt(spcPts.getAttribute("val") || "0", 10);
                    if (!isNaN(val) && val > 0) lineHeight = `${((val / 100) * PX_PER_PT).toFixed(2)}px`;
                }

                const spcBef = Array.from(pPr.children).find((el) => el.localName === "spcBef");
                const spcAft = Array.from(pPr.children).find((el) => el.localName === "spcAft");
                const befPts = Array.from(spcBef?.children || []).find((el) => el.localName === "spcPts");
                const aftPts = Array.from(spcAft?.children || []).find((el) => el.localName === "spcPts");
                if (befPts) {
                    const val = parseInt(befPts.getAttribute("val") || "0", 10);
                    if (!isNaN(val) && val > 0) spaceBefore = ((val / 100) * PX_PER_PT);
                }
                if (aftPts) {
                    const val = parseInt(aftPts.getAttribute("val") || "0", 10);
                    if (!isNaN(val) && val > 0) spaceAfter = ((val / 100) * PX_PER_PT);
                }

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
                const style = mergeStyles(shapeDefault, mergeStyles(paraDefaults, parseRPrStyle(rPr, themeColors)));

                if (!style.fontSize) {
                    style.fontSize = fallbackFontSize;
                }
                if (!style.fontWeight && (placeholderType === "title" || placeholderType === "ctrTitle")) {
                    style.fontWeight = "bold";
                }

                const tNodes = Array.from(r.getElementsByTagName("*")).filter((el) => el.localName === "t");
                const text = tNodes.map((t) => t.textContent || "").join("");

                if (text) {
                    const run = { text, style };
                    const hlink = Array.from(rPr?.children || []).find((el) => el.localName === "hlinkClick");
                    const rId =
                        hlink?.getAttribute("r:id") ||
                        hlink?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
                    if (rId && rels && rels[rId]) {
                        run.href = rels[rId];
                    }
                    runData.push(run);
                }
            }

            if (runData.length > 0) {
                const resolvedAlign =
                    align === "left" && (placeholderType === "title" || placeholderType === "ctrTitle") ? "center" : align;
                textData.push({ align: resolvedAlign, runs: runData, bullet, level, marL, indent, lineHeight, spaceBefore, spaceAfter });
            }
        }

        return textData.length > 0 ? { paragraphs: textData, verticalAlign } : null;
    } catch (e) {
        return null;
    }
}

function extractPlainTextFromTxBody(txBody) {
    if (!txBody) return "";
    const paragraphs = Array.from(txBody.getElementsByTagName("*")).filter((el) => el.localName === "p");
    const texts = paragraphs.map((p) => {
        const tNodes = Array.from(p.getElementsByTagName("*")).filter((el) => el.localName === "t");
        return tNodes.map((t) => t.textContent || "").join("");
    });
    return texts.join("\n");
}

function isPlaceholder(shapeNode) {
    const nvSpPr = Array.from(shapeNode.children).find((el) => el.localName === "nvSpPr");
    if (!nvSpPr) return false;
    const nvPr = Array.from(nvSpPr.children).find((el) => el.localName === "nvPr");
    if (!nvPr) return false;
    const ph = Array.from(nvPr.children).find((el) => el.localName === "ph");
    return !!ph;
}

function getPlaceholderInfo(shapeNode) {
    const nvSpPr = Array.from(shapeNode.children).find((el) => el.localName === "nvSpPr");
    if (!nvSpPr) return null;
    const nvPr = Array.from(nvSpPr.children).find((el) => el.localName === "nvPr");
    if (!nvPr) return null;
    const ph = Array.from(nvPr.children).find((el) => el.localName === "ph");
    if (!ph) return null;
    return {
        type: ph.getAttribute("type") || "body",
        idx: ph.getAttribute("idx") || "0",
        size: ph.getAttribute("sz") || null
    };
}

function collectPlaceholderBoxes(spTree) {
    const map = {};
    const spNodes = Array.from(spTree.children).filter((el) => el.localName === "sp");
    for (const sp of spNodes) {
        const ph = getPlaceholderInfo(sp);
        if (!ph) continue;
        const box = getShapeBox(sp);
        if (!box) continue;
        const key = `${ph.type}:${ph.idx}`;
        map[key] = box;
        map[ph.type] = map[ph.type] || box;
        map[`idx:${ph.idx}`] = map[`idx:${ph.idx}`] || box;
    }
    return map;
}

function getGroupTransform(node) {
    const grpSpPr = Array.from(node.children).find((el) => el.localName === "grpSpPr");
    const xfrm = grpSpPr ? Array.from(grpSpPr.children).find((el) => el.localName === "xfrm") : undefined;
    if (!xfrm) return null;
    const off = Array.from(xfrm.children).find((el) => el.localName === "off");
    const ext = Array.from(xfrm.children).find((el) => el.localName === "ext");
    const chOff = Array.from(xfrm.children).find((el) => el.localName === "chOff");
    const chExt = Array.from(xfrm.children).find((el) => el.localName === "chExt");
    return {
        off: {
            x: parseInt(off?.getAttribute("x") ?? "0", 10),
            y: parseInt(off?.getAttribute("y") ?? "0", 10)
        },
        ext: {
            cx: parseInt(ext?.getAttribute("cx") ?? "0", 10),
            cy: parseInt(ext?.getAttribute("cy") ?? "0", 10)
        },
        chOff: {
            x: parseInt(chOff?.getAttribute("x") ?? "0", 10),
            y: parseInt(chOff?.getAttribute("y") ?? "0", 10)
        },
        chExt: {
            cx: parseInt(chExt?.getAttribute("cx") ?? "1", 10) || 1,
            cy: parseInt(chExt?.getAttribute("cy") ?? "1", 10) || 1
        }
    };
}

function applyGroupTransform(box, transform) {
    if (!transform) return box;
    const scaleX = transform.ext.cx && transform.chExt.cx ? transform.ext.cx / transform.chExt.cx : 1;
    const scaleY = transform.ext.cy && transform.chExt.cy ? transform.ext.cy / transform.chExt.cy : 1;
    return {
        x: transform.off.x + (box.x - transform.chOff.x) * scaleX,
        y: transform.off.y + (box.y - transform.chOff.y) * scaleY,
        cx: box.cx * scaleX,
        cy: box.cy * scaleY
    };
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

async function getLayoutAndMasterPaths(zip, slidePath) {
    try {
        const slideRelsPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
        const slideRelsXml = await zip.file(slideRelsPath)?.async("text");
        let layoutPath = null;
        let masterPath = null;

        if (slideRelsXml) {
            const slideRels = buildRelationshipMap(slideRelsXml);
            const layoutRel = Object.entries(slideRels).find(([_, target]) => target.includes("slideLayout"));
            if (layoutRel) {
                layoutPath = layoutRel[1];
            }
        }

        // Fallback: take the first available slideLayout if none is linked.
        if (!layoutPath) {
            const numMatch = slidePath.match(/slide(\\d+)\\.xml$/);
            const idx = numMatch ? parseInt(numMatch[1], 10) : NaN;
            const candidateByIndex = Number.isFinite(idx)
                ? `ppt/slideLayouts/slideLayout${idx}.xml`
                : null;
            const exists = candidateByIndex && zip.file(candidateByIndex);
            if (exists) {
                layoutPath = candidateByIndex;
            } else {
                const layouts = Object.keys(zip.files).filter(
                    (n) => n.startsWith("ppt/slideLayouts/slideLayout") && n.endsWith(".xml")
                );
                layoutPath = layouts.sort()[0];
            }
        }
        if (layoutPath?.startsWith("../")) {
            layoutPath = layoutPath.replace("../", "");
        }
        if (layoutPath && !layoutPath.startsWith("ppt/")) {
            layoutPath = `ppt/${layoutPath}`;
        }

        if (!layoutPath) return { layoutPath: null, masterPath: null };

        const layoutRelsPath = layoutPath.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
        const layoutRelsXml = await zip.file(layoutRelsPath)?.async("text");

        if (layoutRelsXml) {
            const layoutRels = buildRelationshipMap(layoutRelsXml);
            const masterRel = Object.entries(layoutRels).find(([_, target]) => target.includes("slideMaster"));
            if (masterRel) {
                masterPath = masterRel[1];
            }
        }

        // Fallback: first slide master if not found.
        if (!masterPath) {
            const firstMaster = Object.keys(zip.files)
                .filter((n) => n.startsWith("ppt/slideMasters/slideMaster") && n.endsWith(".xml"))
                .sort()[0];
            masterPath = firstMaster || null;
        }
        if (masterPath?.startsWith("../")) {
            masterPath = masterPath.replace("../", "");
        }
        if (masterPath && !masterPath.startsWith("ppt/")) {
            masterPath = `ppt/${masterPath}`;
        }

        return { layoutPath, masterPath };
    } catch (e) {
        return { layoutPath: null, masterPath: null };
    }
}

async function getRelationships(zip, xmlPath) {
    try {
        const parts = xmlPath.split("/");
        const file = parts.pop();
        if (!file) return {};
        const relPath = `${parts.join("/")}/_rels/${file}.rels`;
        const relFile = zip.file(relPath);
        if (!relFile) return {};
        const relXml = await relFile.async("text");
        return buildRelationshipMap(relXml);
    } catch (e) {
        return {};
    }
}

function resolveMediaPath(currentPath, target) {
    const parts = currentPath.split("/");
    parts.pop(); // remove filename
    let cleanTarget = target;
    while (cleanTarget.startsWith("../")) {
        cleanTarget = cleanTarget.replace("../", "");
        if (parts.length) parts.pop();
    }
    return `${parts.join("/")}/${cleanTarget}`.replace(/\\/g, "/");
}

async function parseBackground(zip, doc, rels, currentPath, themeColors = DEFAULT_THEME_COLORS) {
    try {
        const bg = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "bg" || el.localName === "bgPr");
        if (!bg) return null;
        const fillLike = Array.from(bg.children).find((el) => ["bgPr", "solidFill", "gradFill", "blipFill"].includes(el.localName));
        const targetEl = fillLike?.localName === "bgPr" ? Array.from(fillLike.children)[0] : fillLike;
        if (!targetEl) return null;
        if (targetEl.localName === "solidFill") {
            const color = getColorFromXml(targetEl, themeColors);
            return color ? { color } : null;
        }
        if (targetEl.localName === "gradFill") {
            const stops = Array.from(targetEl.getElementsByTagName("*")).filter((el) => el.localName === "gs");
            const colors = stops.map((s) => getColorFromXml(s, themeColors)).filter(Boolean);
            if (colors.length) return { gradient: colors };
        }
        if (targetEl.localName === "blipFill") {
            const blip = Array.from(targetEl.getElementsByTagName("*")).find((el) => el.localName === "blip");
            const embed =
                blip?.getAttribute("r:embed") ||
                blip?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
            if (embed && rels && rels[embed]) {
                const mediaPath = resolveMediaPath(currentPath, rels[embed]);
                const mediaFile = zip.file(mediaPath);
                if (mediaFile) {
                    const ext = mediaPath.split(".").pop()?.toLowerCase();
                    const mimeTypes = {
                        png: "image/png",
                        jpg: "image/jpeg",
                        jpeg: "image/jpeg",
                        gif: "image/gif",
                        bmp: "image/bmp",
                        svg: "image/svg+xml"
                    };
                    const mime = mimeTypes[ext];
                    if (mime) {
                        const dataUrl = `data:${mime};base64,${await mediaFile.async("base64")}`;
                        return { image: dataUrl };
                    }
                }
            }
        }
        return null;
    } catch (e) {
        return null;
    }
}

async function parseMasterShapes(zip, masterPath, themeColors = DEFAULT_THEME_COLORS) {
    try {
        if (!masterPath) return [];

        const masterXml = await zip.file(masterPath)?.async("text");
        if (!masterXml) {
            return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
        }

        const doc = parseXml(masterXml);
        if (!doc) {
            return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
        }

        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        if (!spTree) {
            return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
        }

        const rels = await getRelationships(zip, masterPath);
        const placeholderBoxes = collectPlaceholderBoxes(spTree);
        const shapes = await parseShapesFromTree(spTree, rels, masterPath, zip, { isMaster: true, skipPlaceholders: true, themeColors, placeholderBoxes });
        const background = await parseBackground(zip, doc, rels, masterPath, themeColors);
        return { shapes, background, rels, placeholderBoxes };
    } catch (e) {
        return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
    }
}

function parseTable(tableNode, themeColors = DEFAULT_THEME_COLORS) {
    const rows = [];
    const trNodes = Array.from(tableNode.children).filter((el) => el.localName === "tr");
    for (const tr of trNodes) {
        const cells = [];
        const tcNodes = Array.from(tr.children).filter((el) => el.localName === "tc");
        for (const tc of tcNodes) {
            const txBody = Array.from(tc.children).find((el) => el.localName === "txBody");
            cells.push(extractPlainTextFromTxBody(txBody));
        }
        rows.push(cells);
    }

    const tblPr = Array.from(tableNode.children).find((el) => el.localName === "tblPr");
    let tableStyle = null;
    if (tblPr) {
        const fill = getShapeFill(tblPr, themeColors);
        if (fill?.type === "solid") {
            tableStyle = { borderColor: fill.color, cellFill: fill.color };
        }
    }

    return { rows, tableStyle };
}

async function parseShapesFromTree(spTree, rels, currentPath, zip, options = {}) {
    const { isMaster = false, skipPlaceholders = false, themeColors = DEFAULT_THEME_COLORS, placeholderBoxes = {} } = options;
    const shapes = [];
    const nodes = Array.from(spTree.children).filter((el) =>
        ["sp", "pic", "graphicFrame", "grpSp"].includes(el.localName)
    );

    for (const node of nodes) {
        const phInfo = getPlaceholderInfo(node);

        if (skipPlaceholders && isPlaceholder(node)) {
            continue;
        }

        if (node.localName === "grpSp") {
            const transform = getGroupTransform(node);
            const innerShapes = await parseShapesFromTree(node, rels, currentPath, zip, options);
            innerShapes.forEach((s) => {
                s.box = applyGroupTransform(s.box, transform);
                shapes.push(s);
            });
            continue;
        }

        let box = getShapeBox(node);
        if (!box && phInfo) {
            const key = `${phInfo.type}:${phInfo.idx}`;
            box = placeholderBoxes[key] || placeholderBoxes[phInfo.type] || placeholderBoxes[`idx:${phInfo.idx}`] || null;
        }
        if (!box) continue;

        if (node.localName === "sp") {
            const spPr = Array.from(node.children).find((el) => el.localName === "spPr");
            const fill = getShapeFill(spPr, themeColors);
            const stroke = getShapeStroke(spPr, themeColors);
            const geom = getShapeGeometry(spPr);
            const customPath = getCustomGeometryPath(spPr);
            const textData = extractTextFromShape(node, rels, themeColors);

            if (textData || fill || stroke || customPath) {
                shapes.push({
                    type: textData ? "text" : "shape",
                    box,
                    fill,
                    stroke,
                    textData,
                    geom,
                    customPath,
                    isMaster
                });
            }
        } else if (node.localName === "pic") {
            const blipEl = Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "blip");
            const embed =
                blipEl?.getAttribute("r:embed") ||
                blipEl?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");

            if (!embed) continue;

            const target = rels?.[embed];
            if (!target) continue;

            const mediaPath = resolveMediaPath(currentPath, target);
            const mediaFile = zip.file(mediaPath);
            if (!mediaFile) continue;

            const ext = mediaPath.split(".").pop()?.toLowerCase();
            const mimeTypes = {
                png: "image/png",
                jpg: "image/jpeg",
                jpeg: "image/jpeg",
                gif: "image/gif",
                bmp: "image/bmp",
                svg: "image/svg+xml"
            };
            const mime = mimeTypes[ext];

            if (!mime) continue;

            try {
                const dataUrl = `data:${mime};base64,${await mediaFile.async("base64")}`;
                shapes.push({
                    type: "image",
                    box,
                    src: dataUrl,
                    mime,
                    isMaster
                });
            } catch (e) {
            }
        } else if (node.localName === "graphicFrame") {
            const graphicData = Array.from(node.getElementsByTagName("*")).find((el) => el.localName === "graphicData");
            const tbl = Array.from(graphicData?.children || []).find((el) => el.localName === "tbl");
            if (tbl) {
                const { rows, tableStyle } = parseTable(tbl, themeColors);
                if (rows.length) {
                    shapes.push({
                        type: "table",
                        data: rows,
                        tableStyle,
                        box,
                        isMaster
                    });
                }
            } else {
                const uri = graphicData?.getAttribute("uri") || "";
                const isDiagram = uri.includes("diagram");
                if (isDiagram) {
                    const frameBox = getFrameBox(node) || box;
                    const drawingTarget = Object.values(rels || {}).find((t) => t.includes("diagrams/drawing"));
                    if (frameBox && drawingTarget) {
                        const diagramPath = resolveMediaPath(currentPath, drawingTarget);
                        const diagramShapes = await parseDiagramDrawing(zip, diagramPath, frameBox, themeColors);
                        shapes.push(...diagramShapes.map((s) => ({ ...s, isMaster, isDiagram: true })));
                    }
                }
            }
        }
    }

    return shapes;
}

async function parseLayoutShapes(zip, layoutPath, themeColors = DEFAULT_THEME_COLORS) {
    try {
        if (!layoutPath) return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
        const xml = await zip.file(layoutPath)?.async("text");
        if (!xml) return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
        const doc = parseXml(xml);
        if (!doc) return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
        const rels = await getRelationships(zip, layoutPath);
        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        const placeholderBoxes = spTree ? collectPlaceholderBoxes(spTree) : {};
        const shapes = spTree
            ? await parseShapesFromTree(spTree, rels, layoutPath, zip, { isMaster: true, skipPlaceholders: true, themeColors, placeholderBoxes })
            : [];
        const background = await parseBackground(zip, doc, rels, layoutPath, themeColors);
        return { shapes, background, rels, placeholderBoxes };
    } catch (e) {
        return { shapes: [], background: null, rels: {}, placeholderBoxes: {} };
    }
}

function normalizeDiagramShapes(shapes, frameBox) {
    if (!shapes.length) return [];
    const minX = Math.min(...shapes.map((s) => s.box.x));
    const minY = Math.min(...shapes.map((s) => s.box.y));
    const maxX = Math.max(...shapes.map((s) => s.box.x + s.box.cx));
    const maxY = Math.max(...shapes.map((s) => s.box.y + s.box.cy));
    const spanX = Math.max(1, maxX - minX);
    const spanY = Math.max(1, maxY - minY);
    // Keep native positioning when possible; only downscale if diagram exceeds frame.
    const scale = Math.min(1, frameBox.cx / spanX, frameBox.cy / spanY);
    const offsetX = frameBox.x - minX * scale;
    const offsetY = frameBox.y - minY * scale;

    return shapes.map((s) => ({
        ...s,
        box: {
            x: offsetX + s.box.x * scale,
            y: offsetY + s.box.y * scale,
            cx: s.box.cx * scale,
            cy: s.box.cy * scale
        }
    }));
}

async function parseDiagramDrawing(zip, diagramPath, frameBox, themeColors = DEFAULT_THEME_COLORS) {
    try {
        const xml = await zip.file(diagramPath)?.async("text");
        if (!xml) return [];
        const doc = parseXml(xml);
        if (!doc) return [];
        const spTree = Array.from(doc.getElementsByTagName("*")).find(
            (el) => el.localName === "spTree" || el.localName === "spTreeUIdx"
        );
        if (!spTree) return [];
        const shapes = await parseShapesFromTree(spTree, {}, diagramPath, zip, {
            isMaster: false,
            skipPlaceholders: false,
            themeColors
        });
        // Ensure the rightArrow background is behind boxes: sort by area descending.
        shapes.sort((a, b) => b.box.cx * b.box.cy - a.box.cx * a.box.cy);
        return normalizeDiagramShapes(shapes, frameBox);
    } catch (e) {
        return [];
    }
}

async function parseSlideShapes(zip, slidePath, rels, themeColors = DEFAULT_THEME_COLORS, placeholderBoxes = {}) {
    try {
        const xml = await zip.file(slidePath)?.async("text");
        if (!xml) {
            return { shapes: [], background: null };
        }

        const doc = parseXml(xml);
        if (!doc) {
            return { shapes: [], background: null };
        }

        const spTree = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "spTree");
        const shapes = spTree
            ? await parseShapesFromTree(spTree, rels, slidePath, zip, { isMaster: false, themeColors, placeholderBoxes })
            : [];
        const background = await parseBackground(zip, doc, rels, slidePath, themeColors);

        return { shapes, background };
    } catch (e) {
        return { shapes: [], background: null };
    }
}

export async function renderPptxSlides(base64, maxSlides = 20) {
    const buffer = decodeBase64ToUint8(base64);
    const zip = await JSZip.loadAsync(buffer);

    const themeColors = await getThemeColors(zip);

    const slideSize = await getSlideSize(zip);
    const slidePaths = (await getSlideOrder(zip)).slice(0, maxSlides);
    const slides = [];

    for (let i = 0; i < slidePaths.length; i++) {
        const slidePath = slidePaths[i];

        try {
            const { layoutPath, masterPath } = await getLayoutAndMasterPaths(zip, slidePath);
            const master = await parseMasterShapes(zip, masterPath, themeColors);
            const layout = await parseLayoutShapes(zip, layoutPath, themeColors);
            const slideRels = await getRelationships(zip, slidePath);
            const placeholderBoxes = { ...(master.placeholderBoxes || {}), ...(layout.placeholderBoxes || {}) };
            const slide = await parseSlideShapes(zip, slidePath, slideRels, themeColors, placeholderBoxes);

            const allShapes = [
                ...(master.shapes || []),
                ...(layout.shapes || []),
                ...(slide.shapes || [])
            ];

            const background = slide.background || layout.background || master.background || null;

            slides.push({
                path: slidePath,
                size: slideSize,
                shapes: allShapes,
                background
            });
        } catch (e) {
        }
    }

    return slides;
}
