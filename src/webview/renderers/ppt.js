import { decodeBase64ToUint8, guessMimeFromBytes, uint8ToBase64 } from "../utils.js";

const CFB = window.CFB;

export async function renderPptSlides(base64) {
    const buffer = decodeBase64ToUint8(base64);
    const cfb = CFB.read(buffer, { type: "array" });
    const pictures = extractPptPictures(cfb);

    const pptStream = findCfbStream(cfb, "PowerPoint Document");
    if (!pptStream) {
        throw new Error("Not a valid PowerPoint file - PowerPoint Document stream not found");
    }

    const streamArray = pptStream instanceof Uint8Array ? pptStream : new Uint8Array(pptStream);
    const slides = parsePptStream(streamArray, pictures);

    // Replace non-renderable image shapes with raster fallbacks when possible.
    const rasterPics = selectRasterPictures(pictures);
    if (rasterPics.length) {
        slides.forEach((slide) => {
            slide.shapes = slide.shapes.map((shape) => {
                if (shape.type === "image" && !shapeHasRenderableImage(shape)) {
                    const pic = rasterPics[0];
                    return { ...shape, src: pic.dataUrl, mime: pic.mime };
                }
                return shape;
            });
        });
    }

    // If a slide has no renderable images (or none at all), add a single full-slide raster fallback.
    if (rasterPics.length) {
        slides.forEach((slide, idx) => {
            const hasRenderable = slide.shapes.some((s) => shapeHasRenderableImage(s));
            const hasAnyImage = slide.shapes.some((s) => s.type === "image");
            if (hasRenderable) return;
            const pic = rasterPics[idx % rasterPics.length];
            slide.shapes.unshift({
                type: "image",
                box: { x: 0, y: 0, cx: slide.size.cx, cy: slide.size.cy },
                src: pic.dataUrl,
                mime: pic.mime,
                isMaster: false
            });
            // If there were non-renderable images, keep them after the background.
            if (hasAnyImage) {
                // no-op; they remain in shapes after unshift
            }
        });
    }

    return slides;
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
        if (
            !lower.includes("pictures") &&
            !lower.endsWith(".png") &&
            !lower.endsWith(".jpg") &&
            !lower.endsWith(".jpeg") &&
            !lower.endsWith(".gif") &&
            !lower.endsWith(".bmp") &&
            !lower.endsWith(".emf") &&
            !lower.endsWith(".wmf")
        ) {
            continue;
        }
        const bytes = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content);
        if (lower.includes("pictures")) {
            const found = extractImagesFromBlob(bytes, name);
            pics.push(...found);
        } else {
            const mime = guessMimeFromBytes(name, bytes);
            const dataUrl = `data:${mime};base64,${uint8ToBase64(bytes)}`;
            pics.push({ name, dataUrl, mime });
        }
    }
    return pics;
}

function isBrowserRenderableImage(mime) {
    const m = (mime || "").toLowerCase();
    return [
        "image/png",
        "image/jpeg",
        "image/jpg",
        "image/gif",
        "image/bmp",
        "image/webp",
        "image/avif"
    ].includes(m);
}

function shapeHasRenderableImage(shape) {
    if (shape.type !== "image") return false;
    if (shape.mime && isBrowserRenderableImage(shape.mime)) return true;
    if (!shape.src) return false;
    return /^data:image\/(png|jpe?g|gif|bmp|webp|avif)/i.test(shape.src);
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
        if (i + 12 < len && bytes[i] === 0x01 && bytes[i + 1] === 0x00 && bytes[i + 2] === 0x00 && bytes[i + 3] === 0x00) {
            const size = bytes[i + 4] | (bytes[i + 5] << 8) | (bytes[i + 6] << 16) | (bytes[i + 7] << 24);
            if (size > 0 && i + size <= len) {
                matches.push({ start: i, end: i + size, mime: "image/emf" });
                i += size - 1;
                continue;
            }
        }

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
            dataUrl: `data:${mime};base64,${uint8ToBase64(slice)}`,
            mime
        });
    }
    return results;
}

function selectRasterPictures(pictures) {
    return pictures
        .filter((p) => isBrowserRenderableImage(p.mime || p.dataUrl || ""))
        .sort((a, b) => (b.dataUrl?.length || 0) - (a.dataUrl?.length || 0));
}

function buildPptTextShapesFromList(texts, slideWidthPx, slideHeightPx) {
    if (!texts || texts.length === 0) return [];

    const shapes = [];
    const splitTitleBody = (allTexts) => {
        if (allTexts.length > 1) {
            return { title: allTexts[0], body: allTexts.slice(1).join("\n") };
        }
        const raw = allTexts[0] || "";
        const parts = raw.split(/\r?\n\r?\n/).map((p) => p.trim()).filter(Boolean);
        if (parts.length > 1) {
            return { title: parts[0], body: parts.slice(1).join("\n\n") };
        }
        const lines = raw.split(/\r?\n+/).map((l) => l.trim()).filter(Boolean);
        if (lines.length > 1) {
            return { title: lines[0], body: lines.slice(1).join("\n") };
        }
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
                        fontSize: "38pt",
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
                            fontSize: "14pt",
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

    while (offset < stream.length - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;

        if (recordEnd > stream.length) break;

        try {
            switch (header.recType) {
                case 0x03E8:
                    const docResult = parsePptDocument(stream, offset, recordEnd);
                    if (docResult.width) slideWidth = docResult.width;
                    if (docResult.height) slideHeight = docResult.height;
                    break;
                case 0x03EE:
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
                        const slideWidthPx = Math.max(1, Math.round(slideWidth / 9525));
                        const slideHeightPx = Math.max(1, Math.round(slideHeight / 9525));
                        slides.push({
                            path: `slide${slides.length}`,
                            size: { cx: slideWidthPx, cy: slideHeightPx },
                            shapes: []
                        });
                    }
                    break;
                case 0x0FF0:
                    const slideText = parsePptSlideListWithText(stream, offset, recordEnd);
                    if (slideText.length > 0) {
                        textBySlide.set(textBySlide.size, slideText);
                    }
                    break;
            }
        } catch {
            // swallow per-record errors to keep parsing
        }

        offset = recordEnd;
    }

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

        if (header.recType === 0x0FA0) {
            const text = readPptTextChars(stream, offset, header.recLen);
            if (text.trim()) texts.push(text);
        } else if (header.recType === 0x0FA8) {
            const text = readPptTextBytes(stream, offset, header.recLen);
            if (text.trim()) texts.push(text);
        } else if (header.recVer === 0xF && header.recLen > 0) {
            const nested = parseProgTags(stream, offset, recordEnd);
            texts.push(...nested);
        }

        offset = recordEnd;
    }

    return texts;
}

function parsePptSlide(stream, offset, endOffset, slideWidthPx, slideHeightPx) {
    const shapes = [];
    const collectedTexts = [];

    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;

        if (header.recType === 0x040C) {
            const drawingShapes = parsePptDrawing(stream, offset, recordEnd, slideWidthPx, slideHeightPx);
            shapes.push(...drawingShapes);
        } else if (header.recType === 0x1388) {
            const texts = parseProgTags(stream, offset, recordEnd);
            if (texts.length > 0) {
                collectedTexts.push(...texts);
            }
        } else if (header.recType === 0x0FF0) {
            const texts = parsePptSlideListWithText(stream, offset, recordEnd);
            if (texts.length > 0) {
                collectedTexts.push(...texts);
            }
        }

        offset = recordEnd;
    }

    return shapes;
}

function parsePptDrawing(stream, offset, endOffset, slideWidth, slideHeight) {
    const shapes = [];

    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;

        if (header.recType === 0xF002 || header.recType === 0xF003 || header.recType === 0xF004) {
            const containerShapes = parsePptShapeContainer(stream, offset, recordEnd, slideWidth, slideHeight);
            shapes.push(...containerShapes);
        } else if (header.recType === 0xF000) {
            // OfficeArtDggContainer not handled
        }

        offset = recordEnd;
    }

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

        if (propId === 0x0004) props.bounds.left = propValue;
        else if (propId === 0x0005) props.bounds.top = propValue;
        else if (propId === 0x0006) props.bounds.right = propValue;
        else if (propId === 0x0007) props.bounds.bottom = propValue;
        else if (propId === 0x0181) props.fillColor = pptColorFromInt(propValue);
        else if (propId === 0x0180) props.fillType = propValue;
        else if (propId === 0x0104) props.blipRef = propValue & 0xFFFF; 
    }

    return props;
}

function parsePptShapeContainer(stream, offset, endOffset, slideWidth, slideHeight) {
    const shapes = [];
    let shapeData = {};

    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;

        if (recordEnd > endOffset) break;

        try {
            if (header.recType === 0xF00B) {
                const props = parsePptShapeProperties(stream, offset, header.recLen);
                Object.assign(shapeData, props);
            } else if (header.recType === 0xF00D) {
                const text = parsePptClientTextbox(stream, offset, recordEnd);
                if (text) {
                    shapeData.text = text;
                }
            } else if (header.recType >= 0xF01A && header.recType <= 0xF020) {
                const blob = stream.slice(offset, recordEnd);
                const mime = guessMimeFromBytes(`blip-${header.recType.toString(16)}`, blob);
                shapeData.image = { dataUrl: `data:${mime};base64,${uint8ToBase64(blob)}`, mime };
            } else if (header.recType === 0xF002 || header.recType === 0xF003 || header.recType === 0xF004) {
                if (shapeData.bounds || shapeData.text || shapeData.image) {
                    const shape = createPptShape(shapeData, slideWidth, slideHeight);
                    if (shape) {
                        shapes.push(shape);
                    }
                    shapeData = {};
                }
                const nested = parsePptShapeContainer(stream, offset, recordEnd, slideWidth, slideHeight);
                shapes.push(...nested);
            }
        } catch {
            // ignore malformed shape data
        }

        offset = recordEnd;
    }

    if (shapeData.bounds || shapeData.text || shapeData.image) {
        const shape = createPptShape(shapeData, slideWidth, slideHeight);
        if (shape) {
            shapes.push(shape);
        }
    }

    return shapes;
}

function parsePptClientData() {
    return null;
}

function parsePptSlideListWithText(stream, offset, endOffset) {
    const texts = [];

    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;

        if (header.recType === 0x0FA0) {
            texts.push(readPptTextChars(stream, offset, header.recLen));
        } else if (header.recType === 0x0FA8) {
            texts.push(readPptTextBytes(stream, offset, header.recLen));
        }

        offset = recordEnd;
    }

    return texts;
}

function parsePptClientTextbox(stream, offset, endOffset) {
    const texts = [];

    while (offset < endOffset - 8) {
        const header = readPptRecordHeader(stream, offset);
        offset += 8;
        const recordEnd = offset + header.recLen;

        if (header.recType === 0x0FA0) {
            texts.push(readPptTextChars(stream, offset, header.recLen));
        } else if (header.recType === 0x0FA8) {
            texts.push(readPptTextBytes(stream, offset, header.recLen));
        }

        offset = recordEnd;
    }

    return texts.join("\n");
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
    const emusToPixels = (emus) => Math.round(emus / 9525);

    let box;
    if (data.bounds && (data.bounds.left || data.bounds.left === 0)) {
        const { left = 0, top = 0, right = slideWidth, bottom = slideHeight } = data.bounds;
        box = {
            x: emusToPixels(left),
            y: emusToPixels(top),
            cx: emusToPixels(right - left),
            cy: emusToPixels(bottom - top)
        };
    } else {
        const fallbackWidth = Math.round(slideWidth * 0.8);
        const fallbackHeight = Math.round(slideHeight * 0.6);
        const offsetX = Math.round((slideWidth - fallbackWidth) / 2);
        const offsetY = Math.round((slideHeight - fallbackHeight) / 2);
        box = { x: offsetX, y: offsetY, cx: fallbackWidth, cy: fallbackHeight };
    }

    if (box.cx <= 0 || box.cy <= 0) {
        box.cx = Math.max(box.cx, Math.round(slideWidth * 0.6));
        box.cy = Math.max(box.cy, Math.round(slideHeight * 0.6));
    }

    const shape = {
        type: data.text ? "text" : data.image ? "image" : "shape",
        box,
        isMaster: false
    };

    // 🔹 keep blipRef so we can attach pictures later
    if (data.blipRef) {
        shape.blipRef = data.blipRef;
    }

    if (data.fillColor && data.fillType !== 0) {
        shape.fill = { type: "solid", color: data.fillColor };
    }

    if (data.image) {
        shape.src = data.image.dataUrl;
        shape.mime = data.image.mime;
    }

    if (data.text) {
        shape.textData = {
            paragraphs: [{
                align: "left",
                runs: [{
                    text: data.text,
                    style: { fontSize: "18pt", color: "#000000" }
                }],
                level: 0,
                marL: 0,
                indent: 0
            }],
            verticalAlign: "center"
        };
    }

    return shape;
}
