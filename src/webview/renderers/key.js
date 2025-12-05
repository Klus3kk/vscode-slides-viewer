import { decodeBase64ToUint8, parseXml, guessMimeFromBytes, uint8ToBase64, extractTrailingNumber, getImageDimensions } from "../utils.js";

/**
 * Minimal Keynote support: uses embedded thumbnails when available.
 */
export async function renderKeySlides(base64, maxSlides = 20) {
    try {
        const buffer = decodeBase64ToUint8(base64);
        const zip = await JSZip.loadAsync(buffer);
        const slides = [];
        const fileNames = Object.keys(zip.files);

        const loadImageSlide = async (name, fallbackSize = { cx: 1280, cy: 720 }) => {
            const file = zip.file(name);
            if (!file) return null;
            const bytes = await file.async("uint8array");
            const mime = guessMimeFromBytes(name, bytes);
            const dims = getImageDimensions(bytes, mime) || fallbackSize;
            const dataUrl = `data:${mime};base64,${uint8ToBase64(bytes)}`;
            return {
                path: name,
                size: dims,
                shapes: [
                    {
                        type: "image",
                        box: { x: 0, y: 0, cx: dims.cx, cy: dims.cy },
                        src: dataUrl,
                        mime,
                        isMaster: false
                    }
                ]
            };
        };

        const indexXml = await zip.file("index.apxl")?.async("text");
        if (indexXml) {
            const doc = parseXml(indexXml);
            if (doc) {
                let defaultSize = { cx: 1280, cy: 720 };
                const sizeEl = Array.from(doc.getElementsByTagName("*")).find((el) => el.localName === "size");
                const wAttr = sizeEl?.getAttribute("sfa:w");
                const hAttr = sizeEl?.getAttribute("sfa:h");
                if (wAttr && hAttr) {
                    const w = parseFloat(wAttr);
                    const h = parseFloat(hAttr);
                    if (w > 0 && h > 0) defaultSize = { cx: Math.round(w), cy: Math.round(h) };
                }

                const slideNodes = Array.from(doc.getElementsByTagName("*")).filter((el) => el.localName === "slide");
                for (const slideEl of slideNodes) {
                    const dataEl = Array.from(slideEl.getElementsByTagName("*")).find((el) => el.localName === "data");
                    const thumbPath = dataEl?.getAttribute("sf:path") || dataEl?.getAttribute("sf:displayname");
                    if (!thumbPath) continue;

                    const normalized = thumbPath.replace(/^\.?\//, "");
                    const candidateNames = [normalized, thumbPath];
                    let slideObj = null;
                    for (const cand of candidateNames) {
                        slideObj = await loadImageSlide(cand, defaultSize);
                        if (slideObj) break;
                    }
                    if (slideObj) {
                        slides.push(slideObj);
                        if (slides.length >= maxSlides) break;
                    }
                }
            }
        }

        if (slides.length === 0) {
            const thumbCandidates = fileNames.filter((name) => {
                const lower = name.toLowerCase();
                return (
                    lower.includes("thumbs/") &&
                    /st\\d+/i.test(lower) &&
                    !zip.files[name].dir &&
                    (lower.endsWith(".jpg") || lower.endsWith(".jpeg") || lower.endsWith(".png"))
                );
            });
            thumbCandidates.sort((a, b) => extractTrailingNumber(a) - extractTrailingNumber(b));

            for (const name of thumbCandidates.slice(0, maxSlides)) {
                const slideObj = await loadImageSlide(name);
                if (slideObj) slides.push(slideObj);
            }
        }

        if (slides.length === 0) {
            for (const name of fileNames) {
                const lower = name.toLowerCase();
                if (zip.files[name].dir) continue;
                if (!lower.includes("quicklook/thumbnail")) continue;
                if (!lower.endsWith(".jpg") && !lower.endsWith(".jpeg") && !lower.endsWith(".png")) continue;
                const slideObj = await loadImageSlide(name);
                if (slideObj) {
                    slides.push(slideObj);
                    break;
                }
            }
        }

        return slides;
    } catch (e) {
        console.error("Error rendering .key:", e);
        return [];
    }
}
