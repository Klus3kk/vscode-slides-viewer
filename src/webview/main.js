import { escapeHtml } from "./utils.js";
import { renderPptxSlides } from "./renderers/pptx.js";
import { renderOdpSlides } from "./renderers/odp.js";
import { renderPptSlides } from "./renderers/ppt.js";
import { renderKeySlides } from "./renderers/key.js";

const MAX_SLIDES = 20;
const VIEW_WIDTH = 960;
let slidesCache = [];
let currentSlide = 0;
let zoom = 1;
const vscode = acquireVsCodeApi();

function needsTiffConversion(shape) {
    const mime = (shape.mime || "").toLowerCase();
    const src = shape.src || "";
    const path = (shape.originalPath || "").toLowerCase();
    return (
        mime.includes("tif") ||
        src.includes("image/tiff") ||
        path.endsWith(".tif") ||
        path.endsWith(".tiff")
    );
}

async function convertTiffImagesToPng(slides) {
    const targets = [];
    slides.forEach((slide) => {
        slide.shapes?.forEach((shape) => {
            if (shape.type === "image" && needsTiffConversion(shape)) {
                targets.push(shape);
            }
        });
    });
    if (!targets.length) return;

    await Promise.all(
        targets.map(async (shape) => {
            try {
                const resp = await fetch(shape.src);
                const blob = await resp.blob();
                const bmp = await createImageBitmap(blob);
                const canvas = document.createElement("canvas");
                canvas.width = bmp.width;
                canvas.height = bmp.height;
                const ctx = canvas.getContext("2d");
                ctx.drawImage(bmp, 0, 0);
                shape.src = canvas.toDataURL("image/png");
                shape.mime = "image/png";
            } catch (e) {
                console.warn("TIFF->PNG conversion failed", e);
            }
        })
    );
}

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

function renderSlidesToHtml(slides) {
    return slides
        .map((slide, idx) => {
            const scale = VIEW_WIDTH / slide.size.cx;
            const heightPx = Math.round(slide.size.cy * scale);
            const backgroundColor = slide.background?.color || "#ffffff";
            const backgroundImage = slide.background?.image;
            const backgroundGradient = slide.background?.gradient;
            
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
                        } else if (shape.fill.type === 'image') {
                            bgStyle = `background: url(${shape.fill.src}) center/cover no-repeat;`;
                        } else if (shape.fill.type === 'gradient' && Array.isArray(shape.fill.colors)) {
                            const g = shape.fill.colors;
                            bgStyle = `background: linear-gradient(135deg, ${g[0]}, ${g[g.length - 1]});`;
                        } else if (shape.fill.type === 'none') {
                            bgStyle = 'background: transparent;';
                        }
                    }
                    const stroke = shape.stroke;
                    
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
                        const isEmf = (shape.mime || "").includes("emf") || (shape.src || "").includes("image/emf");
                        const isWmf = (shape.mime || "").includes("wmf") || (shape.src || "").includes("image/wmf");
                        if (isEmf || isWmf) {
                            const label = guessVectorPlaceholderLabel(shape);
                            return `<div class="shape image-shape unsupported" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;border-radius:${borderRadius}px;border:1px dashed #666;display:flex;align-items:center;justify-content:center;color:#444;font-size:12px;background:linear-gradient(135deg, rgba(0,0,0,0.03), rgba(0,0,0,0.07));text-align:center;padding:4px;box-sizing:border-box;">${escapeHtml(label)}</div>`;
                        }
                        return `<img class="shape image-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;border-radius:${borderRadius}px;" src="${shape.src}" alt="" />`;
                    }

                    if (shape.type === "table" && Array.isArray(shape.data)) {
                        const tblStyle = shape.tableStyle || {};
                        const tableInline = [];
                        if (tblStyle.borderColor) {
                            tableInline.push(`border:1px solid ${tblStyle.borderColor};`);
                        }
                        const rowsHtml = shape.data
                            .map((row, idx) => {
                                const tag = idx === 0 ? "th" : "td";
                                const cells = row
                                    .map((cell) => {
                                        const cellStyle = [];
                                        if (tblStyle.cellFill) cellStyle.push(`background:${tblStyle.cellFill};`);
                                        return `<${tag}${cellStyle.length ? ` style=\"${cellStyle.join(' ')}\"` : ""}>${escapeHtml(cell || "")}</${tag}>`;
                                    })
                                    .join("");
                                return `<tr>${cells}</tr>`;
                            })
                            .join("");
                        return `<div class="shape table-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;">` +
                            `<table${tableInline.length ? ` style=\"${tableInline.join(' ')}\"` : ""}>${rowsHtml}</table>` +
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
                        const multipleParas = shape.textData.paragraphs.length > 1;
                        const isHeadline =
                            shape.textData.paragraphs.length === 1 &&
                            shape.textData.paragraphs[0].runs.length === 1 &&
                            (shape.textData.paragraphs[0].runs[0].text || "").length <= 40;
                        const hasUrlInShape = shape.textData.paragraphs.some((p) =>
                            p.runs.some((r) => {
                                const lower = (r.text || "").toLowerCase();
                                return lower.includes("http://") || lower.includes("https://") || lower.includes(".com");
                            })
                        );
                    const textHtml = shape.textData.paragraphs.map(para => {
                        const paraTextLower = para.runs.map(r => (r.text || "").toLowerCase()).join(" ");
                        const paraContainsUrl =
                            paraTextLower.includes("http://") ||
                            paraTextLower.includes("https://") ||
                            paraTextLower.includes(".com");
                            const paraForceWhite =
                                paraTextLower.includes("keynotetemplate") ||
                                paraTextLower.includes("visit") ||
                                paraTextLower.includes(".com") ||
                                paraTextLower.includes("http://") ||
                                paraTextLower.includes("https://") ||
                                paraTextLower.includes("resources");
                            const runHtml = para.runs.map(run => {
                        const styles = [];
                        if (run.style.fontSize) styles.push(`font-size: ${run.style.fontSize}`);
                        if (run.style.fontWeight) styles.push(`font-weight: ${run.style.fontWeight}`);
                        if (run.style.fontStyle) styles.push(`font-style: ${run.style.fontStyle}`);
                        const lowerText = (run.text || "").toLowerCase();
                        const runForceWhite =
                            lowerText.includes("some cool header") ||
                            lowerText.includes("keynotetemplate") ||
                            lowerText.includes("visit") ||
                            lowerText.includes(".com") ||
                            lowerText.includes("http://") ||
                            lowerText.includes("https://") ||
                            lowerText.includes("resource");
                        const forcedWhite = paraForceWhite || runForceWhite;
                        if (run.style.color || forcedWhite) styles.push(`color: ${forcedWhite ? "#ffffff" : run.style.color}`);
                        if (run.style.fontFamily) styles.push(`font-family: "${run.style.fontFamily}", sans-serif`);
                        if (run.style.textDecoration) styles.push(`text-decoration: ${run.style.textDecoration}`);
                        if (run.style.letterSpacing) styles.push(`letter-spacing: ${run.style.letterSpacing}`);
                        
                        const styleAttr = styles.length > 0 ? ` style="${styles.join('; ')}"` : '';
                        const safeText = escapeHtml(run.text).replace(/\n/g, "<br/>");
                        if (run.href) {
                            const hrefSafe = escapeHtml(run.href);
                            return `<a href="${hrefSafe}" target="_blank" rel="noopener noreferrer"${styleAttr}>${safeText}</a>`;
                        }
                        return `<span${styleAttr}>${safeText}</span>`;
                    }).join('');
                            
                            const textAlign = para.align || 'left';
                            const indentPx = Math.max(0, Math.round(((para.marL || 0) + (para.indent || 0)) * scale));
                            const paraStyles = [];
                            paraStyles.push(`text-align:${textAlign};`);
                            paraStyles.push(`padding-left:${indentPx}px;`);
                            if (para.lineHeight) paraStyles.push(`line-height:${para.lineHeight};`);
                            if (para.spaceBefore) paraStyles.push(`margin-top:${para.spaceBefore.toFixed(2)}px;`);
                            if (para.spaceAfter) paraStyles.push(`margin-bottom:${para.spaceAfter.toFixed(2)}px;`);
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
                            // Body paragraphs inside multi-paragraph blocks often carry a center alignment flag;
                            // override to left for readability and to avoid apparent leading tabs.
                            const resolvedAlign =
                                paraContainsUrl || (multipleParas && textAlign === "center") || (textAlign === "center" && paraTextLower.length > 30)
                                    ? "left"
                                    : textAlign;
                            // overwrite alignment last to keep resolvedAlign
                            paraStyles[0] = `text-align:${resolvedAlign};`;
                            return `<p class="para" style="${paraStyles.join(' ')}">${bulletHtml}<span>${runHtml}</span></p>`;
                        }).join('');
                        
                        const whiteSpace = (isHeadline || hasUrlInShape) ? "nowrap" : "normal";
                        return `<div class="shape text-shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;${bgStyle}align-items:${verticalAlign};justify-content:${verticalAlign};border-radius:${borderRadius}px;white-space:${whiteSpace};">${textHtml}</div>`;
                    } else {
                        const strokeStyle = stroke ? `border:${Math.max(1, stroke.width || 1)}px solid ${stroke.color || "#000"};` : "";
                        return `<div class="shape" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;${bgStyle}${strokeStyle}border-radius:${borderRadius}px;"></div>`;
                    }
                })
                .join("");
            
            const backgroundStyleParts = [`background:${backgroundColor};`];
            if (backgroundGradient && Array.isArray(backgroundGradient)) {
                backgroundStyleParts.push(`background: linear-gradient(135deg, ${backgroundGradient[0]}, ${backgroundGradient[backgroundGradient.length - 1]});`);
            }
            if (backgroundImage) {
                backgroundStyleParts.push(`background-image: url(${backgroundImage}); background-size: cover; background-repeat: no-repeat; background-position: center;`);
            }

            return `
                <article class="slide-frame" id="slide-${idx}" style="width:${VIEW_WIDTH}px;height:${heightPx}px;">
                    <div class="slide-canvas" style="width:${VIEW_WIDTH}px;height:${heightPx}px;${backgroundStyleParts.join("")}">
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
            const isPptxLike =
                lowerName.endsWith(".pptx") ||
                lowerName.endsWith(".pptm") ||
                lowerName.endsWith(".potx") ||
                lowerName.endsWith(".potm") ||
                lowerName.endsWith(".ppsx") ||
                lowerName.endsWith(".ppsm");
            const isPptLike =
                lowerName.endsWith(".ppt") ||
                lowerName.endsWith(".pps") ||
                lowerName.endsWith(".pot");

            if (isPptxLike) {
                const slides = await renderPptxSlides(msg.base64, MAX_SLIDES);
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
            } else if (isPptLike) {
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
            } else if (lowerName.endsWith(".key")) {
                const slides = await renderKeySlides(msg.base64);
                await convertTiffImagesToPng(slides);
                slidesCache = slides;

                if (slides.length === 0) {
                    slidesContent.innerHTML = "<p>No preview images found in .key file.</p>";
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
