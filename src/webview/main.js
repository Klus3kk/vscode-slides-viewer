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
                        const isEmf = (shape.mime || "").includes("emf") || (shape.src || "").includes("image/emf");
                        const isWmf = (shape.mime || "").includes("wmf") || (shape.src || "").includes("image/wmf");
                        if (isEmf || isWmf) {
                            const label = guessVectorPlaceholderLabel(shape);
                            return `<div class="shape image-shape unsupported" style="left:${left}px;top:${top}px;width:${width}px;height:${height}px;border-radius:${borderRadius}px;border:1px dashed #666;display:flex;align-items:center;justify-content:center;color:#444;font-size:12px;background:linear-gradient(135deg, rgba(0,0,0,0.03), rgba(0,0,0,0.07));text-align:center;padding:4px;box-sizing:border-box;">${escapeHtml(label)}</div>`;
                        }
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
                        const safeText = escapeHtml(run.text).replace(/\n/g, "<br/>");
                        return `<span${styleAttr}>${safeText}</span>`;
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
            } else if (lowerName.endsWith(".key")) {
                const slides = await renderKeySlides(msg.base64);
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
