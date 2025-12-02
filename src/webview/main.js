function formatBytes(bytes) {
    if (bytes === 0) {
        return "0 B";
    }

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

function decodeXml(text) {
    return text
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&apos;/g, "'")
        .replace(/&quot;/g, "\"");
}

async function renderPptxSlides(base64) {
    const buffer = decodeBase64ToUint8(base64);
    const zip = await JSZip.loadAsync(buffer);

    const slideFiles = Object.keys(zip.files)
        .filter((name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml"))
        .sort((a, b) => {
            const num = (file) => parseInt(file.match(/slide(\\d+)\\.xml/)?.[1] ?? "0", 10);
            return num(a) - num(b);
        });

    const slides = [];

    for (const file of slideFiles) {
        const xml = await zip.file(file).async("text");
        const texts = Array.from(xml.matchAll(/<a:t[^>]*>(.*?)<\\/a:t>/g)).map((m) => decodeXml(m[1]));
        slides.push({ title: texts[0] ?? "Slide", bullets: texts.slice(1) });
    }

    return slides;
}

window.addEventListener("message", async (event) => {
    const msg = event.data;

    if (msg?.type === "loadFile") {
        const name = document.getElementById("file-name");
        const meta = document.getElementById("file-meta");
        const metaContent = document.getElementById("file-meta-content");
        const slidesEl = document.getElementById("slides");
        const slidesContent = document.getElementById("slides-content");
        const base64SizeKb = (msg.base64.length / 1024).toFixed(1);

        name.textContent = msg.fileName ?? "Presentation";
        metaContent.innerHTML = `<p><strong>Size:</strong> ${formatBytes(msg.size)} (${base64SizeKb} KB base64)</p>`;
        meta.classList.remove("hidden");

        if (msg.fileName?.toLowerCase().endsWith(".pptx")) {
            try {
                const slides = await renderPptxSlides(msg.base64);
                if (slides.length === 0) {
                    slidesContent.innerHTML = "<p>No slides found in this PPTX.</p>";
                    slidesEl.classList.remove("hidden");
                    return;
                }

                slidesContent.innerHTML = slides
                    .map((slide, idx) => {
                        const bullets = slide.bullets.map((b) => `<li>${b || "&nbsp;"}</li>`).join("");
                        return `
                            <article class="slide">
                                <header><span class="badge">Slide ${idx + 1}</span> <strong>${slide.title}</strong></header>
                                <ul>${bullets || "<li>(no text)</li>"}</ul>
                            </article>
                        `;
                    })
                    .join("");
                slidesEl.classList.remove("hidden");
            } catch (err) {
                slidesContent.innerHTML = `<p>Could not render PPTX: ${String(err)}</p>`;
                slidesEl.classList.remove("hidden");
            }
        } else {
            slidesContent.innerHTML = "<p>Rendering currently supports PPTX text only.</p>";
            slidesEl.classList.remove("hidden");
        }
    }
});
