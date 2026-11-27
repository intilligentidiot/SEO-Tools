// ========================================================
// TESLA MECHANICAL DESIGNS — SEO & WEB GROWTH SUITE
// ========================================================

const API_BASE = "";

// --- Tab Switching Logic ---
const tabButtons = document.querySelectorAll(".tab-btn");
const toolPanels = document.querySelectorAll(".tool-panel");

tabButtons.forEach(btn => {
    btn.addEventListener("click", () => {
        const targetTab = btn.getAttribute("data-tab");
        
        tabButtons.forEach(b => b.classList.remove("active"));
        btn.classList.add("active");

        toolPanels.forEach(p => {
            if (p.id === `panel-${targetTab}`) {
                p.classList.remove("hidden");
            } else {
                p.classList.add("hidden");
            }
        });
    });
});

// Helper for SEO Status Pills
function getSeoPill(length, lo, hi) {
    if (length === 0) return `<span class="pill pill-bad">Missing</span>`;
    if (length < lo) return `<span class="pill pill-warn">${length} · Too Short</span>`;
    if (length > hi) return `<span class="pill pill-warn">${length} · Too Long</span>`;
    return `<span class="pill pill-ok">${length} · OK</span>`;
}

function showStatus(elemId, message, isSpinning = false) {
    const el = document.getElementById(elemId);
    if (!el) return;
    el.classList.remove("hidden");
    if (isSpinning) {
        el.innerHTML = `<span class="spinner"></span> <span>${message}</span>`;
    } else {
        el.innerHTML = `<span>${message}</span>`;
    }
}

function hideStatus(elemId) {
    const el = document.getElementById(elemId);
    if (el) el.classList.add("hidden");
}

function downloadBlob(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    a.remove();
}

function downloadText(text, filename, mimeType = "text/plain") {
    const blob = new Blob([text], { type: mimeType });
    downloadBlob(blob, filename);
}

function escapeHtml(text) {
    if (!text) return "";
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


// ========================================================
// 1. TOOL: META TAG SCRAPER
// ========================================================
let currentMetaResults = [];

const metaInput = document.getElementById("meta-url-input");
const metaCount = document.getElementById("meta-url-count");

if (metaInput) {
    metaInput.addEventListener("input", () => {
        const count = metaInput.value.split("\n").filter(u => u.trim()).length;
        metaCount.textContent = `${count} URL${count === 1 ? '' : 's'}`;
    });
}

document.getElementById("meta-scan-btn")?.addEventListener("click", async () => {
    const raw = metaInput.value;
    const urls = raw.split("\n").map(u => u.trim()).filter(Boolean);
    if (urls.length === 0) return alert("Please enter at least one URL.");

    const btn = document.getElementById("meta-scan-btn");
    btn.disabled = true;
    showStatus("meta-status", `Scanning ${urls.length} URL(s)... Please wait.`, true);
    document.getElementById("meta-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/meta-scrape`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "scrape", urls })
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();
        currentMetaResults = data.results || [];
        renderMetaResults();
        hideStatus("meta-status");
    } catch (err) {
        showStatus("meta-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

document.getElementById("meta-clear-btn")?.addEventListener("click", () => {
    metaInput.value = "";
    metaCount.textContent = "0 URLs";
    document.getElementById("meta-results-container").classList.add("hidden");
    hideStatus("meta-status");
    currentMetaResults = [];
});

document.getElementById("meta-export-btn")?.addEventListener("click", async () => {
    if (currentMetaResults.length === 0) return;
    try {
        const resp = await fetch(`${API_BASE}/api/meta-scrape`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "export", results: currentMetaResults })
        });
        if (!resp.ok) throw new Error("Export failed");
        const blob = await resp.blob();
        downloadBlob(blob, "meta_tags_report.xlsx");
    } catch (err) {
        alert("Failed to export Excel: " + err.message);
    }
});

function renderMetaResults() {
    const container = document.getElementById("meta-results-container");
    const list = document.getElementById("meta-results-list");
    const title = document.getElementById("meta-results-title");

    title.textContent = `Scanned ${currentMetaResults.length} page(s)`;
    list.innerHTML = "";

    currentMetaResults.forEach(item => {
        const card = document.createElement("div");
        card.className = "result-card";

        const titlePill = getSeoPill(item.title_length, 30, 60);
        const descPill = getSeoPill(item.description_length, 70, 160);

        card.innerHTML = `
            <a href="${item.url}" target="_blank" rel="noopener" class="result-url">🔗 ${item.url}</a>
            <div class="result-field">
                <div class="field-label">Title Tag ${titlePill}</div>
                <div class="field-content">${item.meta_title ? escapeHtml(item.meta_title) : '<span style="color:#C91D35">[Missing Title]</span>'}</div>
            </div>
            <div class="result-field" style="margin-top: 10px;">
                <div class="field-label">Meta Description ${descPill}</div>
                <div class="field-content">${item.meta_description ? escapeHtml(item.meta_description) : '<span style="color:#C91D35">[Missing Description]</span>'}</div>
            </div>
            ${item.canonical ? `
                <div class="result-field" style="margin-top: 8px;">
                    <div class="field-label">Canonical URL</div>
                    <div class="field-content" style="color:var(--brand-sky); font-size:12px;">${escapeHtml(item.canonical)}</div>
                </div>
            ` : ''}
            ${item.error ? `<div style="color:#C91D35; font-size:12px; margin-top:8px;">⚠️ Warning: ${escapeHtml(item.error)}</div>` : ''}
        `;
        list.appendChild(card);
    });

    container.classList.remove("hidden");
}


// ========================================================
// 2. TOOL: IMAGE SITEMAP GENERATOR (ENHANCED)
// ========================================================
let currentSitemapXml = "";
let currentDiscoveredImages = [];

document.getElementById("sitemap-crawl-btn")?.addEventListener("click", async () => {
    const url = document.getElementById("sitemap-url-input").value.trim();
    const crawlMode = document.getElementById("sitemap-crawl-mode")?.value || "crawler";
    const maxPages = document.getElementById("sitemap-max-pages").value;
    const webpOnly = document.getElementById("sitemap-webp-only").checked;
    const excludeCommon = document.getElementById("sitemap-exclude-common").checked;

    if (!url) return alert("Please enter a website URL.");

    const btn = document.getElementById("sitemap-crawl-btn");
    btn.disabled = true;
    showStatus("sitemap-status", "Crawling pages and discovering images... Please wait.", true);
    document.getElementById("sitemap-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/sitemap-crawler`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                action: "crawl",
                url,
                crawl_mode: crawlMode,
                max_pages: maxPages,
                webp_only: webpOnly,
                exclude_common: excludeCommon
            })
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();

        currentSitemapXml = data.xml_sitemap;
        currentDiscoveredImages = data.images || [];

        document.getElementById("sitemap-results-title").textContent = 
            `Discovered ${data.total_images} image(s) across ${data.pages_crawled} crawled page(s)`;
        document.getElementById("sitemap-xml-preview").textContent = currentSitemapXml;

        const tbody = document.getElementById("sitemap-table-body");
        tbody.innerHTML = "";
        currentDiscoveredImages.slice(0, 150).forEach(img => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td><a href="${img.page_url}" target="_blank" style="color:var(--brand-sky);">${escapeHtml(img.page_url)}</a></td>
                <td><a href="${img.image_url}" target="_blank" style="color:var(--brand-white); font-weight:500;">${escapeHtml(img.image_url)}</a></td>
                <td>${escapeHtml(img.alt_text)}</td>
                <td>${escapeHtml(img.title)}</td>
                <td><span class="pill ${img.is_webp ? 'pill-ok' : 'pill-warn'}">${img.is_webp ? 'WEBP' : 'IMG'}</span></td>
            `;
            tbody.appendChild(tr);
        });

        document.getElementById("sitemap-results-container").classList.remove("hidden");
        hideStatus("sitemap-status");
    } catch (err) {
        showStatus("sitemap-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

document.getElementById("sitemap-copy-xml-btn")?.addEventListener("click", () => {
    if (!currentSitemapXml) return;
    navigator.clipboard.writeText(currentSitemapXml).then(() => {
        alert("Google Image Sitemap XML copied to clipboard!");
    }).catch(() => {
        alert("XML copied!");
    });
});

document.getElementById("sitemap-download-xml-btn")?.addEventListener("click", () => {
    if (!currentSitemapXml) return;
    downloadText(currentSitemapXml, "image_sitemap.xml", "application/xml");
});

document.getElementById("sitemap-download-csv-btn")?.addEventListener("click", () => {
    if (currentDiscoveredImages.length === 0) return;
    let csv = "Page URL,Image URL,Alt Text,Title\n";
    currentDiscoveredImages.forEach(img => {
        csv += `"${img.page_url.replace(/"/g, '""')}","${img.image_url.replace(/"/g, '""')}","${img.alt_text.replace(/"/g, '""')}","${img.title.replace(/"/g, '""')}"\n`;
    });
    downloadText(csv, "discovered_images.csv", "text/csv");
});

document.getElementById("sitemap-clear-btn")?.addEventListener("click", () => {
    document.getElementById("sitemap-url-input").value = "";
    document.getElementById("sitemap-results-container").classList.add("hidden");
    hideStatus("sitemap-status");
});


// ========================================================
// 3. TOOL: DEEP SEO & IMAGE AUDITOR (WITH EXCEL EXPORT)
// ========================================================
let currentAuditedUrl = "";

document.getElementById("audit-btn")?.addEventListener("click", async () => {
    const url = document.getElementById("audit-url-input").value.trim();
    const htmlContent = document.getElementById("audit-html-input")?.value.trim() || "";
    if (!url && !htmlContent) return alert("Please enter a page URL or paste HTML code to audit.");

    currentAuditedUrl = url;
    const btn = document.getElementById("audit-btn");
    btn.disabled = true;
    showStatus("audit-status", "Auditing page elements, images & anchor links...", true);
    document.getElementById("audit-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/seo-audit`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "audit", url, html_content: htmlContent })
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();

        // Populate Score
        const scoreEl = document.getElementById("audit-score");
        scoreEl.textContent = `${data.score}/100`;
        scoreEl.className = `stat-value ${data.score >= 80 ? 'ok' : data.score >= 50 ? 'warn' : 'bad'}`;

        document.getElementById("audit-total-imgs").textContent = data.image_stats.total;
        document.getElementById("audit-missing-alt").textContent = data.image_stats.missing_alt;
        document.getElementById("audit-missing-dims").textContent = data.image_stats.missing_dimensions;

        document.getElementById("audit-img-count").textContent = data.image_stats.total;
        document.getElementById("audit-link-count").textContent = data.link_stats.total;

        // Meta structure
        document.getElementById("audit-title").textContent = data.title || "[None]";
        document.getElementById("audit-title-pill").innerHTML = getSeoPill(data.title_length, 30, 60);

        document.getElementById("audit-desc").textContent = data.description || "[None]";
        document.getElementById("audit-desc-pill").innerHTML = getSeoPill(data.description_length, 70, 160);

        document.getElementById("audit-h1").textContent = data.h1_tags.join(" | ") || "[Missing H1]";
        document.getElementById("audit-canonical").textContent = data.canonical || "[Missing Canonical]";

        // Populate Images Table
        const imgTbody = document.getElementById("audit-images-table");
        imgTbody.innerHTML = "";
        data.images.forEach(img => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${img.index}</td>
                <td><a href="${img.src}" target="_blank" style="color:var(--brand-sky);">${escapeHtml(img.src)}</a></td>
                <td>${img.has_alt ? escapeHtml(img.alt) : '<span class="pill pill-bad">Missing Alt</span>'}</td>
                <td>${img.has_title ? escapeHtml(img.title) : '<span class="pill pill-warn">Missing Title</span>'}</td>
                <td>${img.has_dimensions ? img.dimensions : '<span class="pill pill-bad">Missing Dimensions</span>'}</td>
                <td><span class="pill ${img.is_modern_format ? 'pill-ok' : 'pill-warn'}">${img.format}</span></td>
                <td>${img.loading === 'lazy' ? '<span class="pill pill-ok">lazy</span>' : '<span class="pill pill-warn">eager</span>'}</td>
                <td>${img.status.includes('OK') ? '<span class="pill pill-ok">OK</span>' : '<span class="pill pill-bad">Fix Needed</span>'}</td>
            `;
            imgTbody.appendChild(tr);
        });

        // Populate Anchors Table (from New folder/SEO)
        const anchorTbody = document.getElementById("audit-anchors-table");
        anchorTbody.innerHTML = "";
        data.links.forEach(lnk => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td><a href="${lnk.href}" target="_blank" style="color:var(--brand-sky);">${escapeHtml(lnk.href)}</a></td>
                <td style="font-weight:600;">${escapeHtml(lnk.text)}</td>
                <td>${lnk.title !== 'MISSING' ? escapeHtml(lnk.title) : '<span class="pill pill-bad">Missing Title</span>'}</td>
                <td><span class="pill ${lnk.status.includes('OK') ? 'pill-ok' : 'pill-bad'}">${lnk.status}</span></td>
                <td style="color:var(--brand-sky); font-size:11px;">${escapeHtml(lnk.issue_note)}</td>
            `;
            anchorTbody.appendChild(tr);
        });

        document.getElementById("audit-results-container").classList.remove("hidden");
        hideStatus("audit-status");
    } catch (err) {
        showStatus("audit-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

// View toggles: Images vs Anchors
document.getElementById("view-images-btn")?.addEventListener("click", () => {
    document.getElementById("audit-images-section").classList.remove("hidden");
    document.getElementById("audit-anchors-section").classList.add("hidden");
    document.getElementById("view-images-btn").style.background = "var(--brand-navy-core)";
    document.getElementById("view-images-btn").style.borderColor = "var(--brand-ice-blue)";
    document.getElementById("view-anchors-btn").style.background = "transparent";
    document.getElementById("view-anchors-btn").style.borderColor = "var(--brand-navy-border)";
});

document.getElementById("view-anchors-btn")?.addEventListener("click", () => {
    document.getElementById("audit-anchors-section").classList.remove("hidden");
    document.getElementById("audit-images-section").classList.add("hidden");
    document.getElementById("view-anchors-btn").style.background = "var(--brand-navy-core)";
    document.getElementById("view-anchors-btn").style.borderColor = "var(--brand-ice-blue)";
    document.getElementById("view-images-btn").style.background = "transparent";
    document.getElementById("view-images-btn").style.borderColor = "var(--brand-navy-border)";
});

// Excel Audit Download Handler (The requested missing feature from images checker)
document.getElementById("audit-export-btn")?.addEventListener("click", async () => {
    if (!currentAuditedUrl) return alert("Please run an audit first.");
    const btn = document.getElementById("audit-export-btn");
    btn.disabled = true;
    btn.textContent = "Generating Excel...";

    try {
        const resp = await fetch(`${API_BASE}/api/seo-audit`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "export_excel", url: currentAuditedUrl })
        });
        if (!resp.ok) throw new Error("Failed to generate Excel report");
        const blob = await resp.blob();
        downloadBlob(blob, `seo_audit_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (err) {
        alert("Failed to export Excel audit: " + err.message);
    } finally {
        btn.disabled = false;
        btn.textContent = "↓ Export Detailed Audit (.xlsx)";
    }
});


// ========================================================
// 4. TOOL: BACKLINK OPPORTUNITY FINDER
// ========================================================
document.getElementById("backlink-btn")?.addEventListener("click", async () => {
    const keyword = document.getElementById("backlink-kw").value.trim();
    const category = document.getElementById("backlink-cat").value;
    const country = document.getElementById("backlink-country").value;

    if (!keyword) return alert("Please enter a keyword.");

    const btn = document.getElementById("backlink-btn");
    btn.disabled = true;
    showStatus("backlink-status", "Generating footprints & scraping live opportunities...", true);
    document.getElementById("backlink-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/backlink-finder`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ keyword, category, country })
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();

        // Render footprints
        const fpList = document.getElementById("backlink-footprints-list");
        fpList.innerHTML = "";
        data.footprints.forEach(fp => {
            const card = document.createElement("div");
            card.className = "result-card";
            card.innerHTML = `
                <div style="font-family:var(--font-mono); font-size:13px; font-weight:600; margin-bottom:8px;">
                    🎯 ${escapeHtml(fp.footprint)}
                </div>
                <div class="button-group">
                    <a href="${fp.google_url}" target="_blank" class="btn secondary">Search Google ↗</a>
                    <a href="${fp.bing_url}" target="_blank" class="btn secondary">Search Bing ↗</a>
                    <a href="${fp.duck_url}" target="_blank" class="btn secondary">Search DuckDuckGo ↗</a>
                </div>
            `;
            fpList.appendChild(card);
        });

        // Render live scraped opportunities
        const liveList = document.getElementById("backlink-live-list");
        liveList.innerHTML = "";
        if (data.live_opportunities && data.live_opportunities.length > 0) {
            data.live_opportunities.forEach(item => {
                const card = document.createElement("div");
                card.className = "result-card";
                card.innerHTML = `
                    <a href="${item.url}" target="_blank" class="result-url">🔗 ${escapeHtml(item.url)}</a>
                    <div style="font-weight:600; font-size:13px; margin-bottom:4px;">${escapeHtml(item.title)}</div>
                    <div style="color:var(--brand-sky); font-size:12px;">${escapeHtml(item.snippet)}</div>
                `;
                liveList.appendChild(card);
            });
        } else {
            liveList.innerHTML = `<div style="color:var(--brand-sky); font-size:13px; font-family:var(--font-mono);">Click the direct search buttons above to explore live Google/Bing queries.</div>`;
        }

        document.getElementById("backlink-results-container").classList.remove("hidden");
        hideStatus("backlink-status");
    } catch (err) {
        showStatus("backlink-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

document.getElementById("backlink-clear-btn")?.addEventListener("click", () => {
    document.getElementById("backlink-kw").value = "";
    document.getElementById("backlink-results-container").classList.add("hidden");
    hideStatus("backlink-status");
});


// ========================================================
// 5. TOOL: SCHEMA.ORG & JSON-LD GENERATOR
// ========================================================
let currentSchemaCode = "";

document.getElementById("schema-pdf-btn")?.addEventListener("click", async () => {
    const fileInput = document.getElementById("schema-pdf-upload");
    if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
        return alert("Please select a .pdf file first.");
    }
    const btn = document.getElementById("schema-pdf-btn");
    btn.disabled = true;
    showStatus("schema-pdf-status", "Extracting metadata and text excerpt from PDF...", true);

    try {
        const formData = new FormData();
        formData.append("file", fileInput.files[0]);
        const resp = await fetch(`${API_BASE}/api/pdf-extract-schema`, {
            method: "POST",
            body: formData
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();

        document.getElementById("schema-headline").value = data.headline || "";
        document.getElementById("schema-desc").value = data.description || "";
        document.getElementById("schema-author").value = data.author || "";
        hideStatus("schema-pdf-status");
        document.getElementById("schema-gen-btn")?.click();
    } catch (err) {
        showStatus("schema-pdf-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

document.getElementById("schema-gen-btn")?.addEventListener("click", async () => {
    const schema_type = document.getElementById("schema-type").value;
    const headline = document.getElementById("schema-headline").value.trim();
    const description = document.getElementById("schema-desc").value.trim();
    const author = document.getElementById("schema-author").value.trim();
    const url = document.getElementById("schema-page-url").value.trim();
    const image = document.getElementById("schema-image-url").value.trim();

    try {
        const resp = await fetch(`${API_BASE}/api/schema-gen`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ schema_type, headline, description, author, url, image })
        });
        if (!resp.ok) throw new Error("Generation failed");
        const data = await resp.json();
        currentSchemaCode = data.json_ld_code;

        document.getElementById("schema-code-preview").textContent = currentSchemaCode;
        document.getElementById("schema-results-container").classList.remove("hidden");
    } catch (err) {
        alert("Error: " + err.message);
    }
});

document.getElementById("schema-copy-btn")?.addEventListener("click", () => {
    if (!currentSchemaCode) return;
    navigator.clipboard.writeText(currentSchemaCode);
    alert("JSON-LD code copied to clipboard!");
});

document.getElementById("schema-clear-btn")?.addEventListener("click", () => {
    document.getElementById("schema-headline").value = "";
    document.getElementById("schema-desc").value = "";
    document.getElementById("schema-author").value = "";
    document.getElementById("schema-page-url").value = "";
    document.getElementById("schema-image-url").value = "";
    document.getElementById("schema-results-container").classList.add("hidden");
});


// ========================================================
// 6. TOOL: YOUTUBE SHORTS TREND HUNTER
// ========================================================
let currentShortsResults = [];

document.getElementById("shorts-btn")?.addEventListener("click", async () => {
    const query = document.getElementById("shorts-query").value.trim();
    const max_results = document.getElementById("shorts-max").value;
    if (!query) return alert("Please enter a niche or search query.");

    const btn = document.getElementById("shorts-btn");
    btn.disabled = true;
    showStatus("shorts-status", "Hunting YouTube for viral competitor shorts...", true);
    document.getElementById("shorts-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/shorts-hunter`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ query, max_results })
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();
        currentShortsResults = data.results || [];

        document.getElementById("shorts-results-title").textContent = `Discovered ${currentShortsResults.length} Short(s)`;
        const tbody = document.getElementById("shorts-table-body");
        tbody.innerHTML = "";

        currentShortsResults.forEach((item, idx) => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td style="max-width:340px; font-weight:600; color:var(--brand-white);">${escapeHtml(item.title)}</td>
                <td>${escapeHtml(item.channel)}</td>
                <td><span class="pill pill-ok">${escapeHtml(item.views)}</span></td>
                <td>${escapeHtml(item.duration)}</td>
                <td><a href="${item.url}" target="_blank" class="btn secondary" style="padding:4px 10px; font-size:11px;">Watch ↗</a></td>
            `;
            tbody.appendChild(tr);
        });

        document.getElementById("shorts-results-container").classList.remove("hidden");
        hideStatus("shorts-status");
    } catch (err) {
        showStatus("shorts-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

document.getElementById("shorts-export-csv-btn")?.addEventListener("click", () => {
    if (currentShortsResults.length === 0) return;
    let csv = "Index,Title,Channel,Views,Duration,URL\n";
    currentShortsResults.forEach((item, idx) => {
        csv += `${idx + 1},"${item.title.replace(/"/g, '""')}","${item.channel.replace(/"/g, '""')}","${item.views}","${item.duration}","${item.url}"\n`;
    });
    downloadText(csv, "viral_shorts_trends.csv", "text/csv");
});

document.getElementById("shorts-clear-btn")?.addEventListener("click", () => {
    document.getElementById("shorts-query").value = "";
    document.getElementById("shorts-results-container").classList.add("hidden");
    hideStatus("shorts-status");
});


// ========================================================
// 7. TOOL: UNIVERSAL VIDEO DOWNLOADER (FROM UNIVERSAL_DOWNLOADER.PY)
// ========================================================
document.getElementById("video-fetch-btn")?.addEventListener("click", async () => {
    const url = document.getElementById("video-url-input").value.trim();
    if (!url) return alert("Please enter a video URL (YouTube, Instagram Reel, etc.)");

    const btn = document.getElementById("video-fetch-btn");
    btn.disabled = true;
    showStatus("video-status", "Extracting video metadata and direct media streams...", true);
    document.getElementById("video-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/video-downloader`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url })
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();

        document.getElementById("video-title").textContent = data.title;
        document.getElementById("video-channel").textContent = data.uploader;
        document.getElementById("video-duration").textContent = data.duration;
        document.getElementById("video-thumb-img").src = data.thumbnail || "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' width='240' height='140'><rect fill='%23e2e8f0' width='240' height='140'/><text fill='%2364748b' x='50%' y='50%' dominant-baseline='middle' text-anchor='middle'>No Thumbnail</text></svg>";
        
        const srcLink = document.getElementById("video-source-link");
        srcLink.href = data.source_url;

        const formatsList = document.getElementById("video-formats-list");
        formatsList.innerHTML = "";

        if (data.formats && data.formats.length > 0) {
            data.formats.forEach(fmt => {
                const link = document.createElement("a");
                link.href = fmt.url;
                link.target = "_blank";
                link.className = "btn download";
                link.style.padding = "6px 14px";
                link.style.fontSize = "12px";
                link.textContent = `↓ Download ${fmt.resolution} (${fmt.ext.toUpperCase()})`;
                formatsList.appendChild(link);
            });
        } else if (data.download_url) {
            const link = document.createElement("a");
            link.href = data.download_url;
            link.target = "_blank";
            link.className = "btn download";
            link.style.padding = "6px 14px";
            link.style.fontSize = "12px";
            link.textContent = "↓ Direct Stream Link ↗";
            formatsList.appendChild(link);
        } else {
            formatsList.innerHTML = `<span style="font-size:12px; color:var(--text-muted);">Direct download unavailable for this stream type. Click 'Open Original' above.</span>`;
        }

        document.getElementById("video-results-container").classList.remove("hidden");
        hideStatus("video-status");
    } catch (err) {
        showStatus("video-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});


// ========================================================
// 8. TOOL: FULL ON-PAGE & TECHNICAL CHECKLIST AUDITOR
// ========================================================
let currentChecklistChecks = [];
let currentChecklistUrl = "";

function renderChecklistTable(filter = "ALL") {
    const tbody = document.getElementById("checklist-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";

    const filtered = filter === "ALL"
        ? currentChecklistChecks
        : currentChecklistChecks.filter(c => c.status === filter);

    filtered.forEach((chk, idx) => {
        const pillClass = chk.status === "PASS" ? "pill-ok" : (chk.status === "WARN" ? "pill-warn" : "pill-bad");
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td style="font-weight:600; color:var(--brand-cobalt); white-space:normal;">${escapeHtml(chk.category)}</td>
            <td style="font-weight:600; white-space:normal;">${escapeHtml(chk.particular)}</td>
            <td><span class="pill ${pillClass}">${chk.status}</span></td>
            <td style="max-width:420px; white-space:normal; word-break:break-word; line-height:1.45;">${escapeHtml(chk.value)}</td>
            <td style="max-width:300px; white-space:normal; word-break:break-word; color:var(--text-muted); line-height:1.45;">${escapeHtml(chk.remark)}</td>
        `;
        tbody.appendChild(tr);
    });
}

document.getElementById("checklist-run-btn")?.addEventListener("click", async () => {
    const url = document.getElementById("checklist-url-input").value.trim();
    if (!url) return alert("Please enter a website URL to audit.");

    currentChecklistUrl = url;
    const btn = document.getElementById("checklist-run-btn");
    btn.disabled = true;
    showStatus("checklist-status", "Running 24+ site-level & <head> checks... Please wait.", true);
    document.getElementById("checklist-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/full-checklist-audit`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "audit", url })
        });
        if (!resp.ok) throw new Error(await resp.text());
        const data = await resp.json();

        currentChecklistChecks = data.checks || [];
        const scoreEl = document.getElementById("checklist-score");
        scoreEl.textContent = `${data.score}%`;
        scoreEl.className = `stat-value ${data.score >= 80 ? 'ok' : data.score >= 50 ? 'warn' : 'bad'}`;

        document.getElementById("checklist-passed").textContent = data.summary.passed;
        document.getElementById("checklist-warned").textContent = data.summary.warned;
        document.getElementById("checklist-failed").textContent = data.summary.failed;

        const psLink = document.getElementById("checklist-pagespeed-link");
        if (psLink) {
            psLink.href = `https://pagespeed.web.dev/analysis?url=${encodeURIComponent(data.url)}`;
        }

        renderChecklistTable("ALL");
        document.getElementById("checklist-results-container").classList.remove("hidden");
        hideStatus("checklist-status");
    } catch (err) {
        showStatus("checklist-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

document.querySelectorAll(".checklist-filter-btn").forEach(btn => {
    btn.addEventListener("click", () => {
        document.querySelectorAll(".checklist-filter-btn").forEach(b => {
            b.style.borderColor = "var(--border-color)";
            b.style.color = "var(--text-navy)";
        });
        btn.style.borderColor = "var(--brand-cobalt)";
        btn.style.color = "var(--brand-cobalt)";
        renderChecklistTable(btn.getAttribute("data-filter"));
    });
});

document.getElementById("checklist-export-btn")?.addEventListener("click", async () => {
    if (!currentChecklistUrl) return alert("Please run the checklist first.");
    const btn = document.getElementById("checklist-export-btn");
    btn.disabled = true;
    btn.textContent = "Generating Excel...";

    try {
        const resp = await fetch(`${API_BASE}/api/full-checklist-audit`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "export_excel", url: currentChecklistUrl })
        });
        if (!resp.ok) throw new Error("Export failed");
        const blob = await resp.blob();
        downloadBlob(blob, `full_seo_checklist_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (err) {
        alert("Failed to export Excel: " + err.message);
    } finally {
        btn.disabled = false;
        btn.textContent = "↓ Export Checklist & Old Data (.xlsx)";
    }
});

