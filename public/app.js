// ========================================================
// UNIFIED SEO & WEB TOOLS SUITE — CONTROLLER
// ========================================================

const API_BASE = ""; // Relative path for Vercel / local API

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


// ========================================================
// 1. TOOL: META TAG SCRAPER
// ========================================================
let currentMetaResults = [];

const metaInput = document.getElementById("meta-url-input");
const metaCount = document.getElementById("meta-url-count");

metaInput.addEventListener("input", () => {
    const count = metaInput.value.split("\n").filter(u => u.trim()).length;
    metaCount.textContent = `${count} URL${count === 1 ? '' : 's'}`;
});

document.getElementById("meta-scan-btn").addEventListener("click", async () => {
    const raw = metaInput.value;
    const urls = raw.split("\n").map(u => u.trim()).filter(Boolean);
    if (urls.length === 0) return alert("Please enter at least one URL.");

    const btn = document.getElementById("meta-scan-btn");
    btn.disabled = true;
    showStatus("meta-status", `Scanning ${urls.length} URL(s)... This may take a few seconds.`, true);
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

document.getElementById("meta-clear-btn").addEventListener("click", () => {
    metaInput.value = "";
    metaCount.textContent = "0 URLs";
    document.getElementById("meta-results-container").classList.add("hidden");
    hideStatus("meta-status");
    currentMetaResults = [];
});

document.getElementById("meta-export-btn").addEventListener("click", async () => {
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

    title.textContent = `Results — ${currentMetaResults.length} page(s)`;
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
                <div class="field-content">${item.meta_title ? escapeHtml(item.meta_title) : '<span style="color:#ef4444">[Missing Title]</span>'}</div>
            </div>
            <div class="result-field" style="margin-top: 10px;">
                <div class="field-label">Meta Description ${descPill}</div>
                <div class="field-content">${item.meta_description ? escapeHtml(item.meta_description) : '<span style="color:#ef4444">[Missing Description]</span>'}</div>
            </div>
            ${item.canonical ? `
                <div class="result-field" style="margin-top: 8px;">
                    <div class="field-label">Canonical URL</div>
                    <div class="field-content" style="color:#87879e; font-size:11px;">${escapeHtml(item.canonical)}</div>
                </div>
            ` : ''}
            ${item.error ? `<div style="color:#ef4444; font-size:11px; margin-top:8px;">⚠️ Warning: ${escapeHtml(item.error)}</div>` : ''}
        `;
        list.appendChild(card);
    });

    container.classList.remove("hidden");
}


// ========================================================
// 2. TOOL: IMAGE SITEMAP GENERATOR
// ========================================================
let currentSitemapXml = "";
let currentDiscoveredImages = [];

document.getElementById("sitemap-crawl-btn").addEventListener("click", async () => {
    const url = document.getElementById("sitemap-url-input").value.trim();
    const maxPages = document.getElementById("sitemap-max-pages").value;
    if (!url) return alert("Please enter a website URL.");

    const btn = document.getElementById("sitemap-crawl-btn");
    btn.disabled = true;
    showStatus("sitemap-status", "Crawling pages and discovering images... Please wait.", true);
    document.getElementById("sitemap-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/sitemap-crawler`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "crawl", url, max_pages: maxPages })
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
        currentDiscoveredImages.slice(0, 100).forEach(img => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td><a href="${img.page_url}" target="_blank" style="color:var(--text-secondary);">${img.page_url}</a></td>
                <td><a href="${img.image_url}" target="_blank" style="color:var(--primary);">${img.image_url}</a></td>
                <td>${escapeHtml(img.alt_text)}</td>
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

document.getElementById("sitemap-copy-xml-btn").addEventListener("click", () => {
    if (!currentSitemapXml) return;
    navigator.clipboard.writeText(currentSitemapXml);
    alert("XML Sitemap copied to clipboard!");
});

document.getElementById("sitemap-download-xml-btn").addEventListener("click", () => {
    if (!currentSitemapXml) return;
    downloadText(currentSitemapXml, "image_sitemap.xml", "application/xml");
});

document.getElementById("sitemap-download-csv-btn").addEventListener("click", () => {
    if (currentDiscoveredImages.length === 0) return;
    let csv = "Page URL,Image URL,Alt Text\n";
    currentDiscoveredImages.forEach(img => {
        csv += `"${img.page_url.replace(/"/g, '""')}","${img.image_url.replace(/"/g, '""')}","${img.alt_text.replace(/"/g, '""')}"\n`;
    });
    downloadText(csv, "discovered_images.csv", "text/csv");
});

document.getElementById("sitemap-clear-btn").addEventListener("click", () => {
    document.getElementById("sitemap-url-input").value = "";
    document.getElementById("sitemap-results-container").classList.add("hidden");
    hideStatus("sitemap-status");
});


// ========================================================
// 3. TOOL: DEEP SEO & IMAGE AUDITOR
// ========================================================
document.getElementById("audit-btn").addEventListener("click", async () => {
    const url = document.getElementById("audit-url-input").value.trim();
    if (!url) return alert("Please enter a page URL to audit.");

    const btn = document.getElementById("audit-btn");
    btn.disabled = true;
    showStatus("audit-status", "Auditing page elements, images & anchors...", true);
    document.getElementById("audit-results-container").classList.add("hidden");

    try {
        const resp = await fetch(`${API_BASE}/api/seo-audit`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "audit", url })
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

        // Meta structure
        document.getElementById("audit-title").textContent = data.title || "[None]";
        document.getElementById("audit-title-pill").innerHTML = getSeoPill(data.title_length, 30, 60);

        document.getElementById("audit-desc").textContent = data.description || "[None]";
        document.getElementById("audit-desc-pill").innerHTML = getSeoPill(data.description_length, 70, 160);

        document.getElementById("audit-h1").textContent = data.h1_tags.join(" | ") || "[Missing H1]";
        document.getElementById("audit-canonical").textContent = data.canonical || "[Missing Canonical]";

        // Populate Images Table
        const tbody = document.getElementById("audit-images-table");
        tbody.innerHTML = "";
        data.images.forEach(img => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${img.index}</td>
                <td><a href="${img.src}" target="_blank" style="color:var(--primary);">${img.src}</a></td>
                <td>${img.has_alt ? escapeHtml(img.alt) : '<span class="pill pill-bad">Missing Alt</span>'}</td>
                <td>${img.has_dimensions ? img.dimensions : '<span class="pill pill-warn">Missing WxH</span>'}</td>
                <td><span class="pill ${img.is_modern_format ? 'pill-ok' : 'pill-warn'}">${img.format}</span></td>
                <td>${img.loading === 'lazy' ? '<span class="pill pill-ok">lazy</span>' : '<span class="pill pill-warn">eager</span>'}</td>
            `;
            tbody.appendChild(tr);
        });

        document.getElementById("audit-results-container").classList.remove("hidden");
        hideStatus("audit-status");
    } catch (err) {
        showStatus("audit-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});


// ========================================================
// 4. TOOL: BACKLINK OPPORTUNITY FINDER
// ========================================================
document.getElementById("backlink-btn").addEventListener("click", async () => {
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
                    <div style="color:var(--text-secondary); font-size:12px;">${escapeHtml(item.snippet)}</div>
                `;
                liveList.appendChild(card);
            });
        } else {
            liveList.innerHTML = `<div style="color:var(--text-secondary); font-size:13px; font-family:var(--font-mono);">Click the search links above to open live Google/Bing queries directly.</div>`;
        }

        document.getElementById("backlink-results-container").classList.remove("hidden");
        hideStatus("backlink-status");
    } catch (err) {
        showStatus("backlink-status", `Error: ${err.message}`);
    } finally {
        btn.disabled = false;
    }
});

document.getElementById("backlink-clear-btn").addEventListener("click", () => {
    document.getElementById("backlink-kw").value = "";
    document.getElementById("backlink-results-container").classList.add("hidden");
    hideStatus("backlink-status");
});


// ========================================================
// 5. TOOL: SCHEMA.ORG & JSON-LD GENERATOR
// ========================================================
let currentSchemaCode = "";

document.getElementById("schema-gen-btn").addEventListener("click", async () => {
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

document.getElementById("schema-copy-btn").addEventListener("click", () => {
    if (!currentSchemaCode) return;
    navigator.clipboard.writeText(currentSchemaCode);
    alert("JSON-LD code copied to clipboard!");
});

document.getElementById("schema-clear-btn").addEventListener("click", () => {
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

document.getElementById("shorts-btn").addEventListener("click", async () => {
    const query = document.getElementById("shorts-query").value.trim();
    const max_results = document.getElementById("shorts-max").value;
    if (!query) return alert("Please enter a niche or search query.");

    const btn = document.getElementById("shorts-btn");
    btn.disabled = true;
    showStatus("shorts-status", "Searching YouTube for viral shorts...", true);
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
                <td style="max-width:320px; font-weight:600;">${escapeHtml(item.title)}</td>
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

document.getElementById("shorts-export-csv-btn").addEventListener("click", () => {
    if (currentShortsResults.length === 0) return;
    let csv = "Index,Title,Channel,Views,Duration,URL\n";
    currentShortsResults.forEach((item, idx) => {
        csv += `${idx + 1},"${item.title.replace(/"/g, '""')}","${item.channel.replace(/"/g, '""')}","${item.views}","${item.duration}","${item.url}"\n`;
    });
    downloadText(csv, "viral_shorts_trends.csv", "text/csv");
});

document.getElementById("shorts-clear-btn").addEventListener("click", () => {
    document.getElementById("shorts-query").value = "";
    document.getElementById("shorts-results-container").classList.add("hidden");
    hideStatus("shorts-status");
});

// Helper
function escapeHtml(text) {
    if (!text) return "";
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
