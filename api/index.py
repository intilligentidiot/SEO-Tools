import io
import re
import csv
import json
import time
import urllib.parse
from datetime import datetime
from urllib.parse import urljoin, urlparse

from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__, static_folder="../public", static_url_path="")
CORS(app)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

# -------------------------------------------------------------
# Static Frontend Serving
# -------------------------------------------------------------
@app.route("/")
def index():
    return app.send_static_file("index.html")

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.utcnow().isoformat(),
        "tools": [
            "meta-scrape",
            "sitemap-crawler",
            "seo-audit",
            "backlink-finder",
            "schema-gen",
            "shorts-hunter"
        ]
    })


# -------------------------------------------------------------
# 1. TOOL: META TAG SCRAPER & EXPORTER
# -------------------------------------------------------------
def fetch_meta_for_url(url: str):
    if not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url

    result = {
        "url": url,
        "status_code": 0,
        "meta_title": "",
        "title_length": 0,
        "title_status": "Missing",
        "meta_description": "",
        "description_length": 0,
        "description_status": "Missing",
        "canonical": "",
        "og_title": "",
        "og_description": "",
        "og_image": "",
        "error": ""
    }

    try:
        resp = requests.get(url, headers=HEADERS, timeout=8, allow_redirects=True)
        result["status_code"] = resp.status_code

        soup = BeautifulSoup(resp.text, "html.parser")

        # Title
        title_tag = soup.find("title")
        if title_tag and title_tag.string:
            title = title_tag.get_text(strip=True)
            result["meta_title"] = title
            result["title_length"] = len(title)
            if len(title) == 0:
                result["title_status"] = "Missing"
            elif len(title) < 30:
                result["title_status"] = "Too Short"
            elif len(title) > 60:
                result["title_status"] = "Too Long"
            else:
                result["title_status"] = "OK"

        # Description
        desc_tag = soup.find("meta", attrs={"name": re.compile(r"^description$", re.I)})
        if desc_tag and desc_tag.get("content"):
            desc = desc_tag["content"].strip()
            result["meta_description"] = desc
            result["description_length"] = len(desc)
            if len(desc) == 0:
                result["description_status"] = "Missing"
            elif len(desc) < 70:
                result["description_status"] = "Too Short"
            elif len(desc) > 160:
                result["description_status"] = "Too Long"
            else:
                result["description_status"] = "OK"

        # Canonical
        canon_tag = soup.find("link", attrs={"rel": re.compile(r"^canonical$", re.I)})
        if canon_tag and canon_tag.get("href"):
            result["canonical"] = canon_tag["href"].strip()

        # OpenGraph
        og_t = soup.find("meta", property="og:title")
        if og_t: result["og_title"] = og_t.get("content", "").strip()

        og_d = soup.find("meta", property="og:description")
        if og_d: result["og_description"] = og_d.get("content", "").strip()

        og_img = soup.find("meta", property="og:image")
        if og_img: result["og_image"] = og_img.get("content", "").strip()

    except Exception as e:
        result["error"] = str(e)
        result["title_status"] = "Error"
        result["description_status"] = "Error"

    return result


@app.route("/api/meta-scrape", methods=["POST"])
def api_meta_scrape():
    data = request.get_json(silent=True) or {}
    action = data.get("action", "scrape")

    if action == "scrape":
        urls = data.get("urls", [])
        if not urls:
            return jsonify({"error": "No URLs provided"}), 400
        
        # Limit to 30 per request to respect Vercel serverless timeouts
        results = [fetch_meta_for_url(u.strip()) for u in urls[:30] if u.strip()]
        return jsonify({"results": results, "count": len(results)})

    elif action == "export":
        results = data.get("results", [])
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Meta Data Audit"

        header_fill = PatternFill("solid", start_color="1F3864")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        border = Border(left=Side(style="thin", color="CCCCCC"),
                        right=Side(style="thin", color="CCCCCC"),
                        top=Side(style="thin", color="CCCCCC"),
                        bottom=Side(style="thin", color="CCCCCC"))
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        headers = [
            ("URL", 45),
            ("Status", 12),
            ("Meta Title", 45),
            ("Title Length", 14),
            ("Title Status", 14),
            ("Meta Description", 65),
            ("Desc Length", 14),
            ("Desc Status", 14),
            ("Canonical", 40),
            ("OG Title", 35),
            ("OG Image", 35),
            ("Error", 25)
        ]

        for col_idx, (h_name, width) in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        for row_idx, r in enumerate(results, 2):
            vals = [
                r.get("url", ""),
                r.get("status_code", ""),
                r.get("meta_title", ""),
                r.get("title_length", 0),
                r.get("title_status", ""),
                r.get("meta_description", ""),
                r.get("description_length", 0),
                r.get("description_status", ""),
                r.get("canonical", ""),
                r.get("og_title", ""),
                r.get("og_image", ""),
                r.get("error", "")
            ]
            for col_idx, v in enumerate(vals, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=v)
                cell.border = border

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="meta_tags_report.xlsx"
        )

    return jsonify({"error": "Invalid action"}), 400


# -------------------------------------------------------------
# 2. TOOL: XML IMAGE SITEMAP GENERATOR & CRAWLER
# -------------------------------------------------------------
@app.route("/api/sitemap-crawler", methods=["POST"])
def api_sitemap_crawler():
    data = request.get_json(silent=True) or {}
    action = data.get("action", "crawl")
    base_url = data.get("url", "").strip()

    if not base_url:
        return jsonify({"error": "URL is required"}), 400
    if not base_url.startswith("http://") and not base_url.startswith("https://"):
        base_url = "https://" + base_url

    parsed_base = urlparse(base_url)
    domain = parsed_base.netloc
    max_pages = min(int(data.get("max_pages", 10)), 25)

    if action == "crawl":
        visited = set()
        to_visit = [base_url]
        discovered_images = []

        while to_visit and len(visited) < max_pages:
            current_url = to_visit.pop(0)
            if current_url in visited:
                continue
            visited.add(current_url)

            try:
                resp = requests.get(current_url, headers=HEADERS, timeout=6)
                if resp.status_code != 200:
                    continue
                soup = BeautifulSoup(resp.text, "html.parser")

                # Extract images
                for img in soup.find_all("img"):
                    src = img.get("src") or img.get("data-src") or ""
                    if not src:
                        srcset = img.get("srcset")
                        if srcset:
                            src = srcset.split(",")[0].strip().split(" ")[0]
                    if src and not src.startswith("data:"):
                        full_img_url = urljoin(current_url, src)
                        alt = img.get("alt", "").strip()
                        title = img.get("title", "").strip()
                        discovered_images.append({
                            "page_url": current_url,
                            "image_url": full_img_url,
                            "alt_text": alt if alt else (title if title else "N/A"),
                            "is_webp": full_img_url.lower().endswith(".webp")
                        })

                # Internal link discovery for further crawling
                if len(visited) < max_pages:
                    for a in soup.find_all("a", href=True):
                        link = urljoin(current_url, a["href"])
                        p_link = urlparse(link)
                        if p_link.netloc == domain and link not in visited and link not in to_visit:
                            # Avoid non-html assets
                            if not p_link.path.lower().endswith((".png", ".jpg", ".jpeg", ".pdf", ".zip", ".rar")):
                                to_visit.append(link)

            except Exception:
                continue

        # Group images by page for XML generation
        pages_dict = {}
        for item in discovered_images:
            pu = item["page_url"]
            if pu not in pages_dict:
                pages_dict[pu] = []
            pages_dict[pu].append(item)

        # Build Google Image Sitemap XML
        xml_lines = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"',
            '        xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">'
        ]
        for pu, imgs in pages_dict.items():
            xml_lines.append("  <url>")
            xml_lines.append(f"    <loc>{pu}</loc>")
            for img in imgs[:100]:  # up to 100 images per page
                xml_lines.append("    <image:image>")
                safe_loc = img["image_url"].replace("&", "&amp;")
                xml_lines.append(f"      <image:loc>{safe_loc}</image:loc>")
                if img["alt_text"] and img["alt_text"] != "N/A":
                    safe_title = img["alt_text"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    xml_lines.append(f"      <image:title>{safe_title}</image:title>")
                xml_lines.append("    </image:image>")
            xml_lines.append("  </url>")
        xml_lines.append("</urlset>")
        sitemap_xml = "\n".join(xml_lines)

        return jsonify({
            "pages_crawled": len(visited),
            "total_images": len(discovered_images),
            "images": discovered_images[:200],
            "xml_sitemap": sitemap_xml
        })

    return jsonify({"error": "Invalid action"}), 400


# -------------------------------------------------------------
# 3. TOOL: DEEP SEO & IMAGE AUDITOR
# -------------------------------------------------------------
@app.route("/api/seo-audit", methods=["POST"])
def api_seo_audit():
    data = request.get_json(silent=True) or {}
    action = data.get("action", "audit")
    url = data.get("url", "").strip()

    if not url:
        return jsonify({"error": "URL is required"}), 400
    if not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url

    if action == "audit":
        try:
            resp = requests.get(url, headers=HEADERS, timeout=10)
            soup = BeautifulSoup(resp.text, "html.parser")

            # Document Title & Meta
            title_tag = soup.find("title")
            title = title_tag.get_text(strip=True) if title_tag else ""

            desc_tag = soup.find("meta", attrs={"name": re.compile(r"^description$", re.I)})
            desc = desc_tag.get("content", "").strip() if desc_tag else ""

            viewport_tag = soup.find("meta", attrs={"name": "viewport"})
            has_viewport = bool(viewport_tag)

            canonical_tag = soup.find("link", attrs={"rel": "canonical"})
            canonical = canonical_tag.get("href", "").strip() if canonical_tag else ""

            h1_tags = [h.get_text(strip=True) for h in soup.find_all("h1")]
            h2_tags = [h.get_text(strip=True) for h in soup.find_all("h2")]

            # Image Audit
            images_data = []
            missing_alt_count = 0
            missing_dimensions_count = 0
            improper_format_count = 0
            lazy_loaded_count = 0

            PERMISSIBLE_FORMATS = {"webp", "svg", "gif"}

            all_imgs = soup.find_all("img")
            for idx, img in enumerate(all_imgs, 1):
                src = img.get("src") or img.get("data-src") or ""
                full_src = urljoin(url, src) if src else ""
                alt = img.get("alt")
                width = img.get("width")
                height = img.get("height")
                loading = img.get("loading", "eager").lower()

                # Extension check
                clean_path = urlparse(full_src).path.lower()
                ext = clean_path.split(".")[-1] if "." in clean_path else "unknown"

                is_alt_missing = (alt is None or alt.strip() == "")
                is_dim_missing = (not width or not height)
                is_format_warning = (ext not in PERMISSIBLE_FORMATS and ext in {"jpg", "jpeg", "png"})

                if is_alt_missing: missing_alt_count += 1
                if is_dim_missing: missing_dimensions_count += 1
                if is_format_warning: improper_format_count += 1
                if loading == "lazy": lazy_loaded_count += 1

                images_data.append({
                    "index": idx,
                    "src": full_src,
                    "alt": alt.strip() if alt else "[MISSING]",
                    "has_alt": not is_alt_missing,
                    "dimensions": f"{width}x{height}" if not is_dim_missing else "Missing",
                    "has_dimensions": not is_dim_missing,
                    "format": ext.upper(),
                    "is_modern_format": ext in PERMISSIBLE_FORMATS,
                    "loading": loading
                })

            # Links Audit
            links_data = []
            missing_anchor_titles = 0
            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                title_attr = a.get("title", "").strip()
                anchor_text = a.get_text(strip=True)
                if not title_attr:
                    missing_anchor_titles += 1
                links_data.append({
                    "href": urljoin(url, href),
                    "text": anchor_text[:40],
                    "title": title_attr if title_attr else "[MISSING]"
                })

            # Calculate Health Score (out of 100)
            score = 100
            if not title: score -= 15
            elif len(title) < 30 or len(title) > 65: score -= 5

            if not desc: score -= 15
            elif len(desc) < 70 or len(desc) > 165: score -= 5

            if len(h1_tags) == 0: score -= 15
            elif len(h1_tags) > 1: score -= 5

            if not has_viewport: score -= 15

            total_imgs = len(all_imgs)
            if total_imgs > 0:
                alt_ratio = (total_imgs - missing_alt_count) / total_imgs
                dim_ratio = (total_imgs - missing_dimensions_count) / total_imgs
                score -= int((1 - alt_ratio) * 20)
                score -= int((1 - dim_ratio) * 10)

            score = max(5, min(100, score))

            return jsonify({
                "url": url,
                "score": score,
                "title": title,
                "title_length": len(title),
                "description": desc,
                "description_length": len(desc),
                "h1_tags": h1_tags,
                "h2_tags": h2_tags[:5],
                "canonical": canonical,
                "has_viewport": has_viewport,
                "image_stats": {
                    "total": total_imgs,
                    "missing_alt": missing_alt_count,
                    "missing_dimensions": missing_dimensions_count,
                    "improper_format": improper_format_count,
                    "lazy_loaded": lazy_loaded_count
                },
                "link_stats": {
                    "total": len(links_data),
                    "missing_title": missing_anchor_titles
                },
                "images": images_data[:100]
            })

        except Exception as e:
            return jsonify({"error": f"Failed to audit: {str(e)}"}), 500

    return jsonify({"error": "Invalid action"}), 400


# -------------------------------------------------------------
# 4. TOOL: BACKLINK OPPORTUNITY FINDER
# -------------------------------------------------------------
CATEGORIES = {
    "guest_blogs": [
        '"{kw}" "write for us"',
        '"{kw}" "guest post by"',
        '"{kw}" "submit a guest post"',
        '"{kw}" inurl:write-for-us',
        '"{kw}" "contribute to our blog"'
    ],
    "business_listings": [
        '"{kw}" "add your business"',
        '"{kw}" "submit listing"',
        '"{kw}" "add url" directory',
        '"{kw}" "business directory"'
    ],
    "classifieds": [
        '"{kw}" "post free ad"',
        '"{kw}" classifieds',
        '"{kw}" "submit classified ad"'
    ],
    "pdf_sharing": [
        '"{kw}" "upload presentation"',
        '"{kw}" "share pdf"',
        '"{kw}" filetype:pdf'
    ],
    "communities": [
        '"{kw}" forum',
        '"{kw}" "powered by vbulletin"',
        '"{kw}" "community discussion"'
    ]
}

@app.route("/api/backlink-finder", methods=["POST"])
def api_backlink_finder():
    data = request.get_json(silent=True) or {}
    keyword = data.get("keyword", "").strip()
    category = data.get("category", "guest_blogs")
    country = data.get("country", "")

    if not keyword:
        return jsonify({"error": "Keyword is required"}), 400

    templates = CATEGORIES.get(category, CATEGORIES["guest_blogs"])
    queries = [tpl.replace("{kw}", keyword) for tpl in templates]

    # Add country keyword if present
    country_suffix = f" {country}" if country else ""
    generated_search_urls = []

    for q in queries:
        full_query = q + country_suffix
        google_url = f"https://www.google.com/search?q={urllib.parse.quote(full_query)}"
        bing_url = f"https://www.bing.com/search?q={urllib.parse.quote(full_query)}"
        duck_url = f"https://duckduckgo.com/html/?q={urllib.parse.quote(full_query)}"
        generated_search_urls.append({
            "footprint": full_query,
            "google_url": google_url,
            "bing_url": bing_url,
            "duck_url": duck_url
        })

    # Perform DuckDuckGo live scrape on first footprint to get real link candidates
    live_opportunities = []
    try:
        sample_q = queries[0] + country_suffix
        duck_search = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(sample_q)}"
        resp = requests.post(
            "https://html.duckduckgo.com/html/",
            data={"q": sample_q},
            headers=HEADERS,
            timeout=8
        )
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            for result in soup.find_all("div", class_="result__body")[:15]:
                title_elem = result.find("a", class_="result__snippet") or result.find("a", class_="result__url")
                url_elem = result.find("a", class_="result__url")
                raw_href = url_elem.get("href", "") if url_elem else ""

                # DuckDuckGo redirect unwrapping
                actual_url = raw_href
                if "uddg=" in raw_href:
                    match = re.search(r"uddg=([^&]+)", raw_href)
                    if match:
                        actual_url = urllib.parse.unquote(match.group(1))

                snippet_elem = result.find("a", class_="result__snippet")
                snippet = snippet_elem.get_text(strip=True) if snippet_elem else ""

                if actual_url.startswith("http"):
                    live_opportunities.append({
                        "url": actual_url,
                        "title": snippet[:80] or actual_url,
                        "snippet": snippet,
                        "category": category.replace("_", " ").title()
                    })
    except Exception:
        pass

    return jsonify({
        "keyword": keyword,
        "category": category,
        "country": country,
        "footprints": generated_search_urls,
        "live_opportunities": live_opportunities
    })


# -------------------------------------------------------------
# 5. TOOL: SCHEMA.ORG & JSON-LD GENERATOR
# -------------------------------------------------------------
@app.route("/api/schema-gen", methods=["POST"])
def api_schema_gen():
    data = request.get_json(silent=True) or {}
    schema_type = data.get("schema_type", "Article")
    headline = data.get("headline", "").strip() or "Untitled Document"
    description = data.get("description", "").strip()
    author = data.get("author", "").strip() or "Editorial Team"
    page_url = data.get("url", "").strip() or "https://example.com/article"
    image_url = data.get("image", "").strip() or "https://example.com/image.jpg"
    publisher = data.get("publisher", "").strip() or "Organization"
    date_published = data.get("date_published", "").strip() or datetime.utcnow().strftime("%Y-%m-%d")

    if schema_type == "Article":
        schema_obj = {
            "@context": "https://schema.org",
            "@type": "Article",
            "mainEntityOfPage": {
                "@type": "WebPage",
                "@id": page_url
            },
            "headline": headline,
            "description": description,
            "image": image_url,
            "author": {
                "@type": "Person",
                "name": author
            },
            "publisher": {
                "@type": "Organization",
                "name": publisher,
                "logo": {
                    "@type": "ImageObject",
                    "url": image_url
                }
            },
            "datePublished": date_published,
            "dateModified": date_published
        }
    elif schema_type == "FAQ":
        questions = data.get("faqs", [
            {"q": "What is this service?", "a": "We provide top quality web and SEO solutions."}
        ])
        schema_obj = {
            "@context": "https://schema.org",
            "@type": "FAQPage",
            "mainEntity": [
                {
                    "@type": "Question",
                    "name": item.get("q", ""),
                    "acceptedAnswer": {
                        "@type": "Answer",
                        "text": item.get("a", "")
                    }
                } for item in questions if item.get("q")
            ]
        }
    else:
        schema_obj = {
            "@context": "https://schema.org",
            "@type": "WebPage",
            "name": headline,
            "description": description,
            "url": page_url
        }

    json_ld_string = json.dumps(schema_obj, indent=2)
    html_snippet = f'<script type="application/ld+json">\n{json_ld_string}\n</script>'

    return jsonify({
        "schema_json": schema_obj,
        "json_ld_code": html_snippet
    })


# -------------------------------------------------------------
# 6. TOOL: YOUTUBE SHORTS TREND HUNTER
# -------------------------------------------------------------
@app.route("/api/shorts-hunter", methods=["POST"])
def api_shorts_hunter():
    data = request.get_json(silent=True) or {}
    query = data.get("query", "").strip()
    max_results = min(int(data.get("max_results", 10)), 20)

    if not query:
        return jsonify({"error": "Search query is required"}), 400

    results = []
    try:
        search_url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(query + ' shorts')}"
        resp = requests.get(search_url, headers=HEADERS, timeout=8)

        # YouTube embeds initialData in page HTML
        match = re.search(r"var ytInitialData = ({.*?});</script>", resp.text)
        if match:
            raw_data = json.loads(match.group(1))
            sections = (
                raw_data.get("contents", {})
                .get("twoColumnSearchResultsRenderer", {})
                .get("primaryContents", {})
                .get("sectionListRenderer", {})
                .get("contents", [])
            )
            for section in sections:
                item_section = section.get("itemSectionRenderer", {}).get("contents", [])
                for item in item_section:
                    # Video renderer
                    vr = item.get("videoRenderer")
                    if vr:
                        vid_id = vr.get("videoId")
                        title = vr.get("title", {}).get("runs", [{}])[0].get("text", "N/A")
                        channel = vr.get("ownerText", {}).get("runs", [{}])[0].get("text", "N/A")
                        views = vr.get("viewCountText", {}).get("simpleText", "N/A")
                        duration = vr.get("lengthText", {}).get("simpleText", "N/A")

                        if vid_id:
                            results.append({
                                "video_id": vid_id,
                                "title": title,
                                "channel": channel,
                                "views": views,
                                "duration": duration,
                                "url": f"https://www.youtube.com/shorts/{vid_id}"
                            })
                            if len(results) >= max_results:
                                break
                if len(results) >= max_results:
                    break
    except Exception as e:
        return jsonify({"error": f"Failed to search YouTube: {str(e)}"}), 500

    return jsonify({
        "query": query,
        "count": len(results),
        "results": results
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
