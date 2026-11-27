import io
import re
import csv
import json
import time
import urllib.parse
from datetime import datetime
from urllib.parse import urljoin, urlparse
from xml.sax.saxutils import escape as xml_escape

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import xml.etree.ElementTree as ET

try:
    import yt_dlp
except ImportError:
    yt_dlp = None

try:
    import pypdf
except ImportError:
    pypdf = None

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
# 2. TOOL: XML IMAGE SITEMAP GENERATOR & CRAWLER (ENHANCED)
# -------------------------------------------------------------
CSS_URL_RE = re.compile(r"""url\(\s*['"]?([^'")]+)['"]?\s*\)""", re.IGNORECASE)

def title_from_filename(url: str) -> str:
    path = urlparse(url).path
    filename = path.rsplit("/", 1)[-1] if "/" in path else path
    stem = filename.rsplit(".", 1)[0]
    return stem.replace("-", " ").replace("_", " ").title()

@app.route("/api/sitemap-crawler", methods=["POST"])
def api_sitemap_crawler():
    data = request.get_json(silent=True) or {}
    action = data.get("action", "crawl")
    base_url = data.get("url", "").strip()
    filter_webp_only = data.get("webp_only", False)
    exclude_common = data.get("exclude_common", True)

    if not base_url:
        return jsonify({"error": "URL is required"}), 400
    if not base_url.startswith("http://") and not base_url.startswith("https://"):
        base_url = "https://" + base_url

    parsed_base = urlparse(base_url)
    domain = parsed_base.netloc
    max_pages = min(int(data.get("max_pages", 10)), 25)

    if action == "crawl":
        crawl_mode = data.get("crawl_mode", "crawler")
        visited = set()
        to_visit = [base_url]
        page_images_map = {}  # url -> list of {image_url, alt, title}

        if crawl_mode == "sitemap_xml":
            def _get_sitemap_pages(sm_url, depth=0):
                if depth > 2:
                    return []
                found = []
                try:
                    r = requests.get(sm_url, headers=HEADERS, timeout=7)
                    if r.status_code == 200:
                        root = ET.fromstring(r.content)
                        ns = {"ns": "http://www.sitemaps.org/schemas/sitemap/0.9"}
                        sub_sitemaps = root.findall(".//ns:sitemap/ns:loc", ns)
                        if sub_sitemaps:
                            for sm in sub_sitemaps[:5]:
                                if sm.text:
                                    found.extend(_get_sitemap_pages(sm.text.strip(), depth + 1))
                        else:
                            for u in root.findall(".//ns:url/ns:loc", ns):
                                if u.text:
                                    found.append(u.text.strip())
                except Exception:
                    pass
                return found

            sm_pages = _get_sitemap_pages(base_url)
            if sm_pages:
                to_visit = sm_pages[:max_pages]

        while to_visit and len(visited) < max_pages:
            current_url = to_visit.pop(0)
            if current_url in visited:
                continue
            visited.add(current_url)

            try:
                resp = requests.get(current_url, headers=HEADERS, timeout=7)
                if resp.status_code != 200:
                    continue
                soup = BeautifulSoup(resp.text, "html.parser")
                page_imgs = []

                # 1. <img> tags
                for img in soup.find_all("img"):
                    candidates = []
                    if img.get("src"): candidates.append(img["src"])
                    if img.get("data-src"): candidates.append(img["data-src"])
                    if img.get("srcset"):
                        for part in img["srcset"].split(","):
                            p_url = part.strip().split(" ")[0]
                            if p_url: candidates.append(p_url)

                    alt = img.get("alt", "").strip()
                    title = img.get("title", "").strip()

                    for src in candidates:
                        if not src or src.startswith("data:"):
                            continue
                        abs_img = urljoin(current_url, src)
                        is_webp = abs_img.lower().endswith(".webp")
                        if filter_webp_only and not is_webp:
                            continue
                        
                        display_title = title if title else (alt if alt else title_from_filename(abs_img))
                        page_imgs.append({
                            "page_url": current_url,
                            "image_url": abs_img,
                            "alt_text": alt or "N/A",
                            "title": display_title,
                            "is_webp": is_webp
                        })

                # 2. CSS background images (inline styles & <style> tags)
                for tag in soup.find_all(style=True):
                    for match in CSS_URL_RE.finditer(tag["style"]):
                        raw_src = match.group(1).strip()
                        if raw_src and not raw_src.startswith("data:"):
                            abs_img = urljoin(current_url, raw_src)
                            is_webp = abs_img.lower().endswith(".webp")
                            if filter_webp_only and not is_webp:
                                continue
                            page_imgs.append({
                                "page_url": current_url,
                                "image_url": abs_img,
                                "alt_text": "CSS Background",
                                "title": title_from_filename(abs_img),
                                "is_webp": is_webp
                            })

                # Deduplicate per page
                seen = set()
                unique_page_imgs = []
                for item in page_imgs:
                    if item["image_url"] not in seen:
                        seen.add(item["image_url"])
                        unique_page_imgs.append(item)

                page_images_map[current_url] = unique_page_imgs

                # Discover internal links for queue
                if len(visited) < max_pages:
                    for a in soup.find_all("a", href=True):
                        link = urljoin(current_url, a["href"]).split("#")[0]
                        p_link = urlparse(link)
                        if p_link.netloc == domain and link not in visited and link not in to_visit:
                            if not p_link.path.lower().endswith((".png", ".jpg", ".jpeg", ".pdf", ".zip", ".rar", ".mp4")):
                                to_visit.append(link)

            except Exception:
                continue

        # Exclude common images (e.g. site-wide logos, icons appearing on >50% of crawled pages)
        if exclude_common and len(page_images_map) > 2:
            counts = {}
            for imgs in page_images_map.values():
                for im in imgs:
                    counts[im["image_url"]] = counts.get(im["image_url"], 0) + 1
            threshold = max(2, int(len(page_images_map) * 0.5))
            common_urls = {u for u, c in counts.items() if c >= threshold}

            filtered_map = {}
            for pu, imgs in page_images_map.items():
                filtered_map[pu] = [im for im in imgs if im["image_url"] not in common_urls]
            page_images_map = filtered_map

        # Flatten list for client
        all_discovered = []
        for imgs in page_images_map.values():
            all_discovered.extend(imgs)

        # Build Standard Google Image Sitemap XML
        xml_lines = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"',
            '        xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">'
        ]
        for pu, imgs in page_images_map.items():
            if not imgs:
                continue
            xml_lines.append("  <url>")
            xml_lines.append(f"    <loc>{xml_escape(pu)}</loc>")
            for im in imgs[:100]:
                xml_lines.append("    <image:image>")
                xml_lines.append(f"      <image:loc>{xml_escape(im['image_url'])}</image:loc>")
                if im["title"] and im["title"] != "N/A":
                    xml_lines.append(f"      <image:title>{xml_escape(im['title'])}</image:title>")
                xml_lines.append("    </image:image>")
            xml_lines.append("  </url>")
        xml_lines.append("</urlset>")
        sitemap_xml = "\n".join(xml_lines)

        return jsonify({
            "pages_crawled": len(visited),
            "total_images": len(all_discovered),
            "images": all_discovered[:300],
            "xml_sitemap": sitemap_xml
        })

    return jsonify({"error": "Invalid action"}), 400


# -------------------------------------------------------------
# 3. TOOL: DEEP SEO & IMAGE AUDITOR (WITH EXCEL DOWNLOAD)
# -------------------------------------------------------------
@app.route("/api/seo-audit", methods=["POST"])
def api_seo_audit():
    data = request.get_json(silent=True) or {}
    action = data.get("action", "audit")
    url = data.get("url", "").strip()
    html_content = data.get("html_content", "").strip()

    if not url and not html_content:
        return jsonify({"error": "URL or HTML content is required"}), 400
    if url and not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url

    try:
        if html_content:
            raw_html = html_content
            if not url:
                url = "https://raw-html-input.local"
        else:
            resp = requests.get(url, headers=HEADERS, timeout=10)
            raw_html = resp.text
        soup = BeautifulSoup(raw_html, "html.parser")

        # Basic Head & Meta
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

        # Deep Image Audits
        images_data = []
        missing_alt_count = 0
        missing_title_count = 0
        missing_dimensions_count = 0
        improper_format_count = 0
        lazy_loaded_count = 0

        PERMISSIBLE_FORMATS = {"webp", "svg", "gif"}

        all_imgs = soup.find_all("img")
        for idx, img in enumerate(all_imgs, 1):
            src = img.get("src") or img.get("data-src") or ""
            full_src = urljoin(url, src) if src else ""
            alt = img.get("alt")
            title_attr = img.get("title")
            width = img.get("width")
            height = img.get("height")
            loading = img.get("loading", "eager").lower()

            clean_path = urlparse(full_src).path.lower()
            ext = clean_path.split(".")[-1] if "." in clean_path else "unknown"

            is_alt_missing = (alt is None or alt.strip() == "")
            is_title_missing = (title_attr is None or title_attr.strip() == "")
            is_dim_missing = (not width or not height)
            is_format_warning = (ext not in PERMISSIBLE_FORMATS and ext in {"jpg", "jpeg", "png"})

            if is_alt_missing: missing_alt_count += 1
            if is_title_missing: missing_title_count += 1
            if is_dim_missing: missing_dimensions_count += 1
            if is_format_warning: improper_format_count += 1
            if loading == "lazy": lazy_loaded_count += 1

            status = "✅ OK"
            if is_alt_missing or is_dim_missing or is_format_warning:
                status = "⚠️ Needs Optimization"

            images_data.append({
                "index": idx,
                "src": full_src,
                "alt": alt.strip() if alt else "MISSING",
                "has_alt": not is_alt_missing,
                "title": title_attr.strip() if title_attr else "MISSING",
                "has_title": not is_title_missing,
                "dimensions": f"{width}x{height}" if not is_dim_missing else "MISSING",
                "has_dimensions": not is_dim_missing,
                "format": ext.upper(),
                "is_modern_format": ext in PERMISSIBLE_FORMATS,
                "loading": loading,
                "status": status
            })

        # Deep Anchor Link Audits (from New folder/SEO)
        links_data = []
        missing_anchor_titles = 0
        title_matches_text_count = 0

        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            anchor_title = a.get("title", "").strip()
            anchor_text = a.get_text(strip=True)

            is_title_missing = not bool(anchor_title)
            is_title_matches = (bool(anchor_title) and anchor_title.lower() == anchor_text.lower())

            if is_title_missing: missing_anchor_titles += 1
            if is_title_matches: title_matches_text_count += 1

            status = "✅ OK"
            issue_note = "Valid Title"
            if is_title_missing:
                status = "⚠️ Missing Title"
                issue_note = "Missing anchor title tag"
            elif is_title_matches:
                status = "⚠️ Duplicate"
                issue_note = "Title tag identical to anchor text (bad SEO)"

            links_data.append({
                "href": urljoin(url, href),
                "text": anchor_text[:60] or "[No Text / Icon]",
                "title": anchor_title or "MISSING",
                "status": status,
                "issue_note": issue_note
            })

        # Calculate Overall Health Score
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

        # IF ACTION == EXPORT EXCEL (The requested feature from images checker)
        if action == "export_excel":
            wb = openpyxl.Workbook()
            header_fill = PatternFill(start_color="1F4788", end_color="1F4788", fill_type="solid")
            header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
            ok_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            issue_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            border = Border(left=Side(style='thin', color="CCCCCC"),
                            right=Side(style='thin', color="CCCCCC"),
                            top=Side(style='thin', color="CCCCCC"),
                            bottom=Side(style='thin', color="CCCCCC"))
            center_align = Alignment(horizontal='center', vertical='center')

            # --- Sheet 1: Summary Report ---
            ws_sum = wb.active
            ws_sum.title = "Summary Report"
            sum_headers = ['Audited URL', 'Total Images', 'Compliant Images', 'Issues Found', 'Compliance %', 'Total Links', 'Links Missing Titles', 'Timestamp']
            for col, text in enumerate(sum_headers, 1):
                c = ws_sum.cell(row=1, column=col, value=text)
                c.fill = header_fill
                c.font = header_font
                c.alignment = center_align
                c.border = border

            comp_imgs = total_imgs - missing_alt_count - missing_dimensions_count
            comp_pct = f"{(max(0, comp_imgs) / total_imgs * 100):.1f}%" if total_imgs > 0 else "100%"
            row_vals = [url, total_imgs, max(0, comp_imgs), missing_alt_count + missing_dimensions_count, comp_pct, len(links_data), missing_anchor_titles, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
            for col, val in enumerate(row_vals, 1):
                c = ws_sum.cell(row=2, column=col, value=val)
                c.border = border
                c.alignment = center_align

            # --- Sheet 2: Detailed Image Audit ---
            ws_img = wb.create_sheet("Image Audit")
            img_headers = ['Page URL', 'Image URL', 'Alt Text', 'Title Tag', 'Loading Attribute', 'Dimensions', 'Format', 'Status']
            for col, text in enumerate(img_headers, 1):
                c = ws_img.cell(row=1, column=col, value=text)
                c.fill = header_fill
                c.font = header_font
                c.alignment = center_align
                c.border = border

            for r_idx, im in enumerate(images_data, 2):
                r_data = [url, im['src'], im['alt'], im['title'], im['loading'], im['dimensions'], im['format'], im['status']]
                for c_idx, v in enumerate(r_data, 1):
                    cell = ws_img.cell(row=r_idx, column=c_idx, value=v)
                    cell.border = border
                    if v == "MISSING":
                        cell.fill = issue_fill
                    elif c_idx == 8:
                        cell.fill = ok_fill if "OK" in v else issue_fill

            # --- Sheet 3: Anchor Link Audit (from SEO folder) ---
            ws_links = wb.create_sheet("Anchor Links Audit")
            link_headers = ['Page URL', 'Destination URL', 'Anchor Text', 'Title Attribute', 'Status', 'SEO Issue Details']
            for col, text in enumerate(link_headers, 1):
                c = ws_links.cell(row=1, column=col, value=text)
                c.fill = header_fill
                c.font = header_font
                c.alignment = center_align
                c.border = border

            for r_idx, lnk in enumerate(links_data, 2):
                r_data = [url, lnk['href'], lnk['text'], lnk['title'], lnk['status'], lnk['issue_note']]
                for c_idx, v in enumerate(r_data, 1):
                    cell = ws_links.cell(row=r_idx, column=c_idx, value=v)
                    cell.border = border
                    if "Missing" in str(v) or "Duplicate" in str(v):
                        cell.fill = issue_fill
                    elif c_idx == 5 and "OK" in str(v):
                        cell.fill = ok_fill

            # Auto width
            for ws in [ws_sum, ws_img, ws_links]:
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return send_file(
                buf,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="detailed_seo_audit_report.xlsx"
            )

        # Standard JSON audit response
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
                "missing_title": missing_title_count,
                "missing_dimensions": missing_dimensions_count,
                "improper_format": improper_format_count,
                "lazy_loaded": lazy_loaded_count
            },
            "link_stats": {
                "total": len(links_data),
                "missing_title": missing_anchor_titles,
                "duplicate_title": title_matches_text_count
            },
            "images": images_data[:150],
            "links": links_data[:100]
        })

    except Exception as e:
        return jsonify({"error": f"Failed to audit: {str(e)}"}), 500


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

    live_opportunities = []
    try:
        sample_q = queries[0] + country_suffix
        resp = requests.post(
            "https://html.duckduckgo.com/html/",
            data={"q": sample_q},
            headers=HEADERS,
            timeout=8
        )
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            for result in soup.find_all("div", class_="result__body")[:15]:
                url_elem = result.find("a", class_="result__url")
                raw_href = url_elem.get("href", "") if url_elem else ""

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
    author = data.get("author", "").strip() or "Tesla Mechanical Designs"
    page_url = data.get("url", "").strip() or "https://www.teslamechanicaldesigns.com"
    image_url = data.get("image", "").strip() or "https://www.teslamechanicaldesigns.com/images/hero.webp"
    publisher = data.get("publisher", "").strip() or "Tesla Mechanical Designs"
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


@app.route("/api/pdf-extract-schema", methods=["POST"])
def api_pdf_extract_schema():
    if pypdf is None:
        return jsonify({"error": "pypdf library is not installed on the server"}), 500
    if "file" not in request.files:
        return jsonify({"error": "No PDF file uploaded"}), 400

    pdf_file = request.files["file"]
    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Please upload a valid .pdf file"}), 400

    try:
        reader = pypdf.PdfReader(pdf_file.stream)
        info = reader.metadata or {}

        stem = pdf_file.filename.rsplit(".", 1)[0]
        headline = str(info.get("/Title") or "").strip() or stem
        author = str(info.get("/Author") or "").strip() or "Unknown Author"

        raw_date = str(info.get("/CreationDate") or "")
        date_published = datetime.utcnow().strftime("%Y-%m-%d")
        if raw_date:
            try:
                clean_d = raw_date.replace("D:", "").split("+")[0].replace("'", "")
                dt = datetime.strptime(clean_d[:8], "%Y%m%d")
                date_published = dt.strftime("%Y-%m-%d")
            except Exception:
                pass

        description = ""
        if len(reader.pages) > 0:
            first_text = reader.pages[0].extract_text() or ""
            clean_text = " ".join(first_text.split())
            if clean_text:
                description = clean_text[:300] + ("..." if len(clean_text) > 300 else "")

        return jsonify({
            "headline": headline,
            "author": author,
            "date_published": date_published,
            "description": description or "No text description extracted from PDF.",
            "filename": pdf_file.filename
        })
    except Exception as e:
        return jsonify({"error": f"Failed to parse PDF: {str(e)}"}), 500


# -------------------------------------------------------------
# 6. TOOL: YOUTUBE SHORTS TREND HUNTER (ENHANCED YT-DLP)
# -------------------------------------------------------------
@app.route("/api/shorts-hunter", methods=["POST"])
def api_shorts_hunter():
    data = request.get_json(silent=True) or {}
    query = data.get("query", "").strip()
    max_results = min(int(data.get("max_results", 10)), 20)

    if not query:
        return jsonify({"error": "Search query is required"}), 400

    results = []

    # 1. Try yt-dlp first (matching youtube-shorts-hunter/main.py)
    if yt_dlp is not None:
        try:
            ydl_opts = {
                'quiet': True,
                'no_warnings': True,
                'extract_flat': 'in_playlist',
                'default_search': 'ytsearch',
            }
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                search_query = f"ytsearch{max_results}:{query} shorts"
                search_results = ydl.extract_info(search_query, download=False)
                if search_results and 'entries' in search_results:
                    for entry in search_results['entries']:
                        if not entry: continue
                        vid_id = entry.get('id', '')
                        title = entry.get('title', 'N/A')
                        channel = entry.get('channel') or entry.get('uploader') or 'N/A'
                        views_raw = entry.get('view_count')
                        views = f"{views_raw:,}" if isinstance(views_raw, int) else (views_raw or "N/A")
                        duration_raw = entry.get('duration')
                        duration = f"{duration_raw}s" if duration_raw else "Short"

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
        except Exception:
            pass

    # 2. Resilient fallback scraper if yt-dlp was throttled
    if not results:
        try:
            search_url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(query + ' shorts')}"
            resp = requests.get(search_url, headers=HEADERS, timeout=8)
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
            return jsonify({"error": f"Failed to search: {str(e)}"}), 500

    return jsonify({
        "query": query,
        "count": len(results),
        "results": results
    })


# -------------------------------------------------------------
# 7. TOOL: UNIVERSAL VIDEO DOWNLOADER (FROM UNIVERSAL_DOWNLOADER.PY)
# -------------------------------------------------------------
@app.route("/api/video-downloader", methods=["POST"])
def api_video_downloader():
    data = request.get_json(silent=True) or {}
    url = data.get("url", "").strip()
    if not url:
        return jsonify({"error": "Video URL is required"}), 400

    if yt_dlp is None:
        return jsonify({"error": "yt-dlp engine is not available"}), 500

    try:
        ydl_opts = {
            'quiet': True,
            'no_warnings': True,
            'skip_download': True,
        }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            title = info.get("title", "Untitled Video")
            thumbnail = info.get("thumbnail", "")
            duration = info.get("duration", 0)
            uploader = info.get("uploader", info.get("channel", "Unknown Channel"))

            mins = int(duration // 60) if duration else 0
            secs = int(duration % 60) if duration else 0
            duration_str = f"{mins}:{secs:02d}" if duration else "N/A"

            formats = []
            direct_video_url = None
            if "formats" in info:
                for f in info["formats"]:
                    ext = f.get("ext", "mp4")
                    height = f.get("height")
                    f_url = f.get("url")
                    if f_url and ext in ["mp4", "webm"]:
                        if height and height >= 360:
                            formats.append({
                                "resolution": f"{height}p",
                                "ext": ext,
                                "url": f_url
                            })
                            if not direct_video_url or height == 720:
                                direct_video_url = f_url

            if not direct_video_url:
                direct_video_url = info.get("url", "")

            return jsonify({
                "title": title,
                "thumbnail": thumbnail,
                "duration": duration_str,
                "uploader": uploader,
                "source_url": url,
                "download_url": direct_video_url,
                "formats": formats[:6]
            })
    except Exception as e:
        return jsonify({"error": f"Failed to extract video: {str(e)}"}), 500


# -------------------------------------------------------------
# 8. TOOL: FULL ON-PAGE & TECHNICAL CHECKLIST AUDITOR
# -------------------------------------------------------------
@app.route("/api/full-checklist-audit", methods=["POST"])
def api_full_checklist_audit():
    data = request.get_json(silent=True) or {}
    action = data.get("action", "audit")
    url = data.get("url", "").strip()

    if not url:
        return jsonify({"error": "URL is required"}), 400
    if not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url

    parsed = urlparse(url)
    root_domain = parsed.netloc.replace("www.", "")
    base_https_www = f"https://www.{root_domain}"
    base_origin = f"{parsed.scheme}://{parsed.netloc}"

    checks = []
    def add_check(cat, particular, status, value, remark):
        checks.append({
            "category": cat,
            "particular": particular,
            "status": status,  # PASS, WARN, FAIL
            "value": str(value)[:250],
            "remark": remark
        })

    try:
        resp = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
        raw_html = resp.text
        soup = BeautifulSoup(raw_html, "html.parser")
    except Exception as e:
        return jsonify({"error": f"Could not fetch URL: {str(e)}"}), 500

    # --- 1. SITE-LEVEL CHECKS ---
    # Domain Redirect Check
    variants = [
        f"http://{root_domain}",
        f"https://{root_domain}",
        f"http://www.{root_domain}",
        f"https://www.{root_domain}"
    ]
    redirect_ok = True
    redirect_notes = []
    for v in variants:
        try:
            r_v = requests.head(v, headers=HEADERS, timeout=4, allow_redirects=True)
            final_u = r_v.url.rstrip("/")
            if not final_u.startswith("https://www."):
                redirect_ok = False
            redirect_notes.append(f"{v} -> {final_u}")
        except Exception:
            redirect_ok = False
            redirect_notes.append(f"{v} -> Error")
    add_check(
        "Site-Level", "Domain Redirect Check",
        "PASS" if redirect_ok else "WARN",
        " | ".join(redirect_notes[:2]) + "...",
        "All variants should 301 redirect to https://www." if not redirect_ok else "All variants resolve cleanly to https://www."
    )

    # robots.txt
    try:
        r_rob = requests.get(f"{base_origin}/robots.txt", headers=HEADERS, timeout=5)
        if r_rob.status_code == 200:
            has_sm = "sitemap:" in r_rob.text.lower()
            add_check("Site-Level", "robots.txt", "PASS" if has_sm else "WARN",
                      "Found (200 OK)", "Sitemap declared in robots.txt" if has_sm else "robots.txt found, but Sitemap: line is missing")
        else:
            add_check("Site-Level", "robots.txt", "FAIL", f"HTTP {r_rob.status_code}", "robots.txt not found")
    except Exception:
        add_check("Site-Level", "robots.txt", "FAIL", "Error", "Could not fetch robots.txt")

    # sitemap.xml
    try:
        r_sm = requests.get(f"{base_origin}/sitemap.xml", headers=HEADERS, timeout=5)
        if r_sm.status_code == 200 and "<urlset" in r_sm.text or "<sitemapindex" in r_sm.text:
            lastmods = re.findall(r"<lastmod>(.*?)</lastmod>", r_sm.text)
            latest = lastmods[0] if lastmods else "No <lastmod> tag"
            add_check("Site-Level", "sitemap.xml", "PASS" if lastmods else "WARN",
                      f"Found | Latest: {latest}", "Verify timestamps (YYYY-MM-DDTHH:MM:SSZ) & update every 45 days")
        else:
            add_check("Site-Level", "sitemap.xml", "FAIL", f"HTTP {r_sm.status_code}", "sitemap.xml missing or invalid")
    except Exception:
        add_check("Site-Level", "sitemap.xml", "FAIL", "Error", "Could not fetch sitemap.xml")

    # 404 Implementation
    try:
        r_404 = requests.get(f"{base_origin}/nonexistent-404-check-xyz", headers=HEADERS, timeout=5)
        add_check("Site-Level", "404 Implementation",
                  "PASS" if r_404.status_code == 404 else "FAIL",
                  f"HTTP {r_404.status_code}",
                  "Proper 404 status code returned" if r_404.status_code == 404 else "Soft-404 detected! Must return HTTP 404")
    except Exception:
        add_check("Site-Level", "404 Implementation", "WARN", "Error", "Could not test 404 page")

    # --- 2. CRITICAL HTML & <HEAD> TAGS ---
    # Doctype
    has_doctype = raw_html.lstrip().lower().startswith("<!doctype html>")
    add_check("Critical <head>", "HTML Doctype", "PASS" if has_doctype else "FAIL",
              "<!doctype html>" if has_doctype else "Missing", "<!doctype html> must be defined")

    # HTML Lang
    html_tag = soup.find("html")
    lang_val = html_tag.get("lang", "").strip() if html_tag else ""
    valid_langs = {"en", "en-us", "en-au", "en-in"}
    add_check("Critical <head>", "HTML Lang Tag",
              "PASS" if lang_val.lower() in valid_langs else ("WARN" if lang_val else "FAIL"),
              lang_val or "Missing", "Permissible: en, en-US, en-AU, en-IN")

    # Charset
    charset_tag = soup.find("meta", charset=True) or soup.find("meta", attrs={"http-equiv": re.compile("content-type", re.I)})
    charset_val = charset_tag.get("charset", "") if charset_tag and charset_tag.get("charset") else ("utf-8" if charset_tag and "utf-8" in str(charset_tag).lower() else "")
    add_check("Critical <head>", "Charset (UTF-8)", "PASS" if charset_val.lower() == "utf-8" else "FAIL",
              charset_val or "Missing", "Only UTF-8 is permissible")

    # Preconnect & DNS-Prefetch
    preconnects = soup.find_all("link", rel=lambda r: r and ("preconnect" in r or "dns-prefetch" in r))
    add_check("Critical <head>", "DNS-Prefetch & Preconnect", "PASS" if preconnects else "WARN",
              f"{len(preconnects)} found", "Recommended for faster external font/CDN loading")

    # Preload
    preloads = soup.find_all("link", rel=lambda r: r and "preload" in r)
    add_check("Critical <head>", "Viewport Image Preload", "PASS" if preloads else "WARN",
              f"{len(preloads)} found", "Preload largest above-the-fold image for better LCP")

    # Viewport
    vp = soup.find("meta", attrs={"name": re.compile("^viewport$", re.I)})
    vp_content = vp.get("content", "") if vp else ""
    add_check("Critical <head>", "Viewport Tag", "PASS" if "width=device-width" in vp_content else "FAIL",
              vp_content or "Missing", "Required for mobile responsiveness")

    # Canonical & Hreflang
    canon = soup.find("link", attrs={"rel": re.compile("^canonical$", re.I)})
    canon_href = canon.get("href", "").strip() if canon else ""
    hreflangs = [h.get("hreflang", "").lower() for h in soup.find_all("link", attrs={"rel": "alternate", "hreflang": True})]
    add_check("Critical <head>", "Canonical Tag", "PASS" if canon_href else "FAIL",
              canon_href or "Missing", "Canonical URL must be defined")
    add_check("Critical <head>", "Hreflang Tags",
              "PASS" if "x-default" in hreflangs else ("WARN" if hreflangs else "WARN"),
              ", ".join(hreflangs) if hreflangs else "None", "Set x-default, en-us, en-au, en-ca, en-gb if multi-regional")

    # Meta Keywords
    kw_tag = soup.find("meta", attrs={"name": re.compile("^keywords$", re.I)})
    kw_val = kw_tag.get("content", "").strip() if kw_tag else ""
    add_check("Critical <head>", "Meta Keywords (Bing)", "PASS" if kw_val else "WARN",
              kw_val[:120] if kw_val else "Missing", "Important for Bing SEO")

    # Title & Description
    title_tag = soup.find("title")
    title_val = title_tag.get_text(strip=True) if title_tag else ""
    t_len = len(title_val)
    add_check("Critical <head>", "Meta Title",
              "PASS" if 50 <= t_len <= 75 else ("WARN" if t_len > 0 else "FAIL"),
              f"({t_len} chars) {title_val}", "Target: 60–75 characters")

    desc_tag = soup.find("meta", attrs={"name": re.compile("^description$", re.I)})
    desc_val = desc_tag.get("content", "").strip() if desc_tag else ""
    d_len = len(desc_val)
    add_check("Critical <head>", "Meta Description",
              "PASS" if 120 <= d_len <= 165 else ("WARN" if d_len > 0 else "FAIL"),
              f"({d_len} chars) {desc_val}", "Target: 140–160 characters")

    # Robots / Googlebot / Bingbot / Noindex
    rob_tag = soup.find("meta", attrs={"name": re.compile("^robots$", re.I)})
    rob_val = rob_tag.get("content", "").strip() if rob_tag else ""
    has_noindex = "noindex" in rob_val.lower()
    add_check("Critical <head>", "Robots & Noindex Check",
              "FAIL" if has_noindex else ("PASS" if "index" in rob_val.lower() else "WARN"),
              rob_val or "Not set (defaults to index, follow)",
              "WARNING: noindex detected!" if has_noindex else "Ensure index, follow on live pages")

    # --- 3. SOCIAL, SCHEMA & TRACKING ---
    og_required = ["og:title", "og:locale", "og:url", "og:type", "og:description", "og:image"]
    og_found = {}
    for prop in og_required:
        t = soup.find("meta", property=prop)
        if t and t.get("content"):
            og_found[prop] = t["content"].strip()
    og_missing = [p for p in og_required if p not in og_found]
    add_check("Social & Schema", "OpenGraph (OG) Tags",
              "PASS" if not og_missing else ("WARN" if og_found else "FAIL"),
              f"{len(og_found)}/{len(og_required)} present",
              f"Missing: {', '.join(og_missing)}" if og_missing else "All core OG tags present")

    tw_required = ["twitter:card", "twitter:title", "twitter:description", "twitter:image"]
    tw_found = {}
    for name in tw_required:
        t = soup.find("meta", attrs={"name": name})
        if t and t.get("content"):
            tw_found[name] = t["content"].strip()
    tw_missing = [n for n in tw_required if n not in tw_found]
    add_check("Social & Schema", "Twitter Card Tags",
              "PASS" if not tw_missing else ("WARN" if tw_found else "FAIL"),
              f"{len(tw_found)}/{len(tw_required)} present",
              f"Missing: {', '.join(tw_missing)}" if tw_missing else "All core Twitter tags present")

    # Schemas (JSON-LD)
    schema_scripts = soup.find_all("script", type="application/ld+json")
    schema_types = []
    for sc in schema_scripts:
        try:
            obj = json.loads(sc.string or "{}")
            if isinstance(obj, dict) and "@type" in obj:
                schema_types.append(str(obj["@type"]))
            elif isinstance(obj, list):
                for item in obj:
                    if isinstance(item, dict) and "@type" in item:
                        schema_types.append(str(item["@type"]))
        except Exception:
            pass
    add_check("Social & Schema", "JSON-LD Schemas", "PASS" if schema_types else "WARN",
              ", ".join(schema_types) if schema_types else "None detected",
              "Verify Organization, WebSite, Service, BreadcrumbList, Article schemas")

    # Google Tag Manager
    gtm_match = re.findall(r"GTM-[A-Z0-9]+", raw_html)
    gtm_ids = list(set(gtm_match))
    add_check("Social & Schema", "Google Tag Manager (GTM)", "PASS" if gtm_ids else "WARN",
              ", ".join(gtm_ids) if gtm_ids else "Not detected", "Verify GTM container ID")

    # --- 4. ICONS, HEADINGS, IMAGES & ANCHORS ---
    favicon = soup.find("link", rel=lambda r: r and "icon" in r)
    fav_href = favicon.get("href", "") if favicon else ""
    add_check("Icons & UI", "Favicon (.ico)", "PASS" if ".ico" in fav_href.lower() else ("WARN" if fav_href else "FAIL"),
              fav_href or "Missing", "Preferred format is .ico")

    apple_title = soup.find("meta", attrs={"name": "apple-mobile-web-app-title"})
    add_check("Icons & UI", "Apple WebView Title", "PASS" if apple_title and apple_title.get("content") else "WARN",
              apple_title.get("content", "") if apple_title else "Missing", "Recommended for iOS home screen")

    apple_icons = soup.find_all("link", rel=lambda r: r and "apple-touch-icon" in r)
    add_check("Icons & UI", "Apple Touch Icons", "PASS" if len(apple_icons) >= 4 else ("WARN" if apple_icons else "FAIL"),
              f"{len(apple_icons)} size(s) defined", "Recommended sizes: 57x57 up to 180x180")

    theme_color = soup.find("meta", attrs={"name": "theme-color"})
    tc_val = theme_color.get("content", "") if theme_color else ""
    add_check("Icons & UI", "Theme Color", "PASS" if tc_val else "WARN",
              tc_val or "Missing", "Should match primary brand hex color")

    # Headings
    h1s = [h.get_text(strip=True) for h in soup.find_all("h1")]
    h2s = [h.get_text(strip=True) for h in soup.find_all("h2")]
    links_in_headings = any(h.find("a") for h in soup.find_all(["h1", "h2", "h3", "h4"]))
    add_check("Content & Media", "H1 & Heading Hierarchy",
              "PASS" if len(h1s) == 1 and not links_in_headings else "WARN",
              f"{len(h1s)} H1(s), {len(h2s)} H2(s)",
              "Links found inside headings!" if links_in_headings else ("Must have exactly 1 H1" if len(h1s) != 1 else "Clean heading structure"))

    # Images summary
    imgs = soup.find_all("img")
    miss_alt = sum(1 for i in imgs if not (i.get("alt") or "").strip())
    miss_dim = sum(1 for i in imgs if not i.get("width") or not i.get("height"))
    bad_fmt = sum(1 for i in imgs if urlparse(i.get("src", "")).path.lower().endswith((".png", ".jpg", ".jpeg")))
    add_check("Content & Media", "Images (Alt, Size, WebP)",
              "PASS" if (miss_alt == 0 and miss_dim == 0 and bad_fmt == 0) else "WARN",
              f"Total: {len(imgs)} | Missing Alt: {miss_alt} | Missing WxH: {miss_dim} | PNG/JPG: {bad_fmt}",
              "Use WebP/SVG, explicit width/height, and descriptive alt text")

    # Anchors summary
    anchors = soup.find_all("a", href=True)
    miss_a_title = sum(1 for a in anchors if not (a.get("title") or "").strip())
    add_check("Content & Media", "Anchor Title Attributes",
              "PASS" if miss_a_title == 0 else "WARN",
              f"Total: {len(anchors)} | Missing Title: {miss_a_title}",
              "Every anchor should have a descriptive title != anchor text")

    passed = sum(1 for c in checks if c["status"] == "PASS")
    warned = sum(1 for c in checks if c["status"] == "WARN")
    failed = sum(1 for c in checks if c["status"] == "FAIL")
    score = int((passed + warned * 0.5) / len(checks) * 100) if checks else 0

    if action == "export_excel":
        wb = openpyxl.Workbook()
        header_fill = PatternFill(start_color="223043", end_color="223043", fill_type="solid")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        warn_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        fail_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        border = Border(left=Side(style='thin', color="CCCCCC"),
                        right=Side(style='thin', color="CCCCCC"),
                        top=Side(style='thin', color="CCCCCC"),
                        bottom=Side(style='thin', color="CCCCCC"))

        # Sheet 1: Checklist Verification
        ws1 = wb.active
        ws1.title = "SEO Checklist Audit"
        h1_cols = ["Sr. #", "Category", "Particular", "Status", "Detected Value", "Remark"]
        for c_idx, text in enumerate(h1_cols, 1):
            cell = ws1.cell(row=1, column=c_idx, value=text)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border

        for r_idx, chk in enumerate(checks, 2):
            vals = [r_idx - 1, chk["category"], chk["particular"], chk["status"], chk["value"], chk["remark"]]
            for c_idx, v in enumerate(vals, 1):
                cell = ws1.cell(row=r_idx, column=c_idx, value=v)
                cell.border = border
                if c_idx == 4:
                    cell.fill = pass_fill if v == "PASS" else (warn_fill if v == "WARN" else fail_fill)

        # Sheet 2: Collected Site Data (Old/Current Data)
        ws2 = wb.create_sheet("Collected Site Data")
        data_rows = [
            ("Audited URL", url),
            ("Audit Timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Overall Checklist Score", f"{score}%"),
            ("Meta Title", title_val),
            ("Title Length", t_len),
            ("Meta Description", desc_val),
            ("Description Length", d_len),
            ("Meta Keywords", kw_val),
            ("Canonical URL", canon_href),
            ("Hreflang Tags", ", ".join(hreflangs)),
            ("Robots Tag", rob_val),
            ("Schemas Detected", ", ".join(schema_types)),
            ("GTM IDs", ", ".join(gtm_ids)),
            ("Theme Color", tc_val),
            ("H1 Tags", " | ".join(h1s)),
            ("H2 Tags (Top 5)", " | ".join(h2s[:5])),
            ("Total Images", len(imgs)),
            ("Images Missing Alt", miss_alt),
            ("Images Missing Dimensions", miss_dim),
            ("Total Links", len(anchors)),
            ("Links Missing Title", miss_a_title),
        ]
        ws2.cell(row=1, column=1, value="Field").fill = header_fill
        ws2.cell(row=1, column=1).font = header_font
        ws2.cell(row=1, column=2, value="Extracted Value").fill = header_fill
        ws2.cell(row=1, column=2).font = header_font
        for r_idx, (k, v) in enumerate(data_rows, 2):
            ws2.cell(row=r_idx, column=1, value=k).border = border
            ws2.cell(row=r_idx, column=2, value=str(v)).border = border

        for ws in [ws1, ws2]:
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="full_seo_checklist_report.xlsx"
        )

    return jsonify({
        "url": url,
        "score": score,
        "summary": {"total": len(checks), "passed": passed, "warned": warned, "failed": failed},
        "checks": checks
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

