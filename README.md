# ⚡ SEO & Web Tools Suite — All-in-One Growth Platform

A unified, modern web application combining 6 high-impact SEO, web auditing, and growth tools into one cohesive interface with a futuristic dark theme and zero hosting cost.

Designed for **100% Free Deployment on Vercel** (Static Frontend + Python Serverless Functions).

---

## 🛠️ The 6 Integrated Tools

| Tool | Capability | Exports |
| :--- | :--- | :--- |
| **⚡ Meta Scraper** | Bulk URL analysis for Title, Meta Description, Length checks, OG & Twitter cards | Excel (`.xlsx`) |
| **🗺️ Image Sitemap Crawler** | Deep crawl of internal pages, extracts WebP, responsive `srcset`, and `data-src` images | Google XML Sitemap + CSV |
| **🔍 SEO & Image Auditor** | On-page diagnostics for missing alt tags, missing dimensions (CLS), lazy loading, and headings | Health Score + Table |
| **🔗 Backlink Finder** | High-converting search footprint generator for Guest Posts, Directories, Classifieds & Communities with country filtering | Live search URLs & scraped candidates |
| **📄 Schema.org Generator** | Generates valid `Schema.org/Article` and `WebPage` structured data for Google rich snippets | 1-Click Code Copy |
| **🎬 Shorts Hunter** | Discovers viral YouTube Shorts in any niche with view counts, channels, and engagement metrics | CSV |

---

## 🚀 Quick Start (Run Locally)

### Windows 1-Click
Simply double-click:
```
run_local.bat
```
This automatically verifies dependencies and launches your browser at `http://localhost:5000`.

### Manual Command Line
```bash
# 1. Install dependencies
pip install -r api/requirements.txt

# 2. Run the application
python api/index.py
```
Open [http://localhost:5000](http://localhost:5000) in your browser.

---

## ☁️ Deploying to Vercel for FREE ($0/Month)

### Method 1: Using GitHub (Recommended)
1. Initialize a git repository and push this folder to GitHub:
   ```bash
   git init
   git add .
   git commit -m "Initial commit of unified SEO suite"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/seo-tools-suite.git
   git push -u origin main
   ```
2. Go to [vercel.com](https://vercel.com) and log in.
3. Click **Add New...** → **Project** and import your GitHub repository.
4. Keep all default build settings (Vercel automatically detects `vercel.json` and Python `/api`).
5. Click **Deploy**. Your unified suite is now live with a free `.vercel.app` URL and automatic HTTPS!

### Method 2: Using Vercel CLI
```bash
# 1. Install Vercel CLI (if not already installed)
npm install -g vercel

# 2. Deploy directly from this directory
vercel
```

---

## 📁 Project Structure

```
unified-tools-suite/
├── api/
│   ├── index.py              # Serverless Python API routes for all 6 tools
│   └── requirements.txt      # Python dependencies (Flask, bs4, openpyxl, etc.)
├── public/
│   ├── index.html            # Main unified dashboard UI
│   ├── style.css             # Dark neo-modern styling & typography
│   └── app.js                # Frontend controllers & export handlers
├── index.html                # Root mirror for static hosts
├── style.css                 # Root mirror
├── app.js                    # Root mirror
├── vercel.json               # Vercel configuration for routing
├── run_local.bat             # 1-click Windows local launcher
└── README.md                 # Documentation
```
