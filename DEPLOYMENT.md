# TimelyAI Deployment & Production Guide

This guide provides end-to-end instructions for deploying the **TimelyAI College Timetable Generator & Faculty Schedule System** to **Vercel**, configuring custom domains, managing environment variables, and troubleshooting production workloads.

---

## 1. Architecture & Serverless Readiness Audit

TimelyAI has been audited and engineered specifically for serverless Python execution on Vercel:

| Dimension | Constraint on Vercel | TimelyAI Implementation | Status |
| :--- | :--- | :--- | :--- |
| **Filesystem** | Root directory is strictly read-only; only `/tmp` is writable and ephemeral | Uses `tempfile.NamedTemporaryFile` in OS temp storage (`/tmp`), with immediate `try...finally` cleanup. Endpoints never depend on disk persistence. | ✅ Passed |
| **Execution Limits** | Hobby: 10–15s max<br>Pro: up to 300s (`maxDuration: 60`) | OR-Tools CP-SAT solver uses `SOLVER_TIME_LIMIT` (default: 8s), generating conflict-free solutions well within serverless timeouts. | ✅ Passed |
| **Bundle Size** | 250 MB uncompressed limit | Dependencies (`ortools` ~81MB, `pandas`+`numpy` ~98MB, `openpyxl` ~2MB, `Flask` ~8MB = ~190MB). Kept within the 250MB threshold; `.vercelignore` strips unnecessary files. | ✅ Passed |
| **Upload Payload** | 4.5 MB serverless body limit | Enforces `MAX_CONTENT_LENGTH = 4 * 1024 * 1024` with user-friendly 413 error handlers. Standard timetable workbooks are ~20KB–300KB. | ✅ Passed |
| **Excel Downloads** | No shared disk across serverless invocations | Client-side export buttons send timetable sessions directly via POST; server constructs workbooks in-memory via `io.BytesIO` and streams them back instantly. | ✅ Passed |
| **PDF Export** | Heavy headless Chrome/Weasyprint crashes in serverless | Client-side `@media print` landscape engine with dedicated print stylesheet and browser PDF dialog. Zero server memory overhead. | ✅ Passed |

---

## 2. Deploying to Vercel

### Method A: Deploy via GitHub (Recommended)

1. **Initialize Git & Push to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit of production TimelyAI timetable generator"
   git branch -M main
   git remote add origin https://github.com/chirag17deswal-maker/Automatic-Timetable-Generator.git
   git push -u origin main
   ```

2. **Import Project into Vercel:**
   - Go to [vercel.com](https://vercel.com) and log in.
   - Click **Add New...** → **Project**.
   - Select your GitHub repository (`Automatic-Timetable-Generator`).

3. **Configure Project Settings:**
   - **Framework Preset:** Leave as `Other`.
   - **Root Directory:** Leave as `./` (default).
   - Vercel automatically detects `vercel.json`, `api/index.py`, and `requirements.txt`.

4. **Add Environment Variables (Optional):**
   - Under the **Environment Variables** section, add your desired values (see table below).
   - Click **Deploy**.

5. **Access Your Live Application:**
   - Once deployment finishes (typically 1–2 minutes), Vercel assigns a live domain:
     `https://your-project-name.vercel.app`

---

### Method B: Deploy via Vercel CLI

1. **Install Vercel CLI globally:**
   ```bash
   npm install -g vercel
   ```

2. **Login to Vercel:**
   ```bash
   vercel login
   ```

3. **Deploy from project directory:**
   ```bash
   cd c:\Users\chira\Downloads\timetable-generator-updated
   vercel --prod
   ```
   Follow the interactive prompts:
   - Set up and deploy: **Yes**
   - Which scope: **[Your Account]**
   - Link to existing project: **No**
   - Project name: **timelyai-timetable-generator**
   - In which directory is your code located: **.**

---

## 3. Connecting a Custom Domain

To connect your own domain (e.g. `timetable.yourcollege.edu` or `timelyai.com`):

1. **Add Domain in Vercel:**
   - Navigate to your project on the [Vercel Dashboard](https://vercel.com).
   - Go to **Settings** → **Domains**.
   - Enter your domain or subdomain (e.g., `timelyai.com` or `schedule.college.edu`) and click **Add**.

2. **Configure DNS Records at your Domain Registrar (Namecheap, GoDaddy, Cloudflare, etc.):**
   - **For Apex Domain (`example.com`):**
     - Type: `A`
     - Name: `@`
     - Value: `76.76.21.21`
   - **For Subdomain (`schedule.example.com`):**
     - Type: `CNAME`
     - Name: `schedule`
     - Value: `cname.vercel-dns.com`

3. **SSL Verification:**
   - Vercel automatically provisions and renews free Let's Encrypt SSL certificates within minutes of DNS propagation.

---

## 4. Environment Variables Reference

Configure these in the Vercel Dashboard under **Project Settings → Environment Variables**:

| Variable Name | Default Value | Description |
| :--- | :--- | :--- |
| `CONTACT_EMAIL` | `chirag1710deswal@gmail.com` | Destination email address for the contact form and public display. |
| `PORTFOLIO_URL` | `https://github.com/chiragdeswal` | Link to author's portfolio or personal site. |
| `GITHUB_URL` | `https://github.com/chiragdeswal` | Link to GitHub profile. |
| `LINKEDIN_URL` | `https://www.linkedin.com/in/chirag-deswal/` | Link to LinkedIn profile. |
| `SOLVER_TIME_LIMIT` | `8` | Maximum seconds OR-Tools solver runs before returning a feasible schedule. Keep at `8`–`10` for Vercel Hobby plan. Can increase to `30`–`60` on Pro plan. |
| `SHOW_SPONSOR_BANNER` | `false` | Set to `true` to enable the top sponsored promotion banner. Stays completely hidden by default. |
| `GOOGLE_SITE_VERIFICATION` | `""` | Google Search Console verification meta tag token (e.g. `google-site-verification=XYZ123`). |
| `SITE_URL` | `""` | Canonical base URL (e.g. `https://timelyai.vercel.app`) for sitemap and OpenGraph tags. |
| `SECRET_KEY` | `(auto-fallback)` | Flask session encryption key. Set to a random secure string. |
| `SMTP_SERVER` | `""` | (Optional) Outgoing SMTP host (e.g., `smtp.gmail.com` or `smtp.sendgrid.net`). |
| `SMTP_PORT` | `587` | (Optional) Outgoing SMTP port. |
| `SMTP_USER` | `""` | (Optional) SMTP username/email. |
| `SMTP_PASSWORD` | `""` | (Optional) SMTP app password. |

> **Note on Email Inquiries:** If SMTP credentials are not configured, contact form submissions are safely logged in the serverless console, and the visitor receives a direct 1-click `mailto:` follow-up link pre-filled with their details.

---

## 5. Google Search Console & SEO Verification

1. Go to [Google Search Console](https://search.google.com/search-console).
2. Choose **URL prefix** and enter your production URL (e.g. `https://your-domain.vercel.app`).
3. Select **HTML tag** verification method.
4. Copy the `content="..."` value (e.g. `4abcdefg123456789`).
5. In Vercel Project Settings, add an environment variable:
   - Key: `GOOGLE_SITE_VERIFICATION`
   - Value: `4abcdefg123456789`
6. Redeploy or promote deployment. Return to Google Search Console and click **Verify**.
7. Submit your sitemap in Search Console:
   - URL: `https://your-domain.vercel.app/sitemap.xml`

---

## 6. Pre-Deployment & Post-Deployment Checklist

- [x] **File Uploads:** Tested with sample Excel file (`sample_data/timetable_input_sample.xlsx`).
- [x] **Solver Time Limits:** Solver configured with `SOLVER_TIME_LIMIT=8`s to prevent Vercel Hobby 10-15s timeouts.
- [x] **Ephemeral Storage:** Temporary files written to `/tmp` and removed in `finally:` blocks.
- [x] **Excel Exports:** Tested Master Timetable, All Faculty Timetables, and Individual Faculty Timetable downloads.
- [x] **PDF Export:** Tested browser landscape `@media print` layout and styling.
- [x] **Responsive Mobile Layout:** Verified table horizontal scrolling, sticky periods, and mobile navigation drawer.
- [x] **Contact Form:** Tested server-side validation, anti-spam honeypot, and timestamp checks.
- [x] **SEO:** Validated `/robots.txt`, `/sitemap.xml`, OpenGraph tags, and Schema.org structured data.
- [x] **Monetization & Credits:** Credit "Created by Chirag Deswal" included; Contact email `chirag1710deswal@gmail.com` configured; sponsored banner toggleable via `SHOW_SPONSOR_BANNER`.

---

## 7. Troubleshooting Common Issues

### Issue 1: Vercel Function Invocation Timeout (504 Gateway Timeout)
- **Cause:** Excel dataset has too many classes or tight constraints that cause the solver to exceed Vercel's default 10s Hobby limit.
- **Solution:**
  1. Ensure `SOLVER_TIME_LIMIT` environment variable is set to `8`.
  2. If using Vercel Pro, set `"maxDuration": 60` in `vercel.json` and increase `SOLVER_TIME_LIMIT` to `30`.
  3. Check the Feasibility Guide in the app to make sure faculty aren't assigned more periods than available working slots.

### Issue 2: HTTP 413 "Request Entity Too Large"
- **Cause:** Uploaded file exceeds Vercel's 4.5MB serverless payload limit.
- **Solution:** College timetable spreadsheets should only contain raw text and numbers, usually under 200KB. Ensure the file does not contain embedded high-resolution images or unused media.

### Issue 3: Read-only Filesystem Error (`[Errno 30] Read-only file system`)
- **Cause:** Attempting to write files to the application root directory in Vercel.
- **Solution:** TimelyAI uses `tempfile.gettempdir()` for all caching and parsing. If you write custom code, never write to relative paths like `./uploads`.
