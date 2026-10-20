# TimelyAI — College Timetable & Faculty Schedule System

[![Python 3.11](https://img.shields.io/badge/Python-3.11-blue.svg)](https://python.org)
[![OR-Tools](https://img.shields.io/badge/Google-OR--Tools%20CP--SAT-red.svg)](https://developers.google.com/optimization)
[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new)

**TimelyAI** is a public-facing, web-based College Timetable Generator and Faculty Workload Management system built with Python, Flask, and Google OR-Tools. It converts multi-department academic requirements from Excel workbooks into 100% clash-free schedules in seconds.

**Created by Chirag Deswal.**  
**Contact:** [chirag1710deswal@gmail.com](mailto:chirag1710deswal@gmail.com)

---

## 🌟 Key Features

- **Google OR-Tools CP-SAT Solver:** Solves combinatorial scheduling constraints mathematically (zero double-bookings for faculty or classes).
- **Multi-Period Lab & Workshop Support:** Allocates continuous 2-period and 3-period sessions seamlessly without splitting over breaks.
- **Shared / Combined Branch Lectures:** Merges multiple branches taking common courses under one teacher at the same period.
- **Balanced Workload & Teacher Limits:** Spreads lectures evenly across the week and respects maximum daily and weekly teaching loads.
- **Interactive In-Browser Customization:** Real-time modal editor allowing instant adjustments with visual clash warnings.
- **Multi-Sheet Excel Export:** Downloads Master Timetable (summary + all classes) and individual Faculty Timetable workbooks.
- **Clean Landscape PDF / Print Export:** Dedicated `@media print` engine formatted for A4 landscape printing and PDF saving.
- **Serverless & Vercel Ready:** Stateless, ephemeral `/tmp` execution, sub-10s solving limits, in-memory streaming, and zero local filesystem dependencies.
- **Built-in Anti-Spam Contact Form:** Honeypot and timestamp protection for collaboration and consulting inquiries.
- **SEO & Google Search Console:** Complete metadata, OpenGraph cards, Twitter cards, dynamic `robots.txt`, and XML `sitemap.xml`.

---

## 🚀 Quick Start (Local Development)

### 1. Clone & Set Up Virtual Environment

```bash
# Clone the repository
git clone https://github.com/chirag17deswal-maker/Automatic-Timetable-Generator.git
cd Automatic-Timetable-Generator

# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Run the Application

```bash
python app.py
```

Open your browser at `http://localhost:5000`.

---

## 📊 Excel Template Structure

TimelyAI expects a `.xlsx` workbook containing 4 mandatory sheets (and 1 optional sheet):

1. **Classes:** `course_name`, `branch_name`, `semester`, `section`
2. **Subjects:** `class_id`, `subject_name`, `weekly_periods`, `periods_per_session`, `subject_type`
3. **Teachers:** `teacher_name`, `subjects` (comma-separated subjects qualified to teach)
4. **Settings:** `working_days`, `periods_per_day`, `period_start_time`, `period_duration`, `lunch_break_after`, `lunch_duration`, `allow_gaps`
5. **Common_Classes (Optional):** `group_name`, `class_ids`, `subject_name`, `weekly_periods`, `periods_per_session`

You can download the pre-filled sample template directly from the website or from `sample_data/timetable_input_sample.xlsx`.

---

## ☁️ Deployment on Vercel

TimelyAI is pre-configured for Vercel deployment with `vercel.json` and `api/index.py`.

See [DEPLOYMENT.md](DEPLOYMENT.md) for detailed deployment steps, custom domain setup, and environment variables.

---

## 💼 Work With Me / Consulting

Need custom university scheduling constraints (classroom capacity, professor availability preferences, student elective clustering), ERP integration, or operations research consulting?

- **Author:** Chirag Deswal
- **Email:** [chirag1710deswal@gmail.com](mailto:chirag1710deswal@gmail.com)
- **GitHub:** [https://github.com/chiragdeswal](https://github.com/chiragdeswal)
- **LinkedIn:** [https://www.linkedin.com/in/chirag-deswal/](https://www.linkedin.com/in/chirag-deswal/)

---

## 📄 License

Created by Chirag Deswal. Built for educational and academic institutions worldwide.
