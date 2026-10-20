import io
import json
import logging
import os
import re
import smtplib
import tempfile
import time
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    url_for,
)
from werkzeug.utils import secure_filename

from services.excel_exporter import (
    generate_all_teachers_workbook,
    generate_classes_workbook,
    generate_single_teacher_workbook,
)
from services.excel_validator import validate_excel
from services.feasibility import TimetableInputError
from services.timetable_optimizer import generate_timetable, get_schedule_config

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("timetable_app")

app = Flask(__name__)

# Security & limits
app.config["SECRET_KEY"] = os.environ.get(
    "SECRET_KEY", "timetable-generator-key-2026-chirag"
)
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024  # 4 MB payload safety limit

BASE_DIR = Path(__file__).resolve().parent

# Ephemeral storage for temporary files (safe for serverless / read-only filesystems)
TEMP_DIR = Path(tempfile.gettempdir()) / "timetable_gen_cache"
try:
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
except Exception:
    TEMP_DIR = Path(tempfile.gettempdir())

SAMPLE_FOLDER = BASE_DIR / "sample_data"
SAMPLE_FILENAME = "timetable_input_sample.xlsx"

ALLOWED_EXTENSIONS = {"xlsx"}

TIMETABLE_CACHE_FILE = TEMP_DIR / "current_timetable.json"
CONFIG_CACHE_FILE = TEMP_DIR / "current_config.json"

# Contact, Author, and Customization settings
AUTHOR_NAME = "Chirag Deswal"
AUTHOR_CREDIT = "Created by Chirag Deswal."
DEFAULT_CONTACT_EMAIL = "chirag1710deswal@gmail.com"
CONTACT_EMAIL = os.environ.get("CONTACT_EMAIL", DEFAULT_CONTACT_EMAIL)
PORTFOLIO_URL = os.environ.get("PORTFOLIO_URL", "https://github.com/chiragdeswal")
GITHUB_URL = os.environ.get("GITHUB_URL", "https://github.com/chiragdeswal")
LINKEDIN_URL = os.environ.get(
    "LINKEDIN_URL", "https://www.linkedin.com/in/chirag-deswal/"
)
SHOW_SPONSOR_BANNER = os.environ.get("SHOW_SPONSOR_BANNER", "false").lower() in (
    "true",
    "1",
    "yes",
)
GOOGLE_SITE_VERIFICATION = os.environ.get("GOOGLE_SITE_VERIFICATION", "")
SITE_URL = os.environ.get("SITE_URL", "").rstrip("/")
try:
    SOLVER_TIME_LIMIT = int(os.environ.get("SOLVER_TIME_LIMIT", "8"))
except (ValueError, TypeError):
    SOLVER_TIME_LIMIT = 8

# SMTP configuration (optional)
SMTP_SERVER = os.environ.get("SMTP_SERVER", "")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587")) if os.environ.get("SMTP_PORT") else 587
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")


@app.context_processor
def inject_global_metadata():
    """Make metadata, author info and SEO parameters accessible to all templates."""
    req_url = request.url_root.rstrip("/")
    base_url = SITE_URL if SITE_URL else req_url
    canonical_url = f"{base_url}{request.path}"
    return {
        "author_name": AUTHOR_NAME,
        "author_credit": AUTHOR_CREDIT,
        "contact_email": CONTACT_EMAIL,
        "portfolio_url": PORTFOLIO_URL,
        "github_url": GITHUB_URL,
        "linkedin_url": LINKEDIN_URL,
        "show_sponsor_banner": SHOW_SPONSOR_BANNER,
        "google_site_verification": GOOGLE_SITE_VERIFICATION,
        "canonical_url": canonical_url,
        "site_url": base_url,
        "current_year": datetime.now(timezone.utc).year,
        "form_timestamp": int(time.time()),
    }


def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def build_class_grids(sessions):
    """Group sessions as class -> day -> period."""
    grids = {}
    for item in sessions:
        days = grids.setdefault(item["class_id"], {})
        periods = days.setdefault(item["day"], {})

        periods[item["period"]] = {"lesson": item, "skip": False}

        for extra in range(1, item.get("periods_per_session", 1)):
            periods[item["period"] + extra] = {"skip": True}

    return grids


def build_teacher_grids(sessions):
    """Group sessions as teacher -> day -> period with shared class deduplication."""
    grids = {}
    for item in sessions:
        teacher = item["teacher"]
        days = grids.setdefault(teacher, {})
        periods = days.setdefault(item["day"], {})

        if item["period"] in periods:
            existing = periods[item["period"]].get("lesson")
            if existing:
                classes_set = set(existing.get("class_ids", [existing["class_id"]]))
                classes_set.add(item["class_id"])
                existing["class_ids"] = list(classes_set)
                existing["display_class"] = ", ".join(sorted(classes_set))
            continue

        lesson_copy = dict(item)
        lesson_copy["class_ids"] = [item["class_id"]] + list(item.get("shared_with", []))
        lesson_copy["display_class"] = ", ".join(sorted(set(lesson_copy["class_ids"])))

        periods[item["period"]] = {"lesson": lesson_copy, "skip": False}

        for extra in range(1, item.get("periods_per_session", 1)):
            periods[item["period"] + extra] = {"skip": True}

    return grids


def compute_teacher_stats(sessions):
    """Calculate weekly workload statistics for each teacher."""
    stats = {}
    for item in sessions:
        t = item.get("teacher")
        if not t:
            continue
        if t not in stats:
            stats[t] = {
                "total_periods": 0,
                "subjects": set(),
                "classes": set(),
                "days": set(),
            }
        stats[t]["total_periods"] += item.get("periods_per_session", 1)
        stats[t]["subjects"].add(item.get("subject", ""))
        stats[t]["classes"].add(item.get("class_id", ""))
        if item.get("shared_with"):
            stats[t]["classes"].update(item.get("shared_with", []))
        stats[t]["days"].add(item.get("day", ""))

    for t, s in stats.items():
        s["subjects"] = sorted(list(s["subjects"]))
        s["classes"] = sorted(list(s["classes"]))
        s["days"] = sorted(list(s["days"]))
    return stats


def get_current_data(request_payload=None):
    """Retrieve sessions and schedule config from request payload, temp cache, or sample fallback."""
    sessions = None
    config = None

    if request_payload and isinstance(request_payload, dict):
        if "sessions" in request_payload and request_payload["sessions"]:
            sessions = request_payload["sessions"]
        if "config" in request_payload and request_payload["config"]:
            config = request_payload["config"]

    if sessions is None and TIMETABLE_CACHE_FILE.exists():
        try:
            with open(TIMETABLE_CACHE_FILE, "r", encoding="utf-8") as f:
                sessions = json.load(f)
        except Exception as e:
            logger.warning(f"Failed to read timetable cache: {e}")
            sessions = None

    if config is None and CONFIG_CACHE_FILE.exists():
        try:
            with open(CONFIG_CACHE_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception as e:
            logger.warning(f"Failed to read config cache: {e}")
            config = None

    # Fallback default configuration if nothing cached yet
    if not config:
        config = {
            "days": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
            "periods": [1, 2, 3, 4, 5, 6, 7],
            "rows": [
                {"type": "period", "period": 1, "start": "09:00", "end": "09:50"},
                {"type": "period", "period": 2, "start": "09:50", "end": "10:40"},
                {"type": "break", "label": "Lunch", "start": "10:40", "end": "11:30"},
                {"type": "period", "period": 3, "start": "11:30", "end": "12:20"},
                {"type": "period", "period": 4, "start": "12:20", "end": "13:10"},
                {"type": "period", "period": 5, "start": "13:10", "end": "14:00"},
                {"type": "period", "period": 6, "start": "14:00", "end": "14:50"},
                {"type": "period", "period": 7, "start": "14:50", "end": "15:40"},
            ],
        }

    return sessions or [], config


def process_timetable_file(file_path):
    """Validate and solve an Excel timetable file, returning processed structures or raising error."""
    errors = validate_excel(file_path)
    if errors:
        return False, "Excel validation failed:\n" + "\n".join(errors), None

    raw_timetable = generate_timetable(file_path, time_limit_seconds=SOLVER_TIME_LIMIT)
    if not raw_timetable:
        return False, "Validation passed, but the solver could not find a timetable that satisfies every rule. Try fewer weekly periods, more teachers for busy subjects, or more working days.", None

    config = get_schedule_config(file_path)
    for idx, s in enumerate(raw_timetable, start=1):
        if "id" not in s:
            s["id"] = f"session_{idx}"

    return True, "Timetable generated successfully!", {
        "raw_timetable": raw_timetable,
        "config": config,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    message = ""
    success = False
    timetable = None
    classes_timetable = {}
    teachers_timetable = {}
    teacher_stats = {}
    days = []
    rows = []
    periods = []
    classes = []
    teachers = []
    all_subjects = []

    if request.method == "POST":
        action = request.form.get("action", "upload")

        # Instant Sample Demo Mode
        if action == "demo":
            sample_path = SAMPLE_FOLDER / SAMPLE_FILENAME
            if sample_path.exists():
                try:
                    ok, msg, res = process_timetable_file(sample_path)
                    if ok:
                        raw_timetable = res["raw_timetable"]
                        config = res["config"]
                        days = config["days"]
                        rows = config["rows"]
                        periods = config.get("periods", [r["period"] for r in rows if r.get("type") == "period"])
                        timetable = raw_timetable
                        classes_timetable = build_class_grids(raw_timetable)
                        teachers_timetable = build_teacher_grids(raw_timetable)
                        teacher_stats = compute_teacher_stats(raw_timetable)
                        classes = sorted(list(classes_timetable.keys()))
                        teachers = sorted(list(teachers_timetable.keys()))
                        all_subjects = sorted(list(set(s["subject"] for s in raw_timetable)))

                        try:
                            with open(TIMETABLE_CACHE_FILE, "w", encoding="utf-8") as f:
                                json.dump(raw_timetable, f, indent=2)
                            with open(CONFIG_CACHE_FILE, "w", encoding="utf-8") as f:
                                json.dump(config, f, indent=2)
                        except Exception as ce:
                            logger.warning(f"Could not write cache files: {ce}")

                        message = f"Demo timetable loaded successfully! ({len(raw_timetable)} sessions, {len(classes)} classes, {len(teachers)} faculty members)"
                        success = True
                    else:
                        message = msg
                        success = False
                except Exception as e:
                    message = f"Error generating demo timetable: {e}"
                    success = False
            else:
                message = "Sample file not found on server."
                success = False

        # File Upload Mode
        else:
            file = request.files.get("file")
            if not file or file.filename == "":
                message = "Please select an Excel (.xlsx) file to upload."
            elif not allowed_file(file.filename):
                message = "Only .xlsx Excel files are allowed."
            else:
                temp_file = None
                try:
                    # Write to secure temporary file with guaranteed cleanup
                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                    temp_path = Path(temp_file.name)
                    temp_file.close()

                    file.save(temp_path)

                    ok, msg, res = process_timetable_file(temp_path)
                    if ok:
                        raw_timetable = res["raw_timetable"]
                        config = res["config"]
                        days = config["days"]
                        rows = config["rows"]
                        periods = config.get("periods", [r["period"] for r in rows if r.get("type") == "period"])

                        timetable = raw_timetable
                        classes_timetable = build_class_grids(raw_timetable)
                        teachers_timetable = build_teacher_grids(raw_timetable)
                        teacher_stats = compute_teacher_stats(raw_timetable)
                        classes = sorted(list(classes_timetable.keys()))
                        teachers = sorted(list(teachers_timetable.keys()))
                        all_subjects = sorted(list(set(s["subject"] for s in raw_timetable)))

                        # Persist to temp cache for direct download endpoints
                        try:
                            with open(TIMETABLE_CACHE_FILE, "w", encoding="utf-8") as f:
                                json.dump(raw_timetable, f, indent=2)
                            with open(CONFIG_CACHE_FILE, "w", encoding="utf-8") as f:
                                json.dump(config, f, indent=2)
                        except Exception as ce:
                            logger.warning(f"Could not write cache files: {ce}")

                        message = f"Timetable generated successfully! ({len(raw_timetable)} sessions scheduled for {len(classes)} classes and {len(teachers)} faculty members)"
                        success = True
                    else:
                        message = msg
                        success = False
                except TimetableInputError as e:
                    message = (
                        "A timetable is not possible with this data:\n"
                        + "\n".join(f"- {problem}" for problem in e.problems)
                    )
                    success = False
                except Exception as e:
                    message = f"Error while generating timetable: {e}"
                    success = False
                finally:
                    # Guarantee ephemeral cleanup
                    if temp_file and Path(temp_file.name).exists():
                        try:
                            Path(temp_file.name).unlink()
                        except Exception:
                            pass

    return render_template(
        "index.html",
        message=message,
        success=success,
        timetable=timetable,
        classes_timetable=classes_timetable,
        teachers_timetable=teachers_timetable,
        teacher_stats=teacher_stats,
        days=days,
        rows=rows,
        periods=periods,
        classes=classes,
        teachers=teachers,
        all_subjects=all_subjects,
    )


@app.route("/contact", methods=["POST"])
def contact():
    """Handle contact form with server-side validation and anti-spam protection."""
    is_ajax = request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.is_json
    data = request.get_json(silent=True) if request.is_json else request.form

    name = data.get("name", "").strip()
    email = data.get("email", "").strip()
    subject = data.get("subject", "").strip() or "TimelyAI Inquiry"
    message = data.get("message", "").strip()

    # Spam Protection Check 1: Honeypot field (hidden from genuine users)
    honeypot = data.get("company_website", "").strip()
    if honeypot:
        logger.warning(f"Spam detected (honeypot triggered): name={name}, email={email}")
        resp_msg = "Your inquiry has been received."
        if is_ajax:
            return jsonify({"success": True, "message": resp_msg})
        flash(resp_msg, "success")
        return redirect(url_for("index") + "#contact")

    # Spam Protection Check 2: Submission speed / token verification
    form_ts = data.get("form_ts", "")
    if form_ts:
        try:
            elapsed = time.time() - float(form_ts)
            if elapsed < 1.2:
                logger.warning(f"Spam detected (instant submission in {elapsed:.2f}s)")
                resp_msg = "Please take a moment before submitting."
                if is_ajax:
                    return jsonify({"success": False, "error": resp_msg}), 400
                flash(resp_msg, "error")
                return redirect(url_for("index") + "#contact")
        except (ValueError, TypeError):
            pass

    # Server-side validation
    errors = []
    if not name or len(name) < 2 or len(name) > 100:
        errors.append("Please provide your full name (2-100 characters).")

    email_regex = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    if not email or not re.match(email_regex, email):
        errors.append("Please provide a valid email address.")

    if not message or len(message) < 10 or len(message) > 4000:
        errors.append("Message must be between 10 and 4,000 characters.")

    if errors:
        error_msg = " ".join(errors)
        if is_ajax:
            return jsonify({"success": False, "error": error_msg}), 400
        flash(error_msg, "error")
        return redirect(url_for("index") + "#contact")

    # Attempt email delivery if SMTP configured; otherwise log inquiry safely
    email_sent = False
    if SMTP_SERVER and SMTP_USER and SMTP_PASSWORD:
        try:
            msg = MIMEMultipart()
            msg["From"] = SMTP_USER
            msg["To"] = CONTACT_EMAIL
            msg["Subject"] = f"[TimelyAI Contact] {subject} - from {name}"
            msg["Reply-To"] = email

            body = (
                f"New inquiry received from TimelyAI Website:\n\n"
                f"Name: {name}\n"
                f"Email: {email}\n"
                f"Subject: {subject}\n"
                f"Timestamp: {datetime.now(timezone.utc).isoformat()}\n\n"
                f"Message:\n{message}\n"
            )
            msg.attach(MIMEText(body, "plain"))

            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)
            email_sent = True
            logger.info(f"Contact email sent successfully from {email}")
        except Exception as e:
            logger.error(f"Failed to send email via SMTP: {e}")

    if not email_sent:
        logger.info(
            f"Contact inquiry logged: Name: {name} | Email: {email} | Subject: {subject} | Msg: {message[:100]}..."
        )

    success_msg = f"Thank you, {name}! Your message has been received. You can also reach Chirag directly at {CONTACT_EMAIL}."
    if is_ajax:
        return jsonify({"success": True, "message": success_msg, "email_sent": email_sent})
    flash(success_msg, "success")
    return redirect(url_for("index") + "#contact")


@app.route("/api/save-timetable", methods=["POST"])
def save_timetable():
    """Receive updated timetable sessions from client-side editor and save to temp cache."""
    try:
        data = request.get_json(force=True)
        sessions = data.get("sessions", [])

        for idx, s in enumerate(sessions, start=1):
            if "id" not in s:
                s["id"] = f"session_{idx}"

        try:
            with open(TIMETABLE_CACHE_FILE, "w", encoding="utf-8") as f:
                json.dump(sessions, f, indent=2)
        except Exception as e:
            logger.warning(f"Could not persist cache: {e}")

        return jsonify({
            "success": True,
            "message": "Timetable saved successfully.",
            "count": len(sessions),
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/download/main-excel", methods=["GET", "POST"])
def download_main_excel():
    """Download the main/master timetable as an Excel workbook."""
    payload = request.get_json(silent=True) if request.is_json else None
    sessions, config = get_current_data(payload)
    if not sessions:
        # Fallback to generating sample data
        sample_path = SAMPLE_FOLDER / SAMPLE_FILENAME
        if sample_path.exists():
            try:
                sessions = generate_timetable(sample_path, time_limit_seconds=SOLVER_TIME_LIMIT) or []
                config = get_schedule_config(sample_path)
            except Exception:
                pass

    buf = generate_classes_workbook(sessions, config)
    return send_file(
        buf,
        as_attachment=True,
        download_name="College_Timetable_Master.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download/teacher-excel/<teacher_name>", methods=["GET", "POST"])
def download_teacher_excel(teacher_name):
    """Download an individual teacher's timetable as an Excel workbook."""
    payload = request.get_json(silent=True) if request.is_json else None
    sessions, config = get_current_data(payload)
    if not sessions:
        sample_path = SAMPLE_FOLDER / SAMPLE_FILENAME
        if sample_path.exists():
            try:
                sessions = generate_timetable(sample_path, time_limit_seconds=SOLVER_TIME_LIMIT) or []
                config = get_schedule_config(sample_path)
            except Exception:
                pass

    buf = generate_single_teacher_workbook(teacher_name, sessions, config)
    safe_name = "".join(c for c in teacher_name if c.isalnum() or c in (" ", "_", "-")).strip()
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"Timetable_Faculty_{safe_name}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download/all-teachers-excel", methods=["GET", "POST"])
def download_all_teachers_excel():
    """Download all teachers' individual timetables combined in one workbook."""
    payload = request.get_json(silent=True) if request.is_json else None
    sessions, config = get_current_data(payload)
    if not sessions:
        sample_path = SAMPLE_FOLDER / SAMPLE_FILENAME
        if sample_path.exists():
            try:
                sessions = generate_timetable(sample_path, time_limit_seconds=SOLVER_TIME_LIMIT) or []
                config = get_schedule_config(sample_path)
            except Exception:
                pass

    buf = generate_all_teachers_workbook(sessions, config)
    return send_file(
        buf,
        as_attachment=True,
        download_name="Timetable_All_Faculty.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/sample-template")
def download_sample():
    """Download the official Excel input template."""
    return send_from_directory(
        SAMPLE_FOLDER,
        SAMPLE_FILENAME,
        as_attachment=True,
        download_name="timetable_input_template.xlsx",
    )


@app.route("/robots.txt")
def robots_txt():
    """Dynamic robots.txt for search engines."""
    domain = (SITE_URL if SITE_URL else request.url_root).rstrip("/")
    content = f"User-agent: *\nAllow: /\n\nSitemap: {domain}/sitemap.xml\n"
    return app.response_class(content, mimetype="text/plain")


@app.route("/sitemap.xml")
def sitemap_xml():
    """Dynamic XML sitemap for SEO."""
    domain = (SITE_URL if SITE_URL else request.url_root).rstrip("/")
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>{domain}/</loc>
    <lastmod>{now}</lastmod>
    <changefreq>weekly</changefreq>
    <priority>1.0</priority>
  </url>
  <url>
    <loc>{domain}/sample-template</loc>
    <lastmod>{now}</lastmod>
    <changefreq>monthly</changefreq>
    <priority>0.8</priority>
  </url>
</urlset>"""
    return app.response_class(xml, mimetype="application/xml")


@app.errorhandler(413)
def request_entity_too_large(error):
    flash("The uploaded file is too large. Excel workbooks must be under 4MB.", "error")
    return redirect(url_for("index"))


@app.errorhandler(404)
def page_not_found(error):
    return render_template("404.html"), 404


@app.errorhandler(500)
def internal_server_error(error):
    logger.error(f"Internal server error: {error}")
    return render_template("500.html"), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
