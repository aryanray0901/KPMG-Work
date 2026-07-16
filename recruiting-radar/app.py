"""
Recruiting Radar
----------------
Everything that happens before a consulting interview: knowing the real
deadlines (including the early diversity/sophomore ones most students
miss entirely), understanding the real numbers behind the process,
networking effectively without an existing network, and getting a resume
past the screen. Deliberately does not cover case interview practice,
frameworks, or fit questions, Case Prep already covers that.

All data is real and sourced (see content/ modules and the README), not
generated. The deadline countdowns are computed live against the actual
current date every time the page loads, which is what "live" means here:
genuinely current status without needing any external API.

Single self-contained Flask application. No external services.
"""

import os
import re
import time
from datetime import date

from flask import Flask, request, render_template, redirect, url_for, session, jsonify

from content.deadlines import DEADLINES, deadlines_sorted, days_until
from content.stats import REALITY_STATS
from content.networking import (
    NETWORKING_TECHNIQUES, COLD_OUTREACH_TEMPLATES, COFFEE_CHAT_QUESTIONS, RESPONSE_RATE_GUIDANCE,
)
from content.resume_rules import RESUME_MISTAKES
from content.nontarget_playbook import NONTARGET_PLAYBOOK

app = Flask(__name__)
app.secret_key = "recruiting-radar-local-secret"


def format_date(d):
    """'%-d' is a Unix-only strftime extension and raises on Windows;
    this achieves the same 'July 16, 2026' (no leading zero) formatting
    in a way that works on every platform."""
    return f"{d.strftime('%B')} {d.day}, {d.year}"


app.jinja_env.filters["fmtdate"] = format_date


@app.after_request
def _no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response


# ---------------------------------------------------------------------------
# RESUME CHECKER (rule-based, transparent about what it checks and why)
# ---------------------------------------------------------------------------

FIRST_PERSON_RE = re.compile(r"\bI\b")
WE_RE = re.compile(r"\bwe\b", re.IGNORECASE)
NUMBER_RE = re.compile(r"\d")
SECTION_KEYWORDS = {
    "education": ["education", "university", "college", "gpa"],
    "experience": ["experience", "internship", "analyst", "associate"],
    "leadership": ["leadership", "extracurricular", "activities", "organization", "club"],
    "skills": ["skills", "certifications", "languages", "software"],
}
WEAK_VERBS = ["helped", "worked on", "was responsible for", "assisted", "participated in"]


def check_resume(text):
    """Returns a list of {check, status ('pass'/'flag'), detail} dicts.
    Pattern-based, not a simulation of human judgment, transparent about
    exactly what it's checking so the limits are obvious."""
    results = []
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    word_count = len(text.split())

    # Length check
    if word_count == 0:
        return [{"check": "Length", "status": "flag", "detail": "No text detected."}]
    if word_count > 650:
        results.append({"check": "Length", "status": "flag",
                         "detail": f"{word_count} words. A one-page resume is typically 400-600 words; "
                                   f"this reads long enough that it may be running past one page."})
    elif word_count < 150:
        results.append({"check": "Length", "status": "flag",
                         "detail": f"{word_count} words. This looks short for a full resume, make sure "
                                   f"this is the complete document."})
    else:
        results.append({"check": "Length", "status": "pass",
                         "detail": f"{word_count} words, a reasonable range for a one-page resume."})

    # I vs we
    i_count = len(FIRST_PERSON_RE.findall(text))
    we_count = len(WE_RE.findall(text))
    if we_count > i_count and we_count >= 2:
        results.append({"check": "Individual ownership ('I' vs 'we')", "status": "flag",
                         "detail": f"Found {we_count} instance(s) of 'we' versus {i_count} of 'I'. "
                                   f"Recruiters need to see your individual contribution; consider "
                                   f"rewriting team-credit sentences around what you specifically did."})
    else:
        results.append({"check": "Individual ownership ('I' vs 'we')", "status": "pass",
                         "detail": f"{we_count} instance(s) of 'we' versus {i_count} of 'I' (or bullets "
                                   f"structured without either, which is normal for resume fragments)."})

    # Quantification: what share of bullet-like lines contain a number
    bullet_lines = [l for l in lines if len(l) > 15 and not l.isupper()]
    quantified = [l for l in bullet_lines if NUMBER_RE.search(l)]
    if bullet_lines:
        pct = len(quantified) / len(bullet_lines) * 100
        if pct < 40:
            results.append({"check": "Quantified impact", "status": "flag",
                             "detail": f"Only about {pct:.0f}% of your lines contain a number. Strong "
                                       f"consulting resumes quantify most bullets (%, $, headcount, "
                                       f"time saved), 'improved a process' is much weaker than "
                                       f"'cut process time 30%.'"})
        else:
            results.append({"check": "Quantified impact", "status": "pass",
                             "detail": f"About {pct:.0f}% of your lines contain a number, a solid sign "
                                       f"most bullets are quantified rather than just descriptive."})

    # Weak verb / passive framing check
    weak_hits = [w for w in WEAK_VERBS if w in text.lower()]
    if weak_hits:
        results.append({"check": "Verb strength", "status": "flag",
                         "detail": f"Found weak framing ({', '.join(weak_hits)}). These describe "
                                   f"proximity to work rather than ownership of it; consider stronger "
                                   f"action verbs (led, built, negotiated, redesigned) if you actually "
                                   f"drove the outcome."})
    else:
        results.append({"check": "Verb strength", "status": "pass",
                         "detail": "No common weak/passive phrasing detected."})

    # Section presence (heuristic keyword search, not a layout parser)
    text_lower = text.lower()
    missing_sections = [name for name, kws in SECTION_KEYWORDS.items()
                         if not any(kw in text_lower for kw in kws)]
    if missing_sections:
        results.append({"check": "Section coverage", "status": "flag",
                         "detail": f"Didn't detect clear signals of: {', '.join(missing_sections)}. "
                                   f"This is a keyword search, not a layout parser, so double-check "
                                   f"manually, but a missing leadership/extracurricular section "
                                   f"specifically is one of the most commonly cited resume mistakes."})
    else:
        results.append({"check": "Section coverage", "status": "pass",
                         "detail": "Detected signals of education, experience, leadership, and skills content."})

    return results


# ---------------------------------------------------------------------------
# ROUTES
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    today = date.today()
    upcoming = [d for d in deadlines_sorted() if d["date"] >= today][:5]
    return render_template("index.html", today=today, upcoming=upcoming,
                            deadline_count=len(DEADLINES), stat_count=len(REALITY_STATS))


@app.route("/deadlines", methods=["GET"])
def deadlines_page():
    today = date.today()
    tier_filter = request.args.get("tier", "")
    type_filter = request.args.get("type", "")
    items = deadlines_sorted()
    if tier_filter:
        items = [d for d in items if d["tier"] == tier_filter]
    if type_filter:
        items = [d for d in items if d["type"] == type_filter]
    tiers = sorted(set(d["tier"] for d in DEADLINES))
    types = sorted(set(d["type"] for d in DEADLINES))
    return render_template("deadlines.html", items=items, today=today, tiers=tiers, types=types,
                            active_tier=tier_filter, active_type=type_filter, days_until=days_until)


@app.route("/stats", methods=["GET"])
def stats_page():
    return render_template("stats.html", stats=REALITY_STATS)


@app.route("/networking", methods=["GET"])
def networking_page():
    contacts = session.get("contacts", [])
    return render_template("networking.html", techniques=NETWORKING_TECHNIQUES,
                            templates=COLD_OUTREACH_TEMPLATES, questions=COFFEE_CHAT_QUESTIONS,
                            guidance=RESPONSE_RATE_GUIDANCE, contacts=contacts)


@app.route("/networking/add", methods=["POST"])
def networking_add():
    contacts = session.get("contacts", [])
    contacts.append({
        "name": request.form.get("name", "").strip(),
        "firm": request.form.get("firm", "").strip(),
        "status": request.form.get("status", "Reached out"),
        "notes": request.form.get("notes", "").strip(),
    })
    session["contacts"] = contacts[-40:]  # keep the session cookie small
    return redirect(url_for("networking_page"))


@app.route("/networking/clear", methods=["POST"])
def networking_clear():
    session["contacts"] = []
    return redirect(url_for("networking_page"))


@app.route("/resume-check", methods=["GET"])
def resume_check_page():
    return render_template("resume_check.html", mistakes=RESUME_MISTAKES, results=None)


@app.route("/api/resume-check", methods=["POST"])
def api_resume_check():
    data = request.get_json(force=True)
    text = (data.get("text") or "").strip()
    if not text:
        return jsonify({"error": "Paste your resume text first."}), 400
    results = check_resume(text)
    return jsonify({"results": results})


@app.route("/nontarget", methods=["GET"])
def nontarget_page():
    return render_template("nontarget.html", playbook=NONTARGET_PLAYBOOK)


@app.route("/tracker", methods=["GET"])
def tracker_page():
    applications = session.get("applications", [])
    return render_template("tracker.html", applications=applications, deadlines=deadlines_sorted())


@app.route("/tracker/add", methods=["POST"])
def tracker_add():
    applications = session.get("applications", [])
    applications.append({
        "firm": request.form.get("firm", "").strip(),
        "program": request.form.get("program", "").strip(),
        "status": request.form.get("status", "Planning to apply"),
        "notes": request.form.get("notes", "").strip(),
    })
    session["applications"] = applications[-40:]
    return redirect(url_for("tracker_page"))


@app.route("/tracker/clear", methods=["POST"])
def tracker_clear():
    session["applications"] = []
    return redirect(url_for("tracker_page"))


def _open_browser():
    import webbrowser
    import sys as _sys
    time.sleep(1.2)
    url = "http://127.0.0.1:5110"
    chrome_candidates = []
    if _sys.platform == "darwin":
        chrome_candidates.append("open -a 'Google Chrome' %s")
    elif _sys.platform.startswith("win"):
        for p in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                  r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"):
            if os.path.exists(p):
                chrome_candidates.append(p.replace("\\", "\\\\") + " %s")
    else:
        chrome_candidates += ["google-chrome %s", "chromium-browser %s", "chromium %s"]
    for template in chrome_candidates:
        try:
            webbrowser.get(template).open(url)
            return
        except webbrowser.Error:
            continue
    webbrowser.open(url)


if __name__ == "__main__":
    import sys
    import threading
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        threading.Thread(target=_open_browser, daemon=True).start()
    print("Starting Recruiting Radar at http://127.0.0.1:5110")
    app.run(host="127.0.0.1", port=5110, debug=("--debug" in sys.argv))
