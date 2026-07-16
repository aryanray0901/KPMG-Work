"""
Verifies Recruiting Radar's content and logic: every deadline entry has
required fields and a real source link, the live countdown math is
correct, the resume checker correctly distinguishes a strong resume from
a weak one, and every content module loads without error.

Usage: python tools/verify_content.py
"""
import os
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from content.deadlines import DEADLINES, days_until, upcoming_deadlines, passed_deadlines
from content.stats import REALITY_STATS
from content.networking import NETWORKING_TECHNIQUES, COLD_OUTREACH_TEMPLATES, COFFEE_CHAT_QUESTIONS
from content.resume_rules import RESUME_MISTAKES
from content.nontarget_playbook import NONTARGET_PLAYBOOK
from app import check_resume

failures = []


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" — {detail}" if detail and not condition else ""))
    if not condition:
        failures.append(name)


print("=== Deadline data ===")
check("at least 12 deadlines tracked", len(DEADLINES) >= 12, f"got {len(DEADLINES)}")
missing_fields = [d["program"] for d in DEADLINES
                   if not all(d.get(f) for f in ["firm", "tier", "program", "type", "eligible", "date", "source_name", "source_url"])]
check("every deadline has all required fields", not missing_fields, f"missing in: {missing_fields}")
bad_urls = [d["program"] for d in DEADLINES if not d["source_url"].startswith("http")]
check("every deadline source is a real URL", not bad_urls, f"bad: {bad_urls}")

mbb_firms = {d["firm"] for d in DEADLINES if d["tier"] == "MBB"}
check("all three MBB firms are covered", {"McKinsey", "BCG", "Bain"} <= mbb_firms, f"got {mbb_firms}")
diversity_entries = [d for d in DEADLINES if "Diversity" in d["type"]]
check("diversity/early-access programs are tracked separately", len(diversity_entries) >= 5,
      f"got {len(diversity_entries)}")

print("\n=== Live countdown math ===")
test_date = date(2026, 7, 16)
future_deadline = {"date": date(2026, 8, 11)}
past_deadline = {"date": date(2026, 3, 29)}
check("days_until computes correctly for a future date",
      days_until(future_deadline, today=test_date) == 26, f"got {days_until(future_deadline, today=test_date)}")
check("days_until computes correctly for a past date (negative)",
      days_until(past_deadline, today=test_date) < 0, f"got {days_until(past_deadline, today=test_date)}")
check("upcoming_deadlines excludes passed dates",
      all(d["date"] >= test_date for d in upcoming_deadlines(today=test_date)))
check("passed_deadlines excludes future dates",
      all(d["date"] < test_date for d in passed_deadlines(today=test_date)))

print("\n=== Stats and playbook content ===")
check("at least 6 sourced stats", len(REALITY_STATS) >= 6, f"got {len(REALITY_STATS)}")
bad_stat_sources = [s["label"] for s in REALITY_STATS if not s["source_url"].startswith("http")]
check("every stat has a real source URL", not bad_stat_sources, f"bad: {bad_stat_sources}")

check("at least 2 networking techniques with sources",
      len(NETWORKING_TECHNIQUES) >= 2 and all(t["source_url"].startswith("http") for t in NETWORKING_TECHNIQUES))
check("at least 4 outreach templates", len(COLD_OUTREACH_TEMPLATES) >= 4, f"got {len(COLD_OUTREACH_TEMPLATES)}")
check("at least 5 coffee chat questions", len(COFFEE_CHAT_QUESTIONS) >= 5, f"got {len(COFFEE_CHAT_QUESTIONS)}")
check("at least 5 resume mistakes with bad/good examples",
      len(RESUME_MISTAKES) >= 5 and all(m.get("example_bad") and m.get("example_good") for m in RESUME_MISTAKES))
check("at least 4 non-target playbook items with sources",
      len(NONTARGET_PLAYBOOK) >= 4 and all(p["source_url"].startswith("http") for p in NONTARGET_PLAYBOOK))

print("\n=== Resume checker logic ===")
weak_resume = "We worked on improving the marketing process. Helped organize events. Was responsible for social media."
strong_resume = """Jane Doe. Education: State University, GPA 3.6.
Redesigned the email campaign segmentation strategy, increasing open rates 22% over one quarter.
Built a weekly reporting dashboard used by 4 team members, cutting manual reporting time by 5 hours per week.
Led a cross-functional pricing analysis project with 3 stakeholders, informing a $200K pricing decision.
Leadership: President, Marketing Club, grew membership 40% to 85 active members.
Skills: Excel, SQL, Python, Tableau."""

weak_results = check_resume(weak_resume)
strong_results = check_resume(strong_resume)
weak_flags = sum(1 for r in weak_results if r["status"] == "flag")
strong_flags = sum(1 for r in strong_results if r["status"] == "flag")
check("weak resume gets more flags than a strong one", weak_flags > strong_flags,
      f"weak={weak_flags} flags, strong={strong_flags} flags")
check("weak resume is flagged for weak verbs",
      any(r["check"] == "Verb strength" and r["status"] == "flag" for r in weak_results))
check("strong resume passes the quantification check",
      any(r["check"] == "Quantified impact" and r["status"] == "pass" for r in strong_results))
check("empty resume returns a flag, not a crash", check_resume("")[0]["status"] == "flag")

print("\n" + "=" * 40)
if failures:
    print(f"FAIL: {len(failures)} check(s) failed: {failures}")
    sys.exit(1)
else:
    print("ALL CHECKS PASSED")
