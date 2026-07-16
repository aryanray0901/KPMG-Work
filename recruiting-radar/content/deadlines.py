"""
Real recruiting deadline data, compiled from firm career pages and
university career center postings during research for this tool (dates
below reflect the 2026 recruiting cycle, for 2027 internships and
Dec 2026-Aug 2027 full-time start dates, as published by the sources
listed). Recruiting deadlines shift every year and firms sometimes revise
them mid-cycle, so every entry links back to a source and this data
should be treated as a starting point for your own verification, not a
guarantee. Status (passed/upcoming) is computed live against today's
actual date each time the page loads, not hardcoded, so this stays
accurate regardless of when this tool is actually run.
"""

from datetime import date

DEADLINES = [
    # --- McKinsey ---
    {
        "firm": "McKinsey", "tier": "MBB",
        "program": "Business Analyst Internship (Class of 2027)",
        "type": "Internship", "eligible": "Rising seniors",
        "date": date(2026, 3, 29), "certainty": "confirmed",
        "source_name": "Tufts University Career Center", "source_url": "https://careers.tufts.edu/blog/2026/03/06/mbb-2026-application-deadlines-mckinsey-bain-bcg/",
    },
    {
        "firm": "McKinsey", "tier": "MBB",
        "program": "Sophomore Summer Business Analyst (SSBA) — diversity program, Class of 2028",
        "type": "Diversity / Sophomore", "eligible": "Sophomores (Black, Hispanic/Latinx, Indigenous)",
        "date": date(2026, 6, 10), "certainty": "confirmed",
        "source_name": "Villanova University Career Connections", "source_url": "https://connections.villanova.edu/blog/2025/06/03/2026-opportunity-mbb-application-deadlines-mckinsey-bain-bcg/",
        "note": "Applications typically go live around May 1. This is consistently the earliest MBB deadline of the entire cycle, most students don't hear about it until it's already closed.",
    },
    {
        "firm": "McKinsey", "tier": "MBB",
        "program": "Business Analyst full-time (Dec 2026-Aug 2027 grads)",
        "type": "Full-Time", "eligible": "Seniors / final-year grad students",
        "date": date(2026, 8, 11), "certainty": "confirmed",
        "source_name": "RoadToOffer", "source_url": "https://www.roadtooffer.com/blog/mckinsey-deadline",
        "note": "Applications opened July 1, 2026. Rolling admissions: interview slots commonly fill 4-6 weeks before the stated deadline, so applying in week one is a real advantage, not just a formality.",
    },
    {
        "firm": "McKinsey", "tier": "MBB",
        "program": "Insight (advanced degree diversity program)",
        "type": "Diversity / Advanced Degree", "eligible": "Advanced degree candidates",
        "date": date(2026, 3, 1), "certainty": "approximate",
        "source_name": "Hacking the Case Interview", "source_url": "https://www.hackingthecaseinterview.com/pages/consulting-recruiting-timeline",
        "note": "Accepted Insight participants can move onto an early interview path, sometimes weeks ahead of the general applicant pool.",
    },

    # --- BCG ---
    {
        "firm": "BCG", "tier": "MBB",
        "program": "Summer Associate Intern (Class of 2027)",
        "type": "Internship", "eligible": "Rising seniors",
        "date": date(2026, 4, 15), "certainty": "approximate",
        "source_name": "Columbia Career Education", "source_url": "https://www.careereducation.columbia.edu/news/what-you-need-know-consulting-recruiting-summer-2026",
    },
    {
        "firm": "BCG", "tier": "MBB",
        "program": "Associate full-time (Dec 2026-Aug 2027 grads)",
        "type": "Full-Time", "eligible": "Seniors / final-year grad students",
        "date": date(2026, 7, 7), "certainty": "confirmed",
        "source_name": "RoadToOffer", "source_url": "https://www.roadtooffer.com/blog/consulting-application-deadlines-2026",
        "note": "This is consistently the earliest full-time MBB deadline. A McKinsey-only plan can quietly cost you a BCG application if you're not tracking both.",
    },
    {
        "firm": "BCG", "tier": "MBB",
        "program": "Growing Future Leaders — diversity program",
        "type": "Diversity / Sophomore", "eligible": "Sophomores (underrepresented groups)",
        "date": date(2026, 9, 24), "certainty": "confirmed",
        "source_name": "Villanova University Career Connections", "source_url": "https://connections.villanova.edu/blog/2025/06/03/2026-opportunity-mbb-application-deadlines-mckinsey-bain-bcg/",
    },
    {
        "firm": "BCG", "tier": "MBB",
        "program": "Bridge to Consulting — freshman program",
        "type": "Diversity / Freshman", "eligible": "Freshmen",
        "date": date(2026, 10, 1), "certainty": "approximate",
        "source_name": "Hacking the Case Interview", "source_url": "https://www.hackingthecaseinterview.com/pages/bcg-diversity-programs",
        "note": "Shorter and less competitive than Growing Future Leaders. A good entry point if you're not ready for a full internship application yet.",
    },

    # --- Bain ---
    {
        "firm": "Bain", "tier": "MBB",
        "program": "Associate Consultant Intern — Window 1",
        "type": "Internship", "eligible": "Rising seniors",
        "date": date(2026, 3, 29), "certainty": "confirmed",
        "source_name": "Tufts University Career Center", "source_url": "https://careers.tufts.edu/blog/2026/03/06/mbb-2026-application-deadlines-mckinsey-bain-bcg/",
    },
    {
        "firm": "Bain", "tier": "MBB",
        "program": "Associate Consultant Intern — Window 2",
        "type": "Internship", "eligible": "Rising seniors",
        "date": date(2026, 8, 31), "certainty": "confirmed",
        "source_name": "Tufts University Career Center", "source_url": "https://careers.tufts.edu/blog/2026/03/06/mbb-2026-application-deadlines-mckinsey-bain-bcg/",
        "note": "Bain explicitly offers two windows so candidates who want more prep time aren't forced into the earliest deadline. There's no rule against applying to a later window rather than rushing the first.",
    },
    {
        "firm": "Bain", "tier": "MBB",
        "program": "Associate Consultant full-time (Dec 2026-Aug 2027 grads)",
        "type": "Full-Time", "eligible": "Seniors / final-year grad students",
        "date": date(2026, 7, 19), "certainty": "confirmed",
        "source_name": "Yale Office of Career Strategy", "source_url": "https://ocs.yale.edu/blog/2026/03/06/mckinsey-bain-and-bcg-application-deadlines-for-2026-summer-and-full-time-roles/",
    },
    {
        "firm": "Bain", "tier": "MBB",
        "program": "Bain CREW — diversity program",
        "type": "Diversity / Sophomore", "eligible": "Sophomores (underrepresented groups)",
        "date": date(2026, 11, 15), "certainty": "approximate",
        "source_name": "Hacking the Case Interview", "source_url": "https://www.hackingthecaseinterview.com/pages/consulting-diversity-programs",
    },
    {
        "firm": "Bain", "tier": "MBB",
        "program": "Bain BASE — freshman/sophomore program",
        "type": "Diversity / Freshman", "eligible": "Freshmen and sophomores",
        "date": date(2026, 10, 15), "certainty": "approximate",
        "source_name": "RoadToOffer", "source_url": "https://www.roadtooffer.com/blog/consulting-application-deadlines-2026",
        "note": "Runs on its own earlier timeline separate from standard recruiting; track it independently rather than assuming it follows the main Bain calendar.",
    },

    # --- Big 4 (typically later, rolling, and less rigid than MBB) ---
    {
        "firm": "Deloitte", "tier": "Big 4",
        "program": "Consulting Internship (general)",
        "type": "Internship", "eligible": "Rising seniors",
        "date": date(2026, 9, 15), "certainty": "approximate",
        "source_name": "eFinancialCareers", "source_url": "https://www.efinancialcareers.com/news/finance/interview-questions-at-pwc-deloitte-kpmg-and-ey-the-definitive-list",
        "note": "Big 4 firms generally recruit later and more on a rolling, school-by-school basis than MBB. Check your specific school's career portal rather than relying on a single national date.",
    },
    {
        "firm": "Deloitte", "tier": "Big 4",
        "program": "Women's Leadership Launch — diversity program",
        "type": "Diversity / Sophomore", "eligible": "Sophomores who identify as women",
        "date": date(2026, 8, 1), "certainty": "approximate",
        "source_name": "Leland", "source_url": "https://www.joinleland.com/library/a/the-ultimate-timeline-for-full-time-consulting-recruiting",
        "note": "Can open before the fall semester even starts, easy to miss if you're not watching for it over the summer.",
    },
    {
        "firm": "PwC / EY / KPMG", "tier": "Big 4",
        "program": "Consulting Internship (general)",
        "type": "Internship", "eligible": "Rising seniors",
        "date": date(2026, 9, 30), "certainty": "approximate",
        "source_name": "eFinancialCareers", "source_url": "https://www.efinancialcareers.com/news/finance/interview-questions-at-pwc-deloitte-kpmg-and-ey-the-definitive-list",
        "note": "Big 4 timing varies significantly by office and service line. Treat this as a rough planning marker, not a firm date, verify on your target office's page.",
    },
]


def deadlines_sorted():
    return sorted(DEADLINES, key=lambda d: d["date"])


def upcoming_deadlines(today=None):
    today = today or date.today()
    return [d for d in deadlines_sorted() if d["date"] >= today]


def passed_deadlines(today=None):
    today = today or date.today()
    return [d for d in deadlines_sorted() if d["date"] < today]


def days_until(d, today=None):
    today = today or date.today()
    return (d["date"] - today).days
