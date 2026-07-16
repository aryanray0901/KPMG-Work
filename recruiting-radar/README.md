# Recruiting Radar

**Everything that happens before a consulting interview: real deadlines, real networking strategy, and a resume check, all backed by sourced, current data.**

Case interview practice, frameworks, fit questions, and firm interview-format guides are already covered by Case Prep elsewhere in this repository. This tool deliberately covers what that one doesn't: the part of the process that eliminates more candidates than the interview itself. Consulting recruiting has a well-documented set of structural traps: deadlines that land months before students expect them, an application portal that functions as a near-automatic filter without a referral, and a resume screen that takes well under a minute. Every fact and figure in this tool was researched and sourced, not invented, for exactly this reason.

![Recruiting Radar homepage](screenshots/homepage.png)

## What's included

- **Deadline Radar**: a database of real application windows for McKinsey, BCG, Bain, and Big 4 firms, including the early diversity and sophomore-year programs that consistently carry the earliest deadlines of the entire cycle and are the ones most students find out about after they've already closed. Countdown status is computed live against the actual current date every time the page loads, not a static snapshot.
- **The Real Numbers**: sourced statistics, acceptance rates, resume-screen elimination rates, GPA data, referral impact, that explain why the rest of the advice here matters. Numbers without the reasoning behind them are just trivia; this page is meant to make the "why" concrete.
- **Networking Playbook**: real, named techniques from people who used them (including "leapfrog networking," from a first-hand account of breaking into consulting from a non-target school), outreach message templates for LinkedIn and email, a coffee chat question bank, and a personal contact tracker.
- **Resume Check**: paste your resume text for an automated check against the most commonly cited consulting resume mistakes (weak verbs, unquantified bullets, "we" instead of "I," missing sections), plus a reference list of real mistakes with bad/good examples.
- **Non-Target & Low-GPA Playbook**: specific, sourced strategies for candidates without an existing firm network or a top-of-range GPA, the two most commonly cited structural disadvantages in this process.
- **My Tracker**: a simple personal tracker for the firms and programs you're applying to.

## What "live" means here

This tool doesn't call any external API, and none of the trackers require an account or a download. "Live" refers to the deadline countdowns: every date is a real, researched date, and the countdown status (days remaining, or "passed") is computed against your computer's actual current date each time the page loads. Open this tool in six months and the countdowns will reflect that, not a snapshot frozen at build time.

## Setup

- **Mac:** double-click `Start on Mac.command`
- **Windows:** double-click `Start on Windows.bat`

No sample files needed, all content is built in.

Manual setup:
```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Then open `http://127.0.0.1:5110`.

## Verification

```bash
python tools/verify_content.py
```

Checks that every deadline entry has a real source URL and required fields, that all three MBB firms and multiple diversity/early-access programs are represented, that the live countdown math is correct for both future and past dates, and that the resume checker correctly distinguishes a genuinely strong resume from a weak one (more flags on the weak example, specific correct flags on specific known issues). All 20 checks currently pass. Also tested with a real browser: loaded the homepage and extracted the actual rendered countdown text to confirm it matches the underlying data, not just that the page returns a 200.

## Design notes

- **Every fact links to where it came from.** Deadlines, statistics, and networking techniques all cite a real source. Recruiting timelines and firm processes change every year, this tool is a starting point for your own verification against official firm pages, not a permanent source of truth.
- **Approximate dates are labeled as approximate.** Several programs' next-cycle deadlines weren't yet confirmed at the time of research; those entries are explicitly marked rather than presented with false precision.
- **The resume checker is pattern matching, not judgment.** It's transparent about exactly what it checks (word count, verb choice, quantification density, a few keyword-based section checks) so its limits are obvious. It will miss things a human reviewer would catch, and occasionally flag something that's actually fine. Treat it as a first pass.
- **Trackers are session-only by design.** Nothing you enter is written to a file or sent anywhere; it lives in your browser session and clears when you clear it or restart.
- All processing takes place locally. No file is transmitted to an external service.
