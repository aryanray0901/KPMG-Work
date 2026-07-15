"""Firm-by-firm interview format guides. Content is a general, public-knowledge
description of how each firm's interview process is commonly structured, not
insider or confidential information."""

FIRM_GUIDES = [
    {
        "slug": "mckinsey",
        "name": "McKinsey & Company",
        "format": "Interviewer-led case + Personal Experience Interview (PEI)",
        "summary": "McKinsey cases are typically interviewer-led, meaning the interviewer "
                    "steers which part of the problem you address next, rather than you "
                    "running the entire structure end to end. Expect a distinct Personal "
                    "Experience Interview (PEI) in most rounds, a deep-dive behavioral "
                    "conversation on one or two stories, often probing for personal impact, "
                    "entrepreneurial drive, and leadership through difficulty.",
        "what_to_expect": [
            "The interviewer often jumps between sections of the case rather than letting you drive the whole structure.",
            "Math tends to be precise and interviewer-verified; show your work clearly.",
            "The PEI digs deep into one or two stories rather than covering many topics briefly, expect multiple follow-up 'why' questions on the same story.",
            "Many offices now use McKinsey's digital assessment (a gamified problem-solving test) earlier in the process.",
        ],
        "tips": [
            "Practice being interrupted mid-structure and picking the conversation back up cleanly.",
            "For the PEI, prepare stories with a clear personal (not team-only) impact and be ready to go deep on details.",
        ],
    },
    {
        "slug": "bcg",
        "name": "Boston Consulting Group (BCG)",
        "format": "Interviewee-led, hypothesis-driven case",
        "summary": "BCG cases tend to be interviewee-led: you propose the structure and "
                    "drive the conversation, with the interviewer nudging rather than "
                    "steering. BCG places heavy emphasis on hypothesis-driven thinking, "
                    "state an initial hypothesis early and refine it as data comes in, "
                    "rather than exhaustively working through a generic framework.",
        "what_to_expect": [
            "You'll typically be expected to propose the next analytical step yourself.",
            "Interviewers reward a clear point of view stated early, even if it changes as you learn more.",
            "Some BCG interviews include a short written or digital case component (the 'BCG Online Case' / chatbot-style case) before live interviews.",
        ],
        "tips": [
            "Practice stating a hypothesis in your first two minutes, not just a framework.",
            "Be comfortable saying 'I'd guess X because Y, let's test that' rather than only listing categories to explore.",
        ],
    },
    {
        "slug": "bain",
        "name": "Bain & Company",
        "format": "Interviewee-led case, often data-heavy, plus a written case round",
        "summary": "Bain cases are generally interviewee-led and tend to be quite "
                    "data- and chart-heavy, expect to interpret exhibits quickly and "
                    "extract the 'so what.' Many Bain processes include a written case "
                    "round (a packet of information you analyze over a set time, then "
                    "present recommendations), in addition to live case interviews.",
        "what_to_expect": [
            "Multiple exhibits per case is common; practice reading a chart in under a minute and stating the key takeaway.",
            "The written case round rewards clear prioritization since you can't cover everything in the packet.",
            "Bain interviewers often ask you to synthesize a recommendation to a hypothetical CEO in under a minute at the end.",
        ],
        "tips": [
            "Practice a strict 60-second 'elevator synthesis' at the end of every case you rehearse.",
            "For chart-heavy practice, time yourself reading and stating the takeaway from unfamiliar exhibits.",
        ],
    },
    {
        "slug": "deloitte",
        "name": "Deloitte",
        "format": "Structured case, sometimes group case, plus behavioral round",
        "summary": "Deloitte's process (particularly Strategy & Analytics / Consulting) "
                    "commonly includes a structured, sometimes more straightforward "
                    "profitability- or market-entry-style case, alongside a solid behavioral "
                    "component. Group case exercises appear in some regions and service lines, "
                    "where you solve a case alongside other candidates and are assessed on "
                    "collaboration as well as content.",
        "what_to_expect": [
            "Cases are often somewhat more guided than MBB, useful for candidates newer to case interviews.",
            "Group case rounds assess how you build on others' ideas, not just your own analysis.",
            "Behavioral questions frequently probe client service orientation and teamwork specifically.",
        ],
        "tips": [
            "If practicing for a group round, rehearse actively building on a partner's idea out loud, not just waiting for your turn.",
        ],
    },
    {
        "slug": "pwc",
        "name": "PwC",
        "format": "Structured case interview plus strengths-based behavioral interview",
        "summary": "PwC is known for a 'strengths-based' interview philosophy alongside "
                    "traditional case interviews, questions often explore what energizes "
                    "you and where you do your best work, not only what you're good at. "
                    "Case rounds tend to be structured and business-problem-oriented, "
                    "commonly profitability, market entry, or operations.",
        "what_to_expect": [
            "Expect questions like 'what do you enjoy about problem-solving work' alongside standard behavioral prompts.",
            "Cases are typically clearly scoped with a defined business question up front.",
        ],
        "tips": [
            "Prepare honest, specific answers about what kind of work energizes you, generic answers are easy to spot.",
        ],
    },
    {
        "slug": "ey",
        "name": "EY (Ernst & Young)",
        "format": "Structured case plus competency-based behavioral interview",
        "summary": "EY-Parthenon and EY Consulting interviews typically pair a structured "
                    "business case (profitability, market entry, growth strategy) with a "
                    "competency-based behavioral round assessing teamwork, client focus, "
                    "and adaptability against defined competencies.",
        "what_to_expect": [
            "Behavioral questions are often mapped to explicit, named competencies, answers that clearly demonstrate the named competency score well.",
            "EY-Parthenon (the strategy arm) cases tend to run closer to MBB-style difficulty than broader EY Consulting cases.",
        ],
        "tips": [
            "Research the specific competency framework for the practice you're applying to if available, and structure STAR answers to hit it explicitly.",
        ],
    },
    {
        "slug": "kpmg",
        "name": "KPMG",
        "format": "Structured case interview plus behavioral/motivational interview",
        "summary": "KPMG Advisory interviews commonly include a structured business case "
                    "(often profitability, market entry, or operational improvement) and a "
                    "behavioral round covering motivation, teamwork, and client service "
                    "examples. Some practices use a short written exercise or numerical "
                    "reasoning test earlier in the process.",
        "what_to_expect": [
            "Cases are generally clearly scoped and moderately paced relative to MBB.",
            "Motivational questions ('why KPMG,' 'why this practice') are asked directly and specifically, generic answers stand out negatively.",
        ],
        "tips": [
            "Have a specific, researched answer for 'why this firm' that references something beyond brand reputation.",
        ],
    },
]


def get_firm_guide(slug):
    for f in FIRM_GUIDES:
        if f["slug"] == slug:
            return f
    return None
