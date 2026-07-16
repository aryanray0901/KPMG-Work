"""Common consulting resume mistakes, sourced from recruiting guides
written by former consultants and resume-review services. Used both as
reference content and as the basis for the automated resume checker's
rules."""

RESUME_MISTAKES = [
    {
        "mistake": "Writing bullets that describe responsibilities instead of results",
        "example_bad": "Managed customer retention initiatives and worked with cross-functional teams",
        "example_good": "Redesigned customer retention workflow across 3 teams, reducing churn 12% over two quarters",
        "source_name": "Road to Offer", "source_url": "https://www.roadtooffer.com/blog/consulting-resume-guide",
    },
    {
        "mistake": "Overusing 'we' instead of 'I', making individual contribution unclear",
        "example_bad": "We increased sales by improving the outreach process",
        "example_good": "I redesigned the outreach process, which the team adopted; sales increased 18% over the following quarter",
        "source_name": "IGotAnOffer", "source_url": "https://igotanoffer.com/blogs/mckinsey-case-interview-blog/consulting-resume",
    },
    {
        "mistake": "Omitting an extracurricular / leadership section entirely",
        "example_bad": "(section missing)",
        "example_good": "A dedicated section prioritizing leadership roles over simple club membership, ordered by impact rather than chronology",
        "source_name": "IGotAnOffer", "source_url": "https://igotanoffer.com/blogs/mckinsey-case-interview-blog/consulting-resume",
    },
    {
        "mistake": "Unquantified impact statements",
        "example_bad": "Helped organize a successful recruiting event",
        "example_good": "Organized 5 recruiting events with 500+ attendees, resulting in 3 new corporate partnerships and $15K in sponsorship revenue",
        "source_name": "Hacking the Case Interview", "source_url": "https://www.hackingthecaseinterview.com/pages/consulting-resume-guide",
    },
    {
        "mistake": "Going over one page",
        "example_bad": "(2+ pages)",
        "example_good": "One page, every line earning its place",
        "source_name": "IGotAnOffer", "source_url": "https://igotanoffer.com/blogs/mckinsey-case-interview-blog/consulting-resume",
    },
    {
        "mistake": "Leaving jargon or acronyms unexplained for a non-specialist reader",
        "example_bad": "Led development of a novel CRISPRi-based screening assay for XYZ pathway modulation",
        "example_good": "Led development of a new lab technique to test how genes affect a specific disease pathway, cutting experiment time by 40%",
        "source_name": "IGotAnOffer", "source_url": "https://igotanoffer.com/blogs/mckinsey-case-interview-blog/consulting-resume",
        "note": "Especially relevant for thesis-based or highly technical work; the resume screener will not take time to look up unfamiliar terms.",
    },
]
