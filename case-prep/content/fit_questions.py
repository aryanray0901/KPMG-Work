"""Fit / behavioral question bank, organized by category, with STAR guidance,
an illustrative example answer, and links to real preparation resources
written by former or current consulting employees.

The "example answer" for each question is written for this tool, to
illustrate STAR structure. It is not attributed to any real person. The
linked resources, by contrast, are real, publicly available guides;
several are written or reviewed by named former consultants (credited in
the "source" field where the site itself states this). Firms revise their
interview formats over time, so treat linked pages as a starting point for
research, not a guarantee of current format.
"""

FIT_QUESTIONS = [
    {
        "category": "Motivation",
        "resources": [
            {
                "title": "How to Ace Consulting Fit Interviews",
                "source": "CaseCoach (coaching team includes former McKinsey, BCG, and Bain consultants)",
                "url": "https://casecoach.com/b/how-to-ace-consulting-fit-interviews-the-free-casecoach-guide/",
            },
            {
                "title": "31 Consulting Interview Questions (from McKinsey, BCG, etc.)",
                "source": "IGotAnOffer",
                "url": "https://igotanoffer.com/blogs/mckinsey-case-interview-blog/consulting-interview-questions",
            },
        ],
        "questions": [
            {
                "q": "Why consulting?",
                "guidance": "Avoid generic answers about 'variety' or 'smart people.' Anchor on a "
                            "specific reason tied to how you like to work (e.g. structured "
                            "problem-solving across industries) and ideally a concrete experience "
                            "that revealed this, not just an abstract preference.",
                "example_answer": "I realized I wanted consulting during a class project where our "
                                   "team had to help a real small business figure out why a new "
                                   "product line was underperforming. What stuck with me wasn't the "
                                   "answer we found, it was how much I enjoyed the process of turning "
                                   "a vague, anxious 'sales are down' into a structured question we "
                                   "could actually test. I've since sought out more of that: a data "
                                   "analysis role where I got to reframe ambiguous requests into clear "
                                   "hypotheses. Consulting is the first path I've found where that's "
                                   "the whole job, not a side benefit of it.",
            },
            {
                "q": "Why this firm specifically?",
                "guidance": "Research something specific: a practice area, a recent piece of "
                            "published research, a person you spoke with, the firm's approach to a "
                            "topic you care about. 'Great culture and smart people' applies to every "
                            "firm and signals you haven't done the research.",
                "example_answer": "Two things drew me here specifically. First, I read your team's "
                                   "recent published research on mid-market supply chain resilience, "
                                   "and the framing was different from what I'd seen elsewhere, it "
                                   "focused on decision speed under uncertainty rather than just cost "
                                   "optimization, which matches how I think about problems. Second, I "
                                   "spoke with two current consultants here about staffing, and both "
                                   "independently mentioned the same thing: that the firm actively "
                                   "rotates people across industries early on rather than "
                                   "specializing immediately. That's specifically the kind of breadth "
                                   "I'm looking for right now.",
            },
            {
                "q": "Where do you see yourself in five years?",
                "guidance": "Show you understand the firm's career path and have thought about it "
                            "honestly, without over-committing to a specific practice you haven't "
                            "actually experienced yet. It's fine to express genuine uncertainty "
                            "alongside a clear sense of what you want to learn next.",
                "example_answer": "Honestly, I don't think I know yet which practice area I'll end "
                                   "up specializing in, and I'd be suspicious of my own answer if I "
                                   "claimed certainty this early. What I do know is what I want the "
                                   "next few years to build: real fluency across a few industries, "
                                   "comfort leading a workstream rather than just executing one, and "
                                   "enough client-facing experience to know whether I want to go "
                                   "deep in a specialization or stay a generalist longer. I'd rather "
                                   "let the actual project experience here inform that decision than "
                                   "pretend I've already made it.",
            },
        ],
    },
    {
        "category": "Leadership",
        "resources": [
            {
                "title": "Consulting Personal Fit Interviews: The Only Guide You Need",
                "source": "StrategyCase, written by Florian Smeritschnig, former McKinsey Senior Consultant",
                "url": "https://strategycase.com/consulting-personal-fit-interviews-the-only-guide-you-need-to-read/",
            },
            {
                "title": "Consulting Behavioral & Fit Interview Guide",
                "source": "Hacking the Case Interview (coaching credited to Taylor Warfield, former Bain interviewer)",
                "url": "https://www.hackingthecaseinterview.com/pages/consulting-behavioral-fit-interview",
            },
        ],
        "questions": [
            {
                "q": "Tell me about a time you led a team through a difficult situation.",
                "guidance": "Use STAR (Situation, Task, Action, Result). Be specific about your "
                            "individual actions, not just what 'the team' did, interviewers are "
                            "assessing your personal contribution and decision-making, not the "
                            "team's collective success.",
                "example_answer": "Situation: I was leading a five-person volunteer team on a "
                                   "fundraising deadline, and two weeks in, our biggest planned "
                                   "donor pulled out. Task: I needed to close a 40% funding gap in "
                                   "three weeks without losing the team's morale. Action: I called an "
                                   "immediate meeting, was direct about the gap, and asked each "
                                   "person to bring one new lead by the next day rather than trying "
                                   "to solve it myself. I personally took the two hardest leads. "
                                   "Result: we closed the gap with two days to spare, and two team "
                                   "members told me afterward it was the first time they'd felt "
                                   "trusted with real ownership rather than just tasks.",
            },
            {
                "q": "Describe a time you had to persuade someone who disagreed with you.",
                "guidance": "Show you understood their perspective genuinely before persuading, not "
                            "just that you 'won.' The best answers show real listening and often a "
                            "modification of your original position based on their input.",
                "example_answer": "A teammate on a class project strongly disagreed with my "
                                   "proposed structure for our presentation, he thought it buried "
                                   "the recommendation too late. My first instinct was to defend it, "
                                   "but I asked him to walk me through why instead. He was right that "
                                   "our audience (a busy guest panel) would likely tune out before we "
                                   "got to the point. I combined our approaches: kept my analysis "
                                   "structure but moved the recommendation to the first slide. The "
                                   "panel specifically complimented how clear and upfront the "
                                   "recommendation was.",
            },
            {
                "q": "Tell me about a time you took initiative without being asked.",
                "guidance": "Pick an example with real stakes and a measurable outcome, not a "
                            "trivial task. Be clear about why it wasn't simply part of your assigned "
                            "role.",
                "example_answer": "In my internship, I noticed our weekly reporting process "
                                   "involved three people manually re-entering the same data into "
                                   "different formats for different stakeholders, nobody had asked me "
                                   "to look at this, it wasn't part of my role. I spent a weekend "
                                   "building a simple shared template that generated all three "
                                   "versions from one input. It cut the weekly process from about "
                                   "three hours to twenty minutes, and my manager adopted it as the "
                                   "team standard.",
            },
        ],
    },
    {
        "category": "Teamwork",
        "resources": [
            {
                "title": "Consulting Behavioral Interview Questions (STAR method walkthrough)",
                "source": "RocketBlocks",
                "url": "https://www.rocketblocks.me/blog/consulting-behavioral-interview-questions.php",
            },
            {
                "title": "BCG Behavioral Questions: Complete Guide",
                "source": "Hacking the Case Interview",
                "url": "https://www.hackingthecaseinterview.com/pages/bcg-behavioral-questions",
            },
        ],
        "questions": [
            {
                "q": "Tell me about a time you worked with a difficult team member.",
                "guidance": "Avoid framing it as 'they were the problem.' Strong answers show "
                            "empathy for why the person acted as they did and a concrete action you "
                            "took to improve the working relationship, not just endurance.",
                "example_answer": "A groupmate was consistently missing our check-ins and turning "
                                   "in work late, and the rest of the team was getting frustrated. "
                                   "Instead of escalating immediately, I asked to grab coffee with "
                                   "him one-on-one. It turned out he was dealing with a heavy course "
                                   "load in another class he hadn't mentioned. We restructured his "
                                   "part of the workload to fit around his schedule and set up a "
                                   "shorter daily check-in instead of our longer weekly one. His "
                                   "output became reliable again, and the team never had to have an "
                                   "awkward confrontation about it.",
            },
            {
                "q": "Describe a time your team disagreed on the right approach.",
                "guidance": "Focus on how the disagreement was resolved constructively, ideally "
                            "with a clear decision process, rather than who was 'right' in the end.",
                "example_answer": "Our team split on whether to prioritize a broader survey or a "
                                   "few in-depth interviews for a research project, and after twenty "
                                   "minutes we were just repeating our positions. I suggested we "
                                   "write down what each approach was actually optimizing for, before "
                                   "arguing for either. That reframing showed we weren't actually "
                                   "disagreeing on values, we agreed depth mattered more given our "
                                   "timeline, we'd just been arguing past each other. We went with "
                                   "interviews, and having the criteria written down made the next "
                                   "three team decisions faster too.",
            },
            {
                "q": "Tell me about a time you had to rely on someone else to get something done.",
                "guidance": "Shows you can delegate and trust others, a real concern for candidates "
                            "who read as overly independent or controlling in team settings.",
                "example_answer": "I was managing a tight deadline on a project with a data "
                                   "component I wasn't strong in, and my instinct was to muscle "
                                   "through it myself to avoid looking like I couldn't handle my "
                                   "part. Instead I asked a teammate with a stats background to take "
                                   "that piece, gave her the context and deadline, and checked in "
                                   "twice rather than hovering. Her analysis was better than what I "
                                   "would have produced, and it freed me to focus on the parts I was "
                                   "actually strongest at, which is the trade I should have made from "
                                   "the start.",
            },
        ],
    },
    {
        "category": "Failure & Resilience",
        "resources": [
            {
                "title": "Consulting Behavioral & Fit Interview Guide",
                "source": "Hacking the Case Interview",
                "url": "https://www.hackingthecaseinterview.com/pages/consulting-behavioral-fit-interview",
            },
            {
                "title": "Deloitte Behavioral Interview: Questions and Answers",
                "source": "Hacking the Case Interview (Deloitte-specific)",
                "url": "https://www.hackingthecaseinterview.com/pages/deloitte-behavioral-interview",
            },
        ],
        "questions": [
            {
                "q": "Tell me about a time you failed at something important.",
                "guidance": "Pick a genuine failure with real consequences, not a humble-brag "
                            "disguised as a failure. Be specific about what you learned and, ideally, "
                            "how you've since applied that lesson.",
                "example_answer": "I underestimated how long a client-facing deliverable would take "
                                   "during my internship and told my manager it would be ready a day "
                                   "before I actually finished it, purely because I didn't want to "
                                   "seem slow. It went out late, and worse, my manager was blindsided "
                                   "in a client call. I learned that giving an honest, slightly "
                                   "conservative estimate upfront is always better than an optimistic "
                                   "one I might miss. Since then I build in buffer time explicitly "
                                   "and flag risk early rather than staying quiet and hoping I catch "
                                   "up.",
            },
            {
                "q": "Describe a time you received difficult feedback.",
                "guidance": "Show you took it seriously rather than being defensive, and describe a "
                            "specific behavior change that resulted, not just 'I took it well.'",
                "example_answer": "A professor told me directly that my presentations were "
                                   "technically strong but hard to follow because I front-loaded "
                                   "methodology before getting to the point. My first reaction was "
                                   "defensive, I'd worked hard on that methodology. But I sat with it "
                                   "and realized she was right: I was optimizing for showing my work "
                                   "rather than for the audience's understanding. In my next three "
                                   "presentations I deliberately opened with the conclusion first. "
                                   "It's now just how I structure everything, including this answer.",
            },
            {
                "q": "Tell me about a time you had to adapt quickly to a major change.",
                "guidance": "Consulting involves frequent context-switching between projects and "
                            "teams; a good answer shows comfort with ambiguity and a fast learning "
                            "curve.",
                "example_answer": "Two weeks before a major deliverable, the client sponsor for our "
                                   "project changed and the new sponsor had different priorities than "
                                   "what we'd been building toward. Rather than pushing back on the "
                                   "change, I set up a call within 24 hours to understand her actual "
                                   "priorities, then spent a weekend restructuring our existing "
                                   "analysis around the new framing rather than starting over. We hit "
                                   "the original deadline with a deliverable the new sponsor actually "
                                   "wanted, which wouldn't have happened if I'd waited for more "
                                   "clarity before acting.",
            },
        ],
    },
    {
        "category": "Impact & Achievement",
        "resources": [
            {
                "title": "Consulting Personal Fit Interviews: The Only Guide You Need",
                "source": "StrategyCase, written by Florian Smeritschnig, former McKinsey Senior Consultant",
                "url": "https://strategycase.com/consulting-personal-fit-interviews-the-only-guide-you-need-to-read/",
            },
            {
                "title": "Interview questions at PwC, Deloitte, KPMG and EY: the definitive list",
                "source": "eFinancialCareers, compiled from real candidate and current-employee interview reports",
                "url": "https://www.efinancialcareers.com/news/finance/interview-questions-at-pwc-deloitte-kpmg-and-ey-the-definitive-list",
            },
        ],
        "questions": [
            {
                "q": "Tell me about your proudest achievement.",
                "guidance": "Quantify the impact wherever possible. Make sure your personal role is "
                            "unambiguous if this was a team effort.",
                "example_answer": "I redesigned the volunteer scheduling system for a nonprofit I "
                                   "worked with, which previously relied on a group chat where shifts "
                                   "regularly went unfilled. I built a simple sign-up system and "
                                   "personally recruited and trained the three coordinators who now "
                                   "run it. Unfilled shifts dropped from roughly 30% to under 5% "
                                   "within two months, and the system is still in use two years "
                                   "later, run entirely by people I'm not in contact with anymore, "
                                   "which is the part I'm proudest of.",
            },
            {
                "q": "Describe a time you solved a problem creatively.",
                "guidance": "Emphasize the reasoning process that led to the creative solution, not "
                            "just the solution itself, interviewers want to see how you think.",
                "example_answer": "Our student org had almost no budget for a recruiting event but "
                                   "needed to compete with better-funded clubs for attention. Instead "
                                   "of trying to match their marketing spend, I reframed the problem: "
                                   "what could we offer that money couldn't easily buy? We got three "
                                   "alumni in roles students specifically wanted to commit to a "
                                   "small-group Q&A, something the bigger, better-funded clubs hadn't "
                                   "thought to organize. It became our best-attended event that year, "
                                   "at close to zero cost.",
            },
            {
                "q": "Tell me about a time you had to make a decision with incomplete information.",
                "guidance": "Show a structured approach to the decision (what you knew, what you "
                            "assumed, why) rather than just 'I went with my gut.'",
                "example_answer": "I had to recommend whether our team should pivot a project's "
                                   "direction with only partial user feedback data, the full survey "
                                   "results wouldn't be ready for two more weeks and we had a "
                                   "deadline. I laid out what we knew, what we were assuming, and "
                                   "what the cost of being wrong would be in each direction. Since "
                                   "the cost of a wrong pivot was recoverable but the cost of missing "
                                   "the deadline wasn't, I recommended we proceed with the partial "
                                   "data and built in a checkpoint to course-correct once the full "
                                   "results came in. That structure, not just the decision itself, is "
                                   "what I'd want a reader to take away from this.",
            },
        ],
    },
]
