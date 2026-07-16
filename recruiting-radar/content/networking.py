"""Networking playbook content: real named techniques from people who used
them, message templates, and coffee chat question banks."""

NETWORKING_TECHNIQUES = [
    {
        "name": "Leapfrog networking",
        "source_name": "a non-target-school EY Advisory hire, writing about their own recruiting process",
        "source_url": "https://medium.com/@mymiddleground/q-a-breaking-into-consulting-from-a-non-target-school-2752afff704b",
        "description": "Search LinkedIn for every 2nd-degree connection you have at your target firm, "
                        "then filter by your school, location, or shared background. Send a "
                        "customized connection request that names something specific you have in "
                        "common, hometown, school, a club, volunteering, not a generic 'I'd love to "
                        "connect.' The person who coined this term used it to eventually get to know "
                        "an entire team at EY Advisory well enough that multiple people vouched for "
                        "them to the internship recruiter.",
    },
    {
        "name": "Volume over individual outcomes",
        "source_name": "a non-target-school candidate describing their own MBB recruiting process",
        "source_url": "https://austinschlessinger.substack.com/p/how-to-get-an-interview-with-a-bbm-firm-if-you-are-from-a-non-target-school-a9d41dd740e7",
        "description": "One widely-cited first-hand account describes roughly 100 phone calls and "
                        "1,000 emails before receiving a first-round interview invite. The number "
                        "itself isn't a target to hit, response rates vary hugely, but the underlying "
                        "point is real: for non-target candidates, a handful of networking attempts "
                        "is not yet 'networking,' it's a start. Plan for a volume of outreach, not a "
                        "handful of messages.",
    },
]

COLD_OUTREACH_TEMPLATES = [
    {
        "channel": "LinkedIn connection request (short, character-limited)",
        "template": "Hi {name}, I'm a {year} at {school} interested in consulting, and I saw you "
                    "work in {practice_area} at {firm}. Would love to connect and learn about your "
                    "path there.",
        "note": "Keep it under 300 characters. Naming their specific practice area or a shared "
                "background (school, hometown, club) performs better than a generic request.",
    },
    {
        "channel": "Follow-up coffee chat request (after connecting)",
        "template": "Thanks for connecting, {name}! I'm exploring consulting and would really value "
                    "15-20 minutes to hear about your experience at {firm}, particularly {specific "
                    "topic, e.g. 'the transition from your industry background into consulting'}. "
                    "Would you have any time in the next couple of weeks? Happy to work around your "
                    "schedule.",
        "note": "Naming a specific topic (not just 'learn about the firm') signals you've done "
                "homework and makes it easy for them to say yes quickly.",
    },
    {
        "channel": "Cold email (no LinkedIn connection yet)",
        "template": "Subject: {school} student interested in {practice area} at {firm}\n\nHi "
                    "{name},\n\nI'm a {year} at {school} studying {major}, and I came across your "
                    "profile while researching {firm}'s {practice area} work. I'm early in exploring "
                    "a path into consulting and would appreciate 15 minutes of your time to hear "
                    "about your experience, no ask beyond that. Let me know if you're open to a "
                    "quick call in the next few weeks.\n\nThank you for considering it,\n{your name}",
        "note": "Explicitly stating 'no ask beyond that' up front reduces the friction of responding, "
                "people are more likely to say yes when they know exactly what they're agreeing to.",
    },
    {
        "channel": "Follow-up thank-you (after the chat)",
        "template": "Thank you again for taking the time to talk, {name}. Your point about {specific "
                    "thing they said} was really helpful, it's changed how I'm thinking about "
                    "{topic}. I'll keep you posted as I move through the process, and please let me "
                    "know if there's ever anything I can help with on your end.",
        "note": "Reference something specific they said. A generic 'thanks for your time' is easy to "
                "forget; a specific callback shows you were actually listening.",
    },
]

COFFEE_CHAT_QUESTIONS = [
    "What does a typical week actually look like for you right now, not the recruiting-brochure version?",
    "What surprised you most in your first few months that you wish someone had told you beforehand?",
    "How would you describe the difference in culture between this firm and others you considered?",
    "What kind of projects or industries have you gotten to work on so far?",
    "What separates the people who get staffed on the projects they want from those who don't?",
    "What would you look for in a candidate's background if you were reviewing resumes today?",
    "Is there anything about the recruiting process itself I should know that isn't obvious from the firm's website?",
]

RESPONSE_RATE_GUIDANCE = (
    "Response rates to cold outreach vary enormously by firm, season, and how personalized the "
    "message is, there is no single reliable number. What the first-hand accounts above have in "
    "common is a volume mindset: treat a handful of unanswered messages as normal and expected, not "
    "as a signal to stop. Personalizing each message (name-dropping something specific) consistently "
    "gets cited as the biggest lever for response rate, far more than the channel (email vs. LinkedIn) "
    "you use."
)
