"""
Case Prep
---------
A fully in-browser case interview practice platform for consulting
interviews (McKinsey, BCG, Bain, Deloitte, PwC, EY, KPMG style). Nothing
is downloaded or exported; everything happens on the page. Includes a
case library with progressive exhibit reveal and timers, a framework
reference library, a fit/behavioral question bank, a market sizing drill,
a mental math drill, and firm-by-firm interview format guides.

An optional AI Coach can give feedback on practice responses using the
student's own OpenAI or Anthropic API key (entered per-session, never
stored to disk, sent directly to the provider on each request). Without
a key, a rule-based structured self-assessment is used instead, comparing
the student's response against the case's model approach, this is
useful on its own, not just a fallback.
"""

import os
import re
import json
import random

from flask import Flask, request, render_template, jsonify, session
import requests

from content.cases import CASES, get_case, cases_by_firm
from content.firm_guides import FIRM_GUIDES, get_firm_guide
from content.frameworks import FRAMEWORKS, get_framework
from content.fit_questions import FIT_QUESTIONS
from content.market_sizing import MARKET_SIZING_PROMPTS

app = Flask(__name__)
app.secret_key = "case-prep-local-secret"


@app.after_request
def _no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response


# ---------------------------------------------------------------------------
# AI COACH (optional, bring-your-own-key)
# ---------------------------------------------------------------------------

def call_openai(api_key, system_prompt, user_prompt):
    resp = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={
            "model": "gpt-4o-mini",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "max_tokens": 600,
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def call_anthropic(api_key, system_prompt, user_prompt):
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        },
        json={
            "model": "claude-3-5-sonnet-20241022",
            "max_tokens": 600,
            "system": system_prompt,
            "messages": [{"role": "user", "content": user_prompt}],
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    return data["content"][0]["text"]


def get_ai_feedback(provider, api_key, system_prompt, user_prompt):
    if provider == "openai":
        return call_openai(api_key, system_prompt, user_prompt)
    if provider == "anthropic":
        return call_anthropic(api_key, system_prompt, user_prompt)
    raise ValueError(f"Unknown provider: {provider}")


STOPWORDS = {
    "the", "a", "an", "and", "or", "of", "to", "in", "on", "for", "is", "are",
    "was", "were", "with", "this", "that", "it", "as", "at", "by", "be", "from",
    "than", "rather", "whether", "which", "would", "could", "should", "will",
    "their", "there", "these", "those", "then", "when", "where", "what", "who",
    "into", "onto", "over", "under", "about", "between", "before", "after",
    "not", "but", "also", "each", "such", "more", "most", "some", "any", "all",
    "candidate", "case", "exhibit", "strong", "should", "given", "rather",
    "likely", "probably", "example", "instead", "directly", "explicitly",
    "clearly", "generally", "specific", "specifically", "roughly", "approach",
    "question", "answer", "interviewer", "prompt",
}


def extract_keywords(text, min_len=4, max_words=12):
    words = re.findall(r"[a-zA-Z]+", text.lower())
    counts = {}
    for w in words:
        if len(w) >= min_len and w not in STOPWORDS:
            counts[w] = counts.get(w, 0) + 1
    ranked = sorted(counts.items(), key=lambda x: -x[1])
    return [w for w, _ in ranked[:max_words]]


def rule_based_case_feedback(user_text, case):
    """A structured self-assessment used when no AI key is provided. Not a
    simulation of AI feedback, a genuinely different, honest mechanism:
    keyword coverage against the model approach, plus a fixed checklist."""
    word_count = len(user_text.split())
    model_keywords = extract_keywords(case["model_approach"] + " " + case["model_synthesis"])
    user_lower = user_text.lower()
    covered = [k for k in model_keywords if k in user_lower]
    missing = [k for k in model_keywords if k not in user_lower][:6]

    checklist = [
        f"Length: your response is {word_count} words. Case interview answers are usually spoken, "
        f"but if this reads as too thin to cover a full structure, that's worth noticing now.",
        f"Coverage: your response touches on {len(covered)} of the {len(model_keywords)} key "
        f"terms/concepts that came up in the model approach for this case"
        + (f" (missing, consider whether these are relevant: {', '.join(missing)})." if missing else "."),
        "Structure check: did you state an approach before diving into analysis, rather than "
        "jumping straight to a conclusion?",
        f"Pitfall check for this case type: {case.get('pitfall', 'Compare your structure against the model approach below for common traps.')}",
        "Now compare your written response against the model approach and model synthesis below, "
        "and be honest with yourself about what you'd have missed live, under time pressure.",
    ]
    return "\n\n".join(checklist)


def rule_based_fit_feedback(user_text):
    word_count = len(user_text.split())
    has_star_signals = any(w in user_text.lower() for w in
                            ["situation", "task", "result", "so that", "as a result", "which led to"])
    lines = [
        f"Length: {word_count} words. A strong spoken STAR answer is usually 60-90 seconds, "
        f"written out that's often somewhere in the 120-220 word range, use this as a rough gut check, not a hard rule.",
        "Specificity check: does your answer include a concrete situation with real stakes, "
        "or could it apply to almost any story? Concrete numbers, names of projects, or timeframes are good signs.",
        "Result check: " + ("your answer appears to reference an outcome or result explicitly, good."
                             if has_star_signals else
                             "make sure you explicitly state the outcome/result, not just the actions you took."),
        "Ownership check: re-read your answer and count how many sentences use 'I' versus 'we.' "
        "Interviewers are assessing your individual contribution specifically.",
    ]
    return "\n\n".join(lines)


def rule_based_sizing_feedback(user_text, prompt_data):
    word_count = len(user_text.split())
    has_numbers = bool(re.search(r"\d", user_text))
    has_structure_words = any(w in user_text.lower() for w in
                               ["population", "estimate", "assume", "average", "per", "roughly", "approximately"])
    lines = [
        f"Length: {word_count} words.",
        "Numeric check: " + ("your response includes specific numbers, good, market sizing should "
                              "always end in an actual estimate, not just a described approach."
                              if has_numbers else
                              "your response doesn't appear to include specific numbers yet. "
                              "Market sizing answers need to land on an actual estimate, not just describe the method."),
        "Assumption language check: " + ("your response uses estimation language ('assume', "
                                          "'roughly', 'average'), which is appropriate, precision here would be a red flag."
                                          if has_structure_words else
                                          "make sure you're explicitly stating your assumptions ('assume roughly X%...') rather than presenting numbers without justification."),
        f"Suggested approach for this prompt: {prompt_data['approach']}",
        f"Sanity check range: {prompt_data['range_hint']}",
    ]
    return "\n\n".join(lines)


# ---------------------------------------------------------------------------
# ROUTES
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", case_count=len(CASES), firm_count=len(FIRM_GUIDES),
                            framework_count=len(FRAMEWORKS),
                            fit_count=sum(len(c["questions"]) for c in FIT_QUESTIONS),
                            sizing_count=len(MARKET_SIZING_PROMPTS))


@app.route("/cases", methods=["GET"])
def cases_list():
    firm_filter = request.args.get("firm", "")
    cases = cases_by_firm(firm_filter) if firm_filter else CASES
    return render_template("cases_list.html", cases=cases, firms=FIRM_GUIDES, active_firm=firm_filter)


@app.route("/cases/<slug>", methods=["GET"])
def case_detail(slug):
    case = get_case(slug)
    if not case:
        return render_template("not_found.html"), 404
    firm = get_firm_guide(case["firm_style"])
    return render_template("case_detail.html", case=case, firm=firm)


@app.route("/api/case-feedback", methods=["POST"])
def api_case_feedback():
    data = request.get_json(force=True)
    slug = data.get("slug", "")
    user_text = (data.get("text") or "").strip()
    provider = data.get("provider", "")
    api_key = (data.get("api_key") or "").strip()

    case = get_case(slug)
    if not case:
        return jsonify({"error": "Unknown case."}), 404
    if not user_text:
        return jsonify({"error": "Write your approach and synthesis before requesting feedback."}), 400

    if api_key and provider in ("openai", "anthropic"):
        system_prompt = (
            "You are an experienced case interview coach for top consulting firms "
            "(McKinsey, BCG, Bain, and Big 4 consulting). Give direct, specific, constructive "
            "feedback on a candidate's written case response. Be honest about weaknesses. "
            "Keep feedback to about 200 words, structured as short bullet points."
        )
        user_prompt = (
            f"Case prompt: {case['prompt']}\n\n"
            f"Model approach (for your reference, not the candidate's): {case['model_approach']}\n\n"
            f"Model synthesis (for your reference): {case['model_synthesis']}\n\n"
            f"Candidate's written response:\n{user_text}\n\n"
            f"Give specific feedback on their structure, prioritization, and synthesis."
        )
        try:
            feedback = get_ai_feedback(provider, api_key, system_prompt, user_prompt)
            return jsonify({"feedback": feedback, "mode": "ai"})
        except requests.exceptions.RequestException as e:
            return jsonify({"feedback": rule_based_case_feedback(user_text, case), "mode": "fallback",
                             "note": f"AI request failed ({e}); showing structured self-assessment instead."})

    return jsonify({"feedback": rule_based_case_feedback(user_text, case), "mode": "rule_based"})


@app.route("/frameworks", methods=["GET"])
def frameworks_list():
    return render_template("frameworks_list.html", frameworks=FRAMEWORKS)


@app.route("/frameworks/<slug>", methods=["GET"])
def framework_detail(slug):
    fw = get_framework(slug)
    if not fw:
        return render_template("not_found.html"), 404
    return render_template("framework_detail.html", fw=fw)


@app.route("/fit", methods=["GET"])
def fit_bank():
    return render_template("fit_bank.html", categories=FIT_QUESTIONS)


@app.route("/api/fit-feedback", methods=["POST"])
def api_fit_feedback():
    data = request.get_json(force=True)
    user_text = (data.get("text") or "").strip()
    question = (data.get("question") or "").strip()
    provider = data.get("provider", "")
    api_key = (data.get("api_key") or "").strip()

    if not user_text:
        return jsonify({"error": "Write your answer before requesting feedback."}), 400

    if api_key and provider in ("openai", "anthropic"):
        system_prompt = (
            "You are an experienced consulting interview coach. Give direct, constructive "
            "feedback on a candidate's behavioral/fit interview answer using the STAR method "
            "as your evaluation lens. Keep feedback to about 150 words, structured as short bullet points."
        )
        user_prompt = f"Question: {question}\n\nCandidate's answer:\n{user_text}"
        try:
            feedback = get_ai_feedback(provider, api_key, system_prompt, user_prompt)
            return jsonify({"feedback": feedback, "mode": "ai"})
        except requests.exceptions.RequestException as e:
            return jsonify({"feedback": rule_based_fit_feedback(user_text), "mode": "fallback",
                             "note": f"AI request failed ({e}); showing structured self-assessment instead."})

    return jsonify({"feedback": rule_based_fit_feedback(user_text), "mode": "rule_based"})


@app.route("/market-sizing", methods=["GET"])
def market_sizing():
    idx = request.args.get("i")
    if idx is not None and idx.isdigit() and int(idx) < len(MARKET_SIZING_PROMPTS):
        prompt_idx = int(idx)
    else:
        prompt_idx = random.randrange(len(MARKET_SIZING_PROMPTS))
    return render_template("market_sizing.html", prompt=MARKET_SIZING_PROMPTS[prompt_idx],
                            idx=prompt_idx, total=len(MARKET_SIZING_PROMPTS))


@app.route("/api/sizing-feedback", methods=["POST"])
def api_sizing_feedback():
    data = request.get_json(force=True)
    user_text = (data.get("text") or "").strip()
    idx = data.get("idx", 0)
    provider = data.get("provider", "")
    api_key = (data.get("api_key") or "").strip()

    if not user_text:
        return jsonify({"error": "Write your estimate and reasoning before requesting feedback."}), 400
    if not isinstance(idx, int) or idx < 0 or idx >= len(MARKET_SIZING_PROMPTS):
        idx = 0
    prompt_data = MARKET_SIZING_PROMPTS[idx]

    if api_key and provider in ("openai", "anthropic"):
        system_prompt = (
            "You are an experienced case interview coach. Evaluate a candidate's market "
            "sizing reasoning: is the structure logical, are assumptions stated explicitly and "
            "reasonable, does it land on an actual number. Keep feedback to about 150 words."
        )
        user_prompt = f"Prompt: {prompt_data['prompt']}\n\nCandidate's response:\n{user_text}"
        try:
            feedback = get_ai_feedback(provider, api_key, system_prompt, user_prompt)
            return jsonify({"feedback": feedback, "mode": "ai"})
        except requests.exceptions.RequestException as e:
            return jsonify({"feedback": rule_based_sizing_feedback(user_text, prompt_data), "mode": "fallback",
                             "note": f"AI request failed ({e}); showing structured self-assessment instead."})

    return jsonify({"feedback": rule_based_sizing_feedback(user_text, prompt_data), "mode": "rule_based"})


@app.route("/math-drill", methods=["GET"])
def math_drill():
    return render_template("math_drill.html")


@app.route("/firms", methods=["GET"])
def firms_list():
    return render_template("firms_list.html", firms=FIRM_GUIDES)


@app.route("/firms/<slug>", methods=["GET"])
def firm_detail(slug):
    firm = get_firm_guide(slug)
    if not firm:
        return render_template("not_found.html"), 404
    return render_template("firm_detail.html", firm=firm, cases=cases_by_firm(slug))


def _open_browser():
    import webbrowser
    import time
    import sys as _sys
    time.sleep(1.2)
    url = "http://127.0.0.1:5100"
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
    print("Starting Case Prep at http://127.0.0.1:5100")
    app.run(host="127.0.0.1", port=5100, debug=("--debug" in sys.argv))
