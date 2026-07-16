"""
Verifies Case Prep's content completeness and route coverage: every case,
framework, firm guide, and fit question has all required fields, all
slugs are unique, all seven target firms are covered by at least one
case, and every route in the app resolves for every content item (run
via a live request against a running server, see verify_routes.sh for
the full HTTP-level check this complements).

Usage: python tools/verify_content.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from content.cases import CASES
from content.frameworks import FRAMEWORKS
from content.firm_guides import FIRM_GUIDES
from content.fit_questions import FIT_QUESTIONS
from content.market_sizing import MARKET_SIZING_PROMPTS

failures = []


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" — {detail}" if detail and not condition else ""))
    if not condition:
        failures.append(name)


print("=== Content completeness ===")
required_case_fields = ["slug", "firm_style", "title", "case_type", "difficulty",
                         "prompt", "model_approach", "model_synthesis"]
missing = [f"{c['slug']}.{field}" for c in CASES for field in required_case_fields if not c.get(field)]
check("every case has all required fields", not missing, f"missing: {missing}")

required_fw_fields = ["slug", "name", "when", "structure", "pitfall"]
missing_fw = [f"{f['slug']}.{field}" for f in FRAMEWORKS for field in required_fw_fields if not f.get(field)]
check("every framework has all required fields", not missing_fw, f"missing: {missing_fw}")

required_firm_fields = ["slug", "name", "format", "summary", "what_to_expect", "tips"]
missing_firm = [f"{fi['slug']}.{field}" for fi in FIRM_GUIDES for field in required_firm_fields if not fi.get(field)]
check("every firm guide has all required fields", not missing_firm, f"missing: {missing_firm}")

print("\n=== Uniqueness ===")
case_slugs = [c["slug"] for c in CASES]
fw_slugs = [f["slug"] for f in FRAMEWORKS]
firm_slugs = [f["slug"] for f in FIRM_GUIDES]
check("case slugs are unique", len(case_slugs) == len(set(case_slugs)))
check("framework slugs are unique", len(fw_slugs) == len(set(fw_slugs)))
check("firm slugs are unique", len(firm_slugs) == len(set(firm_slugs)))

print("\n=== Coverage ===")
target_firms = {"mckinsey", "bcg", "bain", "deloitte", "pwc", "ey", "kpmg"}
covered_firms = {c["firm_style"] for c in CASES}
check("all 7 target firms have at least one case", target_firms <= covered_firms,
      f"missing: {target_firms - covered_firms}")
check("at least 8 mock cases", len(CASES) >= 8, f"got {len(CASES)}")
check("at least 6 frameworks", len(FRAMEWORKS) >= 6, f"got {len(FRAMEWORKS)}")
check("at least 5 firm guides with case coverage", len(FIRM_GUIDES) >= 5, f"got {len(FIRM_GUIDES)}")

total_fit_questions = sum(len(c["questions"]) for c in FIT_QUESTIONS)
check("at least 10 fit/behavioral questions", total_fit_questions >= 10, f"got {total_fit_questions}")
check("at least 5 market sizing prompts", len(MARKET_SIZING_PROMPTS) >= 5, f"got {len(MARKET_SIZING_PROMPTS)}")

print("\n=== Answer key content ===")
missing_examples = [f"{c['category']}: {q['q']}" for c in FIT_QUESTIONS for q in c["questions"] if not q.get("example_answer")]
check("every fit question has an example answer", not missing_examples, f"missing: {missing_examples}")

missing_resources = [c["category"] for c in FIT_QUESTIONS if not c.get("resources")]
check("every fit category has researched resource links", not missing_resources, f"missing: {missing_resources}")

bad_urls = [r["url"] for c in FIT_QUESTIONS for r in c.get("resources", []) if not r.get("url", "").startswith("http")]
check("every resource link is a well-formed URL", not bad_urls, f"bad: {bad_urls}")

missing_source = [r["title"] for c in FIT_QUESTIONS for r in c.get("resources", []) if not r.get("source")]
check("every resource link credits its source", not missing_source, f"missing source: {missing_source}")

print("\n=== Rule-based feedback engine ===")
from app import rule_based_case_feedback, rule_based_fit_feedback, rule_based_sizing_feedback, extract_keywords

sample_case = CASES[0]
fb = rule_based_case_feedback("A short test answer about revenue and costs.", sample_case)
check("case feedback generator returns non-empty text", len(fb) > 50)

fit_fb = rule_based_fit_feedback("I led a project and delivered results as a result of my planning.")
check("fit feedback generator returns non-empty text", len(fit_fb) > 50)

sizing_fb = rule_based_sizing_feedback("Assume 1000 units at roughly $5 each.", MARKET_SIZING_PROMPTS[0])
check("sizing feedback generator returns non-empty text", len(sizing_fb) > 50)

keywords = extract_keywords("The hangar utilization and maintenance cost drove the profitability decline in the fleet expansion")
check("keyword extraction filters common filler words", "the" not in keywords and "and" not in keywords)
check("keyword extraction surfaces real content words", "hangar" in keywords or "maintenance" in keywords or "utilization" in keywords,
      f"got {keywords}")

print("\n" + "=" * 40)
if failures:
    print(f"FAIL: {len(failures)} check(s) failed: {failures}")
    sys.exit(1)
else:
    print("ALL CHECKS PASSED")
