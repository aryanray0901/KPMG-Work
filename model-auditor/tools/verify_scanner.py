"""
Verifies the Model Auditor scanner against two purpose-built sample
workbooks: one with a deliberate bug in every detection category, and one
that's clean, to confirm the scanner catches real problems without
generating false positives on well-built formulas.

Usage: python tools/verify_scanner.py
"""
import os
import sys
import shutil
import tempfile
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import scan_workbook, build_report_pptx, build_annotated_xlsx, health_grade

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(BASE, "sample_files")
OUT = tempfile.mkdtemp(prefix="ma_verify_")
failures = []


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" — {detail}" if detail and not condition else ""))
    if not condition:
        failures.append(name)


print("=== Flawed sample model (one deliberate bug per category) ===")
issues, stats = scan_workbook(os.path.join(SAMPLES, "sample_flawed_model.xlsx"))
issue_types = {i["type"] for i in issues}
issues_at = defaultdict(list)
for i in issues:
    issues_at[(i["sheet"], i["cell"])].append(i["type"])

check("finds the hardcoded-input bug (B3, margin assumption)",
      "Hardcoded input" in issues_at[("Model", "B3")])
check("finds the row-consistency break (F2, hardcoded add instead of growth)",
      any("Inconsistent" in t for t in issues_at[("Model", "F2")]))
check("finds the circular reference (B5/B6)",
      "Circular reference" in issue_types)
check("finds the divide-by-zero error (C8)",
      len(issues_at[("Model", "C8")]) > 0)
check("health score reflects a heavily flawed model", stats["health_score"] < 60,
      f"got {stats['health_score']}")
check("assigns a failing or near-failing letter grade", health_grade(stats["health_score"]) in ("D", "F"),
      f"got {health_grade(stats['health_score'])}")

report_path = os.path.join(OUT, "report.pptx")
annotated_path = os.path.join(OUT, "annotated.xlsx")
build_report_pptx(issues, stats, "Model Health Report", "Verification run", "sample_flawed_model.xlsx", report_path)
build_annotated_xlsx(os.path.join(SAMPLES, "sample_flawed_model.xlsx"), issues, annotated_path)
check("report and annotated workbook created",
      os.path.exists(report_path) and os.path.exists(annotated_path))

import openpyxl
wb = openpyxl.load_workbook(annotated_path)
check("annotated workbook has an Audit Findings sheet", "Audit Findings" in wb.sheetnames)
f2_cell = wb["Model"]["F2"]
check("flagged cell F2 is highlighted", f2_cell.fill.fgColor.rgb not in (None, "00000000"))
check("flagged cell F2 has a comment explaining the finding", f2_cell.comment is not None)

print("\n=== Clean sample model (no deliberate bugs) ===")
clean_path = os.path.join(SAMPLES, "sample_clean_model.xlsx")
if os.path.exists(clean_path):
    clean_issues, clean_stats = scan_workbook(clean_path)
    check("zero issues found on a well-built model", len(clean_issues) == 0, f"got {len(clean_issues)}")
    check("perfect health score on a well-built model", clean_stats["health_score"] == 100,
          f"got {clean_stats['health_score']}")
else:
    print("[SKIP] sample_clean_model.xlsx not found")

shutil.rmtree(OUT, ignore_errors=True)

print("\n" + "=" * 40)
if failures:
    print(f"FAIL: {len(failures)} check(s) failed: {failures}")
    sys.exit(1)
else:
    print("ALL CHECKS PASSED")
