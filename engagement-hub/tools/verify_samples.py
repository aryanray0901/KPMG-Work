"""
Verifies each Engagement Hub module against its sample data: checks that
parsing/analysis produces the expected structure and that generated PPTX
files have no off-slide or overlapping text (the same rendering check used
to catch a real bug during development).

Usage: python tools/verify_samples.py
"""
import os
import sys
import shutil
import subprocess
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from app import (
    parse_meeting_notes, build_action_tracker_xlsx, build_action_tracker_pptx,
    build_followup_email, analyze_trend, build_summary_pptx, build_summary_xlsx,
    analyze_benchmark, build_benchmark_pptx, build_benchmark_xlsx,
    SOFFICE_PATH, HAVE_FITZ,
)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(BASE, "sample_files")
OUT = tempfile.mkdtemp(prefix="eh_verify_")

failures = []


def check(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {name}" + (f" — {detail}" if detail and not condition else ""))
    if not condition:
        failures.append(name)


def render_and_check_overlap(pptx_path, label):
    if not (SOFFICE_PATH and HAVE_FITZ):
        print(f"[SKIP] {label} render check — LibreOffice/PyMuPDF not available")
        return
    import fitz
    profile_dir = tempfile.mkdtemp(prefix="lo_verify_")
    try:
        result = subprocess.run(
            [SOFFICE_PATH, "--headless", "--norestore",
             f"-env:UserInstallation=file://{profile_dir}",
             "--convert-to", "pdf", "--outdir", OUT, pptx_path],
            capture_output=True, timeout=90,
        )
        pdf_path = os.path.join(OUT, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
        check(f"{label}: converts to PDF", result.returncode == 0 and os.path.exists(pdf_path))
        if not os.path.exists(pdf_path):
            return
        doc = fitz.open(pdf_path)
        issues = 0
        for page in doc:
            spans = []
            for b in page.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l.get("spans", []):
                        if s["text"].strip():
                            spans.append((s["bbox"], s["text"]))
            pw, ph = page.rect.width, page.rect.height
            for bbox, txt in spans:
                x0, y0, x1, y1 = bbox
                if x0 < -1 or y0 < -1 or x1 > pw + 1 or y1 > ph + 1:
                    issues += 1
            for i in range(len(spans)):
                for j in range(i + 1, len(spans)):
                    b1, _ = spans[i]
                    b2, _ = spans[j]
                    ox = max(0, min(b1[2], b2[2]) - max(b1[0], b2[0]))
                    oy = max(0, min(b1[3], b2[3]) - max(b1[1], b2[1]))
                    a1 = (b1[2] - b1[0]) * (b1[3] - b1[1])
                    a2 = (b2[2] - b2[0]) * (b2[3] - b2[1])
                    if ox * oy > 0.3 * min(a1, a2) and ox * oy > 20:
                        issues += 1
        check(f"{label}: no off-slide or overlapping text", issues == 0, f"{issues} issue(s)")
        doc.close()
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


print("=== Module 1: Action Tracker ===")
with open(os.path.join(SAMPLES, "sample_meeting_notes.txt")) as f:
    notes_text = f.read()
attendees = ["Sarah Chen", "James Cole", "Priya Nair", "Marcus Lee"]
items = parse_meeting_notes(notes_text, attendees)
check("parses expected number of candidate lines", len(items) == 7, f"got {len(items)}")
check("assigns known attendees correctly",
      items[0]["owner"] == "Sarah Chen" and items[1]["owner"] == "James Cole",
      f"got {items[0]['owner']}, {items[1]['owner']}")
check("does not guess an owner for an unattributed line",
      items[2]["owner"] == "Unassigned", f"got {items[2]['owner']}")
xlsx_path = os.path.join(OUT, "actions.xlsx")
pptx_path = os.path.join(OUT, "actions.pptx")
build_action_tracker_xlsx(items, "Meridian Retail Kickoff", "July 15, 2026", xlsx_path)
build_action_tracker_pptx(items, "Meridian Retail Kickoff", "July 15, 2026", pptx_path)
build_followup_email(items, "Meridian Retail Kickoff", "July 15, 2026", "Alex Rivera")
check("Excel tracker created", os.path.exists(xlsx_path))
render_and_check_overlap(pptx_path, "Action summary slide")

print("\n=== Module 2: Executive Summary ===")
df = pd.read_excel(os.path.join(SAMPLES, "sample_trend_data.xlsx"))
analysis = analyze_trend(df, "Service Line", ["Q1", "Q2", "Q3", "Q4"], "M", metric_name="Revenue")
check("computes total growth", analysis["total_change_pct"] is not None and analysis["total_change_pct"] > 0)
check("identifies a top contributor", analysis["top_contributor"] is not None)
check("drafts at least 3 highlight bullets", len(analysis["bullets"]) >= 3, f"got {len(analysis['bullets'])}")
pptx_path2 = os.path.join(OUT, "summary.pptx")
build_summary_pptx(analysis, "Q4 Executive Summary", "Advisory practice revenue trend, Q1-Q4",
                    "Service Line", "M", pptx_path2)
build_summary_xlsx(analysis, "Q4 Executive Summary", "Service Line", "M", os.path.join(OUT, "summary.xlsx"))
render_and_check_overlap(pptx_path2, "Executive summary slide")

print("\n=== Module 3: Benchmarking ===")
df2 = pd.read_excel(os.path.join(SAMPLES, "sample_benchmark_data.xlsx"))
rows = analyze_benchmark(df2, "Metric", "Target", ["Peer A", "Peer B", "Peer C", "Peer D"])
check("analyzes all metric rows", len(rows) == 5, f"got {len(rows)}")
dso_row = next(r for r in rows if "Days Sales" in r["metric"])
check("classifies Days Sales Outstanding as lower-is-better", dso_row["higher_is_good"] is False)
revenue_row = next(r for r in rows if "Revenue Growth" in r["metric"])
check("classifies Revenue Growth as higher-is-better", revenue_row["higher_is_good"] is True)
pptx_path3 = os.path.join(OUT, "benchmark.pptx")
build_benchmark_pptx(rows, "Q4 Benchmarking Analysis", "Meridian Retail Group vs. peer set",
                      "Meridian Retail Group", pptx_path3)
build_benchmark_xlsx(rows, "Q4 Benchmarking Analysis", "Meridian Retail Group", os.path.join(OUT, "benchmark.xlsx"))
render_and_check_overlap(pptx_path3, "Benchmarking slide")

shutil.rmtree(OUT, ignore_errors=True)

print("\n" + "=" * 40)
if failures:
    print(f"FAIL: {len(failures)} check(s) failed: {failures}")
    sys.exit(1)
else:
    print("ALL CHECKS PASSED")
