"""
KPMG Deck Refresh
------------
Upload a PowerPoint OR Excel file, upload/paste new data, and this app finds
matching numbers by label, fuzzy-matches across period changes (Q3 -> Q4,
FY25 -> FY26, etc.), rewrites headings to match, and writes new values
straight into the existing cells, runs, and chart data so the original
formatting and file structure remain in place. The app runs locally and does
not send files to an external service.
"""

import os
import re
import io
import json
import time
import uuid
import shutil
import subprocess
import tempfile
import threading
from collections import Counter

from flask import Flask, request, render_template, send_file, redirect, url_for, flash, abort
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from openpyxl import load_workbook
import pandas as pd
from rapidfuzz import fuzz

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SESSIONS_DIR = os.path.join(BASE_DIR, "sessions")
os.makedirs(SESSIONS_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = "deck-refresh-local-secret"
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB

# ---------------------------------------------------------------------------
# REAL SLIDE RENDERING (LibreOffice -> PDF -> PNG per slide)
# Gives an actual visual render of the real PPTX, not a reconstruction.
# Falls back gracefully (rendering_ok=False) if LibreOffice isn't installed.
# ---------------------------------------------------------------------------

SOFFICE_PATH = shutil.which("soffice") or shutil.which("libreoffice")
POWERSHELL_PATH = shutil.which("powershell") or shutil.which("pwsh")
_soffice_lock = threading.Lock()
_powerpoint_lock = threading.Lock()

try:
    import fitz  # PyMuPDF
    HAVE_FITZ = True
except ImportError:
    HAVE_FITZ = False


def _render_pptx_with_powerpoint(pptx_path, out_dir, prefix):
    """Render with the installed Microsoft PowerPoint desktop app on Windows.
    This uses PowerPoint's own export engine, so the browser preview matches
    what the user sees when opening the PPTX in PowerPoint.
    """
    if os.name != "nt" or not POWERSHELL_PATH:
        return None

    os.makedirs(out_dir, exist_ok=True)
    script_path = os.path.join(out_dir, f"render_{prefix}.ps1")
    script = r'''
param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputDir,
    [Parameter(Mandatory=$true)][string]$Prefix
)
$ErrorActionPreference = "Stop"
$powerpoint = $null
$presentation = $null
try {
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerpoint.Presentations.Open($InputPath, -1, 0, 0)
    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    $exportWidth = 1600
    $exportHeight = [int][Math]::Round($exportWidth * ($slideHeight / $slideWidth))
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $outputPath = Join-Path $OutputDir ("{0}_{1}.png" -f $Prefix, $i)
        $presentation.Slides.Item($i).Export($outputPath, "PNG", $exportWidth, $exportHeight)
    }
}
finally {
    if ($presentation -ne $null) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($powerpoint -ne $null) {
        $powerpoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''
    try:
        with open(script_path, "w", encoding="utf-8") as f:
            f.write(script)
        with _powerpoint_lock:
            result = subprocess.run(
                [POWERSHELL_PATH, "-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-File", script_path, "-InputPath", os.path.abspath(pptx_path),
                 "-OutputDir", os.path.abspath(out_dir), "-Prefix", prefix],
                capture_output=True, timeout=180,
            )
        if result.returncode != 0:
            return None
        paths = []
        i = 1
        while True:
            path = os.path.join(out_dir, f"{prefix}_{i}.png")
            if not os.path.exists(path):
                break
            paths.append(path)
            i += 1
        return paths or None
    except Exception:
        return None
    finally:
        try:
            os.remove(script_path)
        except OSError:
            pass


def _convert_pptx_to_pdf(pptx_path, out_dir):
    if not SOFFICE_PATH:
        return None
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        with _soffice_lock:
            result = subprocess.run(
                [SOFFICE_PATH, "--headless", "--norestore",
                 f"-env:UserInstallation=file://{profile_dir}",
                 "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
                capture_output=True, timeout=120,
            )
        pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
        return pdf_path if (result.returncode == 0 and os.path.exists(pdf_path)) else None
    except Exception:
        return None
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def _render_pptx_with_libreoffice(pptx_path, out_dir, prefix):
    if not (SOFFICE_PATH and HAVE_FITZ):
        return None
    os.makedirs(out_dir, exist_ok=True)
    pdf_path = _convert_pptx_to_pdf(pptx_path, out_dir)
    if not pdf_path:
        return None
    try:
        doc = fitz.open(pdf_path)
        paths = []
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img_path = os.path.join(out_dir, f"{prefix}_{i+1}.png")
            pix.save(img_path)
            paths.append(img_path)
        doc.close()
        return paths or None
    except Exception:
        return None
    finally:
        try:
            os.remove(pdf_path)
        except OSError:
            pass


def render_pptx_to_images(pptx_path, out_dir, prefix):
    """Return slide image paths and the rendering engine used.

    Windows first uses the installed Microsoft PowerPoint application. Other
    machines, or Windows machines without PowerPoint, fall back to
    LibreOffice plus PyMuPDF.
    """
    powerpoint_paths = _render_pptx_with_powerpoint(pptx_path, out_dir, prefix)
    if powerpoint_paths:
        return powerpoint_paths, "Microsoft PowerPoint"

    libreoffice_paths = _render_pptx_with_libreoffice(pptx_path, out_dir, prefix)
    if libreoffice_paths:
        return libreoffice_paths, "LibreOffice"

    return None, None

@app.after_request
def _no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response


# ---------------------------------------------------------------------------
# SHARED HELPERS
# ---------------------------------------------------------------------------

NUMBER_RE = re.compile(
    r"(?P<prefix>[$€£]?)\s*(?P<num>-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|\-?\d+(?:\.\d+)?)"
    r"\s*(?P<mag>[KkMmBbTt])?\s*(?P<suffix>%?)"
)
MAGNITUDE = {"K": 1_000, "M": 1_000_000, "B": 1_000_000_000, "T": 1_000_000_000_000}

PERIOD_WORDS = re.compile(
    r"\b(q1|q2|q3|q4|quarter\s*[1-4]|fy\s*\d{2,4}|20\d{2}|h1|h2|jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|"
    r"january|february|march|april|june|july|august|september|october|november|december|ytd|mtd)\b",
    re.IGNORECASE,
)

QUARTER_TOKEN_RE = re.compile(r"\bQ[1-4]\b", re.IGNORECASE)
FY_TOKEN_RE = re.compile(r"\bFY\s?\d{2,4}\b", re.IGNORECASE)
YEAR_TOKEN_RE = re.compile(r"\b20\d{2}\b")


def normalize_label(text):
    t = (text or "").lower()
    t = re.sub(r"\([^)]*\)", " ", t)  # drop unit annotations like "($M)", "(%)"
    t = PERIOD_WORDS.sub(" ", t)
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def looks_numeric(text):
    text = (text or "").strip()
    if not text:
        return False
    m = NUMBER_RE.fullmatch(text)
    return m is not None and re.search(r"\d", text) is not None


def extract_number_value(text):
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    m = NUMBER_RE.search(str(text))
    if not m:
        return None
    raw = m.group("num").replace(",", "")
    try:
        return float(raw)
    except ValueError:
        return None


def format_new_value(original_text, new_value):
    """Render new_value using the same style ($, %, commas, decimals, K/M/B/T) as original_text."""
    original_text = str(original_text)
    currency = ""
    for sym in ("$", "€", "£"):
        if sym in original_text:
            currency = sym
            break
    has_percent = "%" in original_text
    mag_match = re.search(r"\d\s*([KkMmBbTt])\b", original_text)
    magnitude = mag_match.group(1) if mag_match else ""
    has_comma = "," in original_text
    decimals = 0
    m = re.search(r"\.(\d+)", original_text)
    if m:
        decimals = len(m.group(1))
    is_negative = new_value < 0
    abs_val = abs(new_value)
    if decimals:
        num_str = f"{abs_val:,.{decimals}f}" if has_comma else f"{abs_val:.{decimals}f}"
    else:
        if abs_val == int(abs_val):
            num_str = f"{int(abs_val):,}" if has_comma else f"{int(abs_val)}"
        else:
            num_str = f"{abs_val:,.2f}" if has_comma else f"{abs_val:.2f}"
    sign = "-" if is_negative else ""
    result = f"{sign}{currency}{num_str}{magnitude}"
    if has_percent:
        result += "%"
    return result


def detect_period_tokens(texts):
    q_counter, fy_counter, year_counter = Counter(), Counter(), Counter()
    for t in texts:
        if not t:
            continue
        t = str(t)
        for m in FY_TOKEN_RE.finditer(t):
            fy_counter[re.sub(r"\s+", "", m.group(0)).upper()] += 1
        remainder = FY_TOKEN_RE.sub(" ", t)
        for m in QUARTER_TOKEN_RE.finditer(remainder):
            q_counter[m.group(0).upper()] += 1
        remainder2 = QUARTER_TOKEN_RE.sub(" ", remainder)
        for m in YEAR_TOKEN_RE.finditer(remainder2):
            year_counter[m.group(0)] += 1
    return q_counter, fy_counter, year_counter


def compute_period_replacements(deck_texts, source_pairs, pasted_text):
    deck_q, deck_fy, deck_year = detect_period_tokens(deck_texts)
    src_texts = [lbl for lbl, _ in source_pairs]
    if pasted_text:
        src_texts.append(pasted_text)
    src_q, src_fy, src_year = detect_period_tokens(src_texts)

    replacements = {}
    for deck_counter, src_counter in ((deck_q, src_q), (deck_fy, src_fy), (deck_year, src_year)):
        if not deck_counter or not src_counter:
            continue
        old = deck_counter.most_common(1)[0][0]
        new = src_counter.most_common(1)[0][0]
        if old.upper() != new.upper():
            replacements[old] = new
    return replacements


KEY_METRIC_WORDS = {
    "revenue", "cost", "costs", "expense", "expenses", "margin", "income",
    "profit", "budget", "actual", "variance", "ebitda", "ebit", "loss",
}


def match_targets_to_source(targets, source_pairs, threshold=60):
    norm_sources = [(normalize_label(lbl), lbl, val) for lbl, val in source_pairs]
    matches = []
    for t in targets:
        norm_t = normalize_label(t["label"])
        t_tokens = set(norm_t.split())
        t_key = t_tokens & KEY_METRIC_WORDS
        best, best_score, best_metric = None, 0, None
        top_seen_score = 0
        if norm_t:
            for norm_s, orig_s, val in norm_sources:
                if not norm_s:
                    continue
                s_tokens = set(norm_s.split())
                if not s_tokens:
                    continue
                score = fuzz.token_sort_ratio(norm_t, norm_s)
                top_seen_score = max(top_seen_score, score)
                # A long shared descriptive phrase (e.g. "Technology
                # Consulting") can make "...Costs" and "...Revenue" look
                # highly similar by raw character overlap even though
                # they're different metrics. If the target names a specific
                # metric keyword, the candidate must contain that same
                # keyword too -- a bare label with no keyword at all, or one
                # naming a *different* keyword, must not satisfy it.
                s_key = s_tokens & KEY_METRIC_WORDS
                if t_key and not (t_key <= s_key):
                    continue
                # Among everything that survives the keyword gate, prefer
                # the candidate whose token set differs least from the
                # target's -- exact-word overlap is a far more reliable
                # financial-label signal than raw character similarity
                # (which can rank "Advisory Budget" above the correct
                # "Deal Advisory Actual" by a hair, or "Total Revenue"
                # above "Tax Revenue" for a target literally saying "Tax").
                # Ties broken by fuzzy score, then by preferring an
                # explicit "Actual" figure -- the sensible default when the
                # target doesn't say which period-type it wants.
                sym_diff = len(t_tokens ^ s_tokens)
                actual_bonus = 0 if "actual" in s_tokens else 1
                metric = (sym_diff, actual_bonus, -score)
                if best_metric is None or metric < best_metric:
                    best_metric, best_score, best = metric, score, (orig_s, val)
        entry = dict(t)
        if best and best_score >= threshold:
            entry["matched_label"] = best[0]
            entry["new_value"] = best[1]
            entry["score"] = round(best_score, 1)
        else:
            entry["matched_label"] = None
            entry["new_value"] = None
            entry["score"] = round(top_seen_score, 1)
        if entry["new_value"] is not None:
            entry["new_text_preview"] = format_new_value(entry["original"], entry["new_value"])
        else:
            entry["new_text_preview"] = None
        matches.append(entry)
    return matches


# ---------------------------------------------------------------------------
# SOURCE DATA EXTRACTION (label -> new value), shared by pptx & xlsx targets
# ---------------------------------------------------------------------------

def extract_from_spreadsheet(path):
    pairs = []
    try:
        if path.lower().endswith(".csv"):
            frames = {"csv": pd.read_csv(path)}
        else:
            raw = pd.read_excel(path, sheet_name=None)
            frames = raw if isinstance(raw, dict) else {"Sheet1": raw}
    except Exception:
        return pairs

    for sheet_name, frame in frames.items():
        frame = frame.dropna(how="all").dropna(axis=1, how="all")
        if frame.empty:
            continue
        cols = [str(c) for c in frame.columns]
        first_col_is_label = (
            not pd.api.types.is_numeric_dtype(frame.iloc[:, 0]) if frame.shape[1] > 0 else False
        )

        if frame.shape[1] == 2 and first_col_is_label:
            for _, row in frame.iterrows():
                label = str(row.iloc[0]).strip()
                val = row.iloc[1]
                if pd.notna(val) and label and label.lower() != "nan":
                    try:
                        pairs.append((label, float(val)))
                    except (ValueError, TypeError):
                        pass
            continue

        for _, row in frame.iterrows():
            row_label = str(row.iloc[0]).strip() if first_col_is_label else ""
            start = 1 if first_col_is_label else 0
            data_col_count = len(cols) - start
            for c in range(start, len(cols)):
                val = row.iloc[c]
                if pd.isna(val):
                    continue
                try:
                    val = float(val)
                except (ValueError, TypeError):
                    continue
                col_label = cols[c]
                combined = f"{row_label} {col_label}".strip()
                if combined:
                    pairs.append((combined, val))
                # Bare row/column labels are only unambiguous when there's a
                # single data column (or single data row) -- otherwise every
                # row sharing that column header (e.g. every line item's
                # "Q4 Actual") would collide under the same bare label with
                # different values, and whichever happened to be inserted
                # first would silently win.
                if data_col_count == 1 and row_label:
                    pairs.append((row_label, val))
                if len(frame) == 1 and col_label and col_label.lower() not in ("unnamed: 0",):
                    pairs.append((col_label, val))
    return pairs


def extract_from_text(text):
    pairs = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = re.split(r"[:=\t]|(?:\s-\s)", line, maxsplit=1)
        if len(parts) == 2:
            label, rest = parts
            val = extract_number_value(rest)
            if val is not None and label.strip():
                pairs.append((label.strip(), val))
    return pairs


def gather_source_pairs(data_path, pasted_text):
    source_pairs = []
    if data_path:
        fname = data_path.lower()
        if fname.endswith((".xlsx", ".xls", ".csv")):
            source_pairs += extract_from_spreadsheet(data_path)
        elif fname.endswith(".pptx"):
            _, targets = extract_targets_pptx(data_path)
            for t in targets:
                val = extract_number_value(t["original"])
                if val is not None and t["label"]:
                    source_pairs.append((t["label"], val))
    if pasted_text:
        source_pairs += extract_from_text(pasted_text)
    return source_pairs


# ---------------------------------------------------------------------------
# PPTX: target extraction / editing / snapshot
# ---------------------------------------------------------------------------

def slide_title(slide):
    for shape in slide.shapes:
        if shape.has_text_frame and shape.is_placeholder:
            try:
                if shape.placeholder_format.type is not None and "TITLE" in str(shape.placeholder_format.type):
                    return shape.text_frame.text.strip()
            except Exception:
                pass
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            return shape.text_frame.text.strip().split("\n")[0]
    return ""


def extract_targets_pptx(pptx_path):
    prs = Presentation(pptx_path)
    targets = []
    for s_idx, slide in enumerate(prs.slides):
        title = slide_title(slide)
        for sh_idx, shape in enumerate(slide.shapes):
            if shape.has_text_frame:
                all_paras = shape.text_frame.paragraphs
                for p_idx, para in enumerate(all_paras):
                    for r_idx, run in enumerate(para.runs):
                        if looks_numeric(run.text):
                            other_text = "".join(
                                rr.text for k, rr in enumerate(para.runs) if k != r_idx
                            ).strip()
                            if not other_text:
                                # check sibling paragraphs in the same text box
                                # (e.g. a KPI tile: value on one line, label below it)
                                sibling_text = " ".join(
                                    pp.text.strip() for pi, pp in enumerate(all_paras)
                                    if pi != p_idx and pp.text.strip()
                                ).strip()
                                other_text = sibling_text
                            label = other_text or title or shape.name
                            targets.append({
                                "id": f"s{s_idx}-sh{sh_idx}-p{p_idx}-r{r_idx}",
                                "kind": "run",
                                "slide": s_idx + 1,
                                "label": label,
                                "context": f"Slide {s_idx+1}: {title}" if title else f"Slide {s_idx+1}",
                                "original": run.text.strip(),
                                "location": (s_idx, sh_idx, "para", p_idx, r_idx),
                            })

            if shape.has_chart:
                chart = shape.chart
                chart_title = ""
                try:
                    if chart.has_title and chart.chart_title.text_frame.text.strip():
                        chart_title = chart.chart_title.text_frame.text.strip()
                except Exception:
                    pass
                try:
                    plot = chart.plots[0]
                    categories = [str(c) for c in plot.categories]
                    series_list = list(chart.series)
                    multi_series = len(series_list) > 1
                    for se_idx, series in enumerate(series_list):
                        values = list(series.values)
                        for c_idx, cat in enumerate(categories):
                            if c_idx >= len(values) or values[c_idx] is None:
                                continue
                            val = values[c_idx]
                            label = f"{series.name} {cat}".strip() if multi_series else cat
                            targets.append({
                                "id": f"s{s_idx}-sh{sh_idx}-chart-se{se_idx}-c{c_idx}",
                                "kind": "chart_point",
                                "slide": s_idx + 1,
                                "label": label,
                                "context": (f"Slide {s_idx+1}: {chart_title or title} (chart)"
                                            if (chart_title or title) else f"Slide {s_idx+1} (chart)"),
                                "original": (f"{val:g}" if isinstance(val, float) else str(val)),
                                "location": (s_idx, sh_idx, "chart", se_idx, c_idx),
                            })
                except Exception:
                    pass

            if shape.has_table:
                table = shape.table
                nrows, ncols = len(table.rows), len(table.columns)
                col_headers = [table.cell(0, c).text.strip() for c in range(ncols)]
                for r in range(1, nrows):
                    row_header = table.cell(r, 0).text.strip()
                    for c in range(1, ncols):
                        cell_text = table.cell(r, c).text.strip()
                        if looks_numeric(cell_text):
                            label = f"{row_header} {col_headers[c]}".strip()
                            targets.append({
                                "id": f"s{s_idx}-sh{sh_idx}-tbl-r{r}-c{c}",
                                "kind": "table_cell",
                                "slide": s_idx + 1,
                                "label": label or f"{shape.name} row{r} col{c}",
                                "context": (f"Slide {s_idx+1}: {title} (table)" if title else f"Slide {s_idx+1} (table)"),
                                "original": cell_text,
                                "location": (s_idx, sh_idx, "table", r, c),
                            })
    return prs, targets


def collect_all_pptx_text(prs):
    texts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        texts.append(cell.text_frame.text)
            if shape.has_chart:
                try:
                    if shape.chart.has_title:
                        texts.append(shape.chart.chart_title.text_frame.text)
                except Exception:
                    pass
    return texts


def _build_period_pattern(replacements):
    if not replacements:
        return None, None
    keys_sorted = sorted(replacements.keys(), key=len, reverse=True)
    pattern = re.compile(r"\b(" + "|".join(re.escape(k) for k in keys_sorted) + r")\b", re.IGNORECASE)
    lookup = {k.upper(): v for k, v in replacements.items()}
    return pattern, lookup


def _replace_text(text, pattern, lookup):
    if not text or not pattern:
        return text
    return pattern.sub(lambda m: lookup.get(m.group(0).upper(), m.group(0)), text)


def apply_period_replacements_pptx(prs, replacements):
    pattern, lookup = _build_period_pattern(replacements)
    if not pattern:
        return

    def process_text_frame(tf):
        for para in tf.paragraphs:
            for run in para.runs:
                if run.text and pattern.search(run.text):
                    run.text = _replace_text(run.text, pattern, lookup)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                process_text_frame(shape.text_frame)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        process_text_frame(cell.text_frame)
            if shape.has_chart:
                try:
                    if shape.chart.has_title:
                        process_text_frame(shape.chart.chart_title.text_frame)
                except Exception:
                    pass


def apply_edits_pptx(pptx_path, out_path, confirmed_matches, period_replacements=None):
    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    chart_edits = {}
    pattern, lookup = _build_period_pattern(period_replacements or {})

    for m in confirmed_matches:
        loc = m["location"]
        s_idx, sh_idx = loc[0], loc[1]
        shape = list(slides[s_idx].shapes)[sh_idx]
        if loc[2] == "para":
            p_idx, r_idx = loc[3], loc[4]
            shape.text_frame.paragraphs[p_idx].runs[r_idx].text = m["new_text"]
        elif loc[2] == "table":
            r, c = loc[3], loc[4]
            cell = shape.table.cell(r, c)
            if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
                cell.text_frame.paragraphs[0].runs[0].text = m["new_text"]
                for extra in cell.text_frame.paragraphs[0].runs[1:]:
                    extra.text = ""
            else:
                cell.text = m["new_text"]
        elif loc[2] == "chart":
            se_idx, c_idx = loc[3], loc[4]
            numeric_val = extract_number_value(m["new_text"])
            if numeric_val is None:
                continue
            chart_edits.setdefault((s_idx, sh_idx), {})[(se_idx, c_idx)] = numeric_val

    # Rebuild data for every chart that had at least one confirmed numeric
    # edit OR needs its series/category names updated to the new period.
    charts_needing_period_update = set()
    if pattern:
        for slide_idx, slide in enumerate(slides):
            for sh_idx, shape in enumerate(slide.shapes):
                if not shape.has_chart:
                    continue
                chart = shape.chart
                try:
                    plot = chart.plots[0]
                    names_and_cats = [str(c) for c in plot.categories] + [s.name or "" for s in chart.series]
                    if any(pattern.search(t) for t in names_and_cats):
                        charts_needing_period_update.add((slide_idx, sh_idx))
                except Exception:
                    pass

    for (s_idx, sh_idx) in set(chart_edits.keys()) | charts_needing_period_update:
        shape = list(slides[s_idx].shapes)[sh_idx]
        chart = shape.chart
        plot = chart.plots[0]
        categories = [_replace_text(str(c), pattern, lookup) for c in plot.categories]
        new_chart_data = CategoryChartData()
        new_chart_data.categories = categories
        edits = chart_edits.get((s_idx, sh_idx), {})
        for se_idx, series in enumerate(chart.series):
            values = list(series.values)
            for c_idx in range(len(values)):
                if (se_idx, c_idx) in edits:
                    values[c_idx] = edits[(se_idx, c_idx)]
            new_name = _replace_text(series.name or "", pattern, lookup)
            new_chart_data.add_series(new_name, values)
        chart.replace_data(new_chart_data)

    apply_period_replacements_pptx(prs, period_replacements or {})
    prs.save(out_path)


def snapshot_pptx(path):
    """Everything needed to render a before/after comparison in the browser:
    per slide, every table (as a 2D grid) and every chart (categories/series/
    colors) as plain data."""
    prs = Presentation(path)
    slides_out = []
    for slide in prs.slides:
        title = slide_title(slide)
        tables, charts = [], []
        for shape in slide.shapes:
            if shape.has_table:
                t = shape.table
                rows = [[t.cell(r, c).text for c in range(len(t.columns))] for r in range(len(t.rows))]
                tables.append(rows)
            if shape.has_chart:
                chart = shape.chart
                try:
                    plot = chart.plots[0]
                    cats = [str(c) for c in plot.categories]
                    series_out = []
                    for series in chart.series:
                        vals = [float(v) if v is not None else 0.0 for v in series.values]
                        colors = []
                        for pt in series.points:
                            try:
                                colors.append("#" + str(pt.format.fill.fore_color.rgb))
                            except Exception:
                                colors.append(None)
                        series_out.append({"name": series.name or "", "values": vals, "colors": colors})
                    chart_type = "pie" if "PIE" in str(chart.chart_type) else (
                        "line" if "LINE" in str(chart.chart_type) else "bar"
                    )
                    charts.append({"type": chart_type, "categories": cats, "series": series_out})
                except Exception:
                    pass
        slides_out.append({"title": title or f"Slide", "tables": tables, "charts": charts})
    return slides_out


# ---------------------------------------------------------------------------
# XLSX: target extraction / editing / snapshot
# Excel charts are normally bound to live cell ranges, so simply updating
# cell values (and never touching chart XML) makes the embedded chart
# reflect the new numbers automatically, with all original formatting,
# when the file is opened in Excel.
# ---------------------------------------------------------------------------

def extract_targets_xlsx(path):
    wb = load_workbook(path, data_only=False)
    targets = []
    for ws in wb.worksheets:
        sheet_id = re.sub(r"[^A-Za-z0-9]+", "_", ws.title).strip("_") or "sheet"
        max_row, max_col = ws.max_row, ws.max_column
        if max_row < 2 or max_col < 2:
            continue
        headers = [ws.cell(row=1, column=c).value for c in range(1, max_col + 1)]
        for r in range(2, max_row + 1):
            row_label = ws.cell(row=r, column=1).value
            row_label = str(row_label).strip() if row_label is not None else ""
            for c in range(2, max_col + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    col_header = headers[c - 1]
                    col_header = str(col_header).strip() if col_header is not None else ""
                    label = f"{row_label} {col_header}".strip()
                    original = f"{val:,.2f}" if (val != int(val)) else f"{int(val):,}"
                    targets.append({
                        "id": f"{sheet_id}-r{r}-c{c}",
                        "kind": "xlsx_cell",
                        "sheet": ws.title,
                        "label": label or col_header or row_label,
                        "context": ws.title,
                        "original": original,
                        "location": (ws.title, r, c),
                    })
    wb.close()
    return targets


def collect_all_xlsx_text(path):
    wb = load_workbook(path, data_only=False)
    texts = []
    for ws in wb.worksheets:
        texts.append(ws.title)
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    texts.append(cell.value)
    wb.close()
    return texts


def apply_period_replacements_xlsx(wb, replacements):
    if not replacements:
        return
    keys_sorted = sorted(replacements.keys(), key=len, reverse=True)
    pattern = re.compile(r"\b(" + "|".join(re.escape(k) for k in keys_sorted) + r")\b", re.IGNORECASE)
    lookup = {k.upper(): v for k, v in replacements.items()}

    def sub_func(m):
        return lookup.get(m.group(0).upper(), m.group(0))

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and pattern.search(cell.value):
                    cell.value = pattern.sub(sub_func, cell.value)


def apply_edits_xlsx(xlsx_path, out_path, confirmed_matches, period_replacements=None):
    wb = load_workbook(xlsx_path)
    for m in confirmed_matches:
        sheet, r, c = m["location"]
        ws = wb[sheet]
        ws.cell(row=r, column=c).value = m["new_value"]
    apply_period_replacements_xlsx(wb, period_replacements or {})
    wb.save(out_path)
    wb.close()


def snapshot_xlsx(path):
    """Per-sheet grid of values, plus a simple reconstructed chart (first
    text column as categories, first numeric column as values) so the
    browser can show a comparable visual. This is a rendering aid only —
    the real embedded Excel chart (with its original formatting) is left
    completely untouched in the actual file and updates itself in Excel."""
    wb = load_workbook(path, data_only=True)
    sheets_out = []
    for ws in wb.worksheets:
        max_row, max_col = min(ws.max_row, 60), min(ws.max_column, 12)
        rows = []
        for r in range(1, max_row + 1):
            rows.append([ws.cell(row=r, column=c).value for c in range(1, max_col + 1)])

        chart = None
        if len(rows) >= 2 and len(rows[0]) >= 2:
            cats, vals = [], []
            for r in rows[1:]:
                label = r[0]
                num = None
                for cell_val in r[1:]:
                    if isinstance(cell_val, (int, float)) and not isinstance(cell_val, bool):
                        num = float(cell_val)
                        break
                if label is not None and num is not None:
                    cats.append(str(label))
                    vals.append(num)
            if cats and vals:
                chart = {"type": "bar", "categories": cats, "series": [{"name": "Value", "values": vals, "colors": []}]}

        sheets_out.append({"title": ws.title, "tables": [rows], "charts": [chart] if chart else []})
    wb.close()
    return sheets_out


# ---------------------------------------------------------------------------
# ROUTES
# ---------------------------------------------------------------------------

def _session_dir(sid):
    d = os.path.join(SESSIONS_DIR, sid)
    if not os.path.isdir(d):
        abort(404)
    return d


def _save_session_meta(sess_dir, **kwargs):
    path = os.path.join(sess_dir, "meta.json")
    data = {}
    if os.path.exists(path):
        with open(path) as f:
            data = json.load(f)
    data.update(kwargs)
    with open(path, "w") as f:
        json.dump(data, f)
    return data


def _load_session_meta(sess_dir):
    path = os.path.join(sess_dir, "meta.json")
    if not os.path.exists(path):
        return {}
    with open(path) as f:
        return json.load(f)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/sample/<name>", methods=["GET"])
def sample_file(name):
    allowed = {
        "original": ("kpmg_advisory_q3_original.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        "data": ("kpmg_advisory_q4_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "expected": ("kpmg_advisory_q4_expected.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        "verification": ("sample_verification.txt", "text/plain"),
    }
    if name not in allowed:
        abort(404)
    filename, mimetype = allowed[name]
    path = os.path.join(BASE_DIR, "sample_files", filename)
    if not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=filename, mimetype=mimetype)


@app.route("/process", methods=["POST"])
def process():
    primary_file = request.files.get("primary_file")
    if not primary_file or primary_file.filename == "":
        flash("Please upload a PowerPoint (.pptx) or Excel (.xlsx) file.")
        return redirect(url_for("index"))

    fname = primary_file.filename.lower()
    if fname.endswith(".pptx"):
        file_type = "pptx"
    elif fname.endswith((".xlsx", ".xls")):
        file_type = "xlsx"
    else:
        flash("That file type isn't supported. Upload a .pptx or .xlsx file.")
        return redirect(url_for("index"))

    sid = uuid.uuid4().hex[:12]
    sess_dir = os.path.join(SESSIONS_DIR, sid)
    os.makedirs(sess_dir, exist_ok=True)

    primary_path = os.path.join(sess_dir, "original." + ("pptx" if file_type == "pptx" else "xlsx"))
    primary_file.save(primary_path)

    data_path = None
    data_file = request.files.get("data_file")
    if data_file and data_file.filename:
        data_path = os.path.join(sess_dir, "data_" + data_file.filename)
        data_file.save(data_path)

    pasted_text = request.form.get("pasted_text", "").strip()
    source_pairs = gather_source_pairs(data_path, pasted_text)

    if not source_pairs:
        flash("No usable data found in what you provided (Excel/CSV/text/PPTX). Please check the format.")
        shutil.rmtree(sess_dir, ignore_errors=True)
        return redirect(url_for("index"))

    if file_type == "pptx":
        prs, targets = extract_targets_pptx(primary_path)
        deck_texts = collect_all_pptx_text(prs)
    else:
        targets = extract_targets_xlsx(primary_path)
        deck_texts = collect_all_xlsx_text(primary_path)

    matches = match_targets_to_source(targets, source_pairs)
    period_replacements = compute_period_replacements(deck_texts, source_pairs, pasted_text)

    with open(os.path.join(sess_dir, "matches.json"), "w") as f:
        json.dump(matches, f)
    with open(os.path.join(sess_dir, "period_replacements.json"), "w") as f:
        json.dump(period_replacements, f)
    _save_session_meta(sess_dir, file_type=file_type, original_filename=primary_file.filename,
                        source_count=len(source_pairs))

    return redirect(url_for("review", sid=sid))


@app.route("/review/<sid>", methods=["GET"])
def review(sid):
    sess_dir = _session_dir(sid)
    matches_path = os.path.join(sess_dir, "matches.json")
    if not os.path.exists(matches_path):
        flash("That session has expired. Please upload again.")
        return redirect(url_for("index"))

    with open(matches_path) as f:
        matches = json.load(f)
    period_path = os.path.join(sess_dir, "period_replacements.json")
    period_replacements = {}
    if os.path.exists(period_path):
        with open(period_path) as f:
            period_replacements = json.load(f)

    meta = _load_session_meta(sess_dir)
    matched = [m for m in matches if m["matched_label"]]
    unmatched = [m for m in matches if not m["matched_label"]]

    return render_template(
        "preview.html",
        sid=sid,
        matched=matched,
        unmatched=unmatched,
        source_count=meta.get("source_count", 0),
        period_replacements=period_replacements,
        file_type=meta.get("file_type", "pptx"),
    )


@app.route("/apply/<sid>", methods=["POST"])
def apply(sid):
    sess_dir = _session_dir(sid)
    matches_path = os.path.join(sess_dir, "matches.json")
    if not os.path.exists(matches_path):
        flash("That session has expired. Please upload again.")
        return redirect(url_for("index"))

    meta = _load_session_meta(sess_dir)
    file_type = meta.get("file_type", "pptx")
    primary_path = os.path.join(sess_dir, "original." + ("pptx" if file_type == "pptx" else "xlsx"))

    with open(matches_path) as f:
        matches = json.load(f)
    period_path = os.path.join(sess_dir, "period_replacements.json")
    period_replacements = {}
    if os.path.exists(period_path):
        with open(period_path) as f:
            period_replacements = json.load(f)

    confirmed = []
    for m in matches:
        if request.form.get(f"apply_{m['id']}") == "on" and m.get("new_text_preview"):
            override = request.form.get(f"value_{m['id']}", "").strip()
            new_text = override if override else m["new_text_preview"]
            entry = dict(m)
            entry["new_text"] = new_text
            if entry.get("new_value") is None:
                entry["new_value"] = extract_number_value(new_text)
            confirmed.append(entry)

    out_path = os.path.join(sess_dir, "updated." + ("pptx" if file_type == "pptx" else "xlsx"))
    if file_type == "pptx":
        apply_edits_pptx(primary_path, out_path, confirmed, period_replacements)
    else:
        apply_edits_xlsx(primary_path, out_path, confirmed, period_replacements)

    rendering_ok = False
    slide_count = 0
    if file_type == "pptx":
        render_dir = os.path.join(sess_dir, "render")
        shutil.rmtree(render_dir, ignore_errors=True)
        orig_images, orig_engine = render_pptx_to_images(primary_path, render_dir, "original")
        new_images, new_engine = render_pptx_to_images(out_path, render_dir, "updated")
        if orig_images and new_images and len(orig_images) == len(new_images):
            rendering_ok = True
            slide_count = len(orig_images)
            render_engine = orig_engine if orig_engine == new_engine else f"{orig_engine} / {new_engine}"
        else:
            render_engine = None
    else:
        render_engine = None

    _save_session_meta(
        sess_dir,
        applied_count=len(confirmed),
        rendering_ok=rendering_ok,
        slide_count=slide_count,
        render_engine=render_engine,
    )
    return redirect(url_for("result", sid=sid))


@app.route("/result/<sid>", methods=["GET"])
def result(sid):
    sess_dir = _session_dir(sid)
    meta = _load_session_meta(sess_dir)
    file_type = meta.get("file_type", "pptx")
    primary_path = os.path.join(sess_dir, "original." + ("pptx" if file_type == "pptx" else "xlsx"))
    updated_path = os.path.join(sess_dir, "updated." + ("pptx" if file_type == "pptx" else "xlsx"))
    if not os.path.exists(updated_path):
        flash("That session has expired. Please upload again.")
        return redirect(url_for("index"))

    rendering_ok = meta.get("rendering_ok", False)

    if rendering_ok:
        return render_template(
            "result.html",
            sid=sid,
            file_type=file_type,
            rendering_ok=True,
            slide_count=meta.get("slide_count", 0),
            applied_count=meta.get("applied_count", 0),
            render_engine=meta.get("render_engine", "PowerPoint renderer"),
            original_filename=meta.get("original_filename", "original.pptx"),
            updated_filename="updated_" + meta.get("original_filename", "presentation.pptx"),
        )

    # Fallback: no LibreOffice/PyMuPDF available (or xlsx), show the
    # reconstructed data/chart comparison instead of real slide images.
    if file_type == "pptx":
        old_snap = snapshot_pptx(primary_path)
        new_snap = snapshot_pptx(updated_path)
    else:
        old_snap = snapshot_xlsx(primary_path)
        new_snap = snapshot_xlsx(updated_path)

    tabs = []
    for i in range(max(len(old_snap), len(new_snap))):
        old_s = old_snap[i] if i < len(old_snap) else {"title": "", "tables": [], "charts": []}
        new_s = new_snap[i] if i < len(new_snap) else {"title": "", "tables": [], "charts": []}
        tabs.append({
            "label": new_s.get("title") or old_s.get("title") or f"{'Slide' if file_type=='pptx' else 'Sheet'} {i+1}",
            "old": old_s,
            "new": new_s,
        })

    return render_template(
        "result.html",
        sid=sid,
        file_type=file_type,
        rendering_ok=False,
        tabs_json=json.dumps(tabs),
        applied_count=meta.get("applied_count", 0),
        soffice_missing=not ((os.name == "nt" and POWERSHELL_PATH) or (SOFFICE_PATH and HAVE_FITZ)),
    )


@app.route("/slide_image/<sid>/<which>/<int:n>", methods=["GET"])
def slide_image(sid, which, n):
    sess_dir = _session_dir(sid)
    prefix = "original" if which == "original" else "updated"
    path = os.path.join(sess_dir, "render", f"{prefix}_{n}.png")
    if not os.path.exists(path):
        abort(404)
    return send_file(path, mimetype="image/png")


@app.route("/download/<sid>/<which>", methods=["GET"])
def download(sid, which):
    sess_dir = _session_dir(sid)
    meta = _load_session_meta(sess_dir)
    file_type = meta.get("file_type", "pptx")
    ext = "pptx" if file_type == "pptx" else "xlsx"
    mimetype = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        if file_type == "pptx"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    if which == "original":
        path = os.path.join(sess_dir, f"original.{ext}")
        name = "original." + ext
    elif which == "updated":
        path = os.path.join(sess_dir, f"updated.{ext}")
        name = "updated." + ext
    else:
        abort(404)
    if not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=name, mimetype=mimetype)


def _open_browser():
    import webbrowser
    import sys as _sys
    time.sleep(1.2)
    url = "http://127.0.0.1:5050"
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
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        threading.Thread(target=_open_browser, daemon=True).start()
    print("Starting KPMG Deck Refresh at http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050, debug=("--debug" in sys.argv))
