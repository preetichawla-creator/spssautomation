"""
STAGE 1: Auto-generate a variable/label mapping file from a Lighthouse Studio SPSS export.

Replaces the manual step of retyping variable names & labels into Excel.
Reads the .sav directly, applies known/reusable rename+cleaning rules, and
flags anything it can't confidently resolve so a human only reviews the
exceptions instead of building the whole file from scratch.

USAGE:
    python 01_generate_mapping.py <input.sav> <output_mapping.xlsx> [--config config.json]
"""

import sys
import re
import json
import argparse
from collections import defaultdict

import pyreadstat
import openpyxl
from openpyxl.styles import PatternFill, Font

# ---------------------------------------------------------------------------
# CONFIG (reusable across studies — edit/extend as new patterns appear)
# ---------------------------------------------------------------------------

# 1. System/admin fields Lighthouse always generates the same way.
#    Build this once, reuse forever across all studies.
DEFAULT_SYSTEM_FIELD_OVERRIDES = {
    "sys_RespNum":        ("Respondent_No", "Respondent Number"),
    "sys_StartTime":      ("Start_Time", "Start Time"),
    "sys_EndTime":        ("End_Time", "End Time"),
    "sys_ElapsedTime":    ("Elapsed_Time", "Elapsed Time"),
}

# 2. Project-specific admin/meta variable overrides (name + label are
#    study-specific business terms, not derivable from the raw label text).
#    Add rows here per study when they don't fit the generic rules.
DEFAULT_MANUAL_OVERRIDES = {
    "check_version": ("Version", "Version"),
    "lang":          ("Language", "Language"),
    "Coun":          ("Country", "Country"),
    "Password":      ("Password", "Password"),
}

# 3. Piping placeholders Lighthouse leaves in raw labels, e.g. [%ListLabel(Brand,1)%]
#    Map each placeholder pattern to the literal text it should resolve to
#    for THIS study. Update per project.
PIPING_REPLACEMENTS = {
    r"\[%\s*ListLabel\(\s*Brand\s*,\s*1\s*\)\s*%\]": "Vivo V19",
    r"\[%\s*ListLabel\(\s*B4heading\s*,\s*\d+\s*\)\s*%\]": "",  # section heading placeholders — no clean text available, strip
}

# 4. Country-code single-letter suffixes used in loop variables (A5x1M, B6aB, etc.)
COUNTRY_CODE_MAP = {
    "M": "Myanmar",
    "B": "Bangladesh",
}

# 5. Variables to always drop from final SPSS output (internal Lighthouse
#    helper variables, page-timing fields, etc.) — confirmed against your
#    team's actual practice: these never appeared in your finished mapping file.
EXCLUDE_PATTERNS = [
    r"^for[A-Z]",       # ALL internal Lighthouse helper variables (forB4a, forB3c,
                        # forCoun, forLang, etc.) — confirmed none of these ever
                        # appear in your team's finished mapping, across every
                        # prefix present in this file, not just forB4a.
    r"^sys_pagetime_",  # per-page timing fields
]

# 6. Topic-abbreviation dictionary for column-coded grids where the suffix is
#    derived from the QUESTION TOPIC, not the variable name itself
#    (e.g. C6B_c1 -> C6.Bangladesh_PMI). Extend per study.
COLUMN_TOPIC_MAP = {
    # (base_without_country, column_number): suffix
    ("C6", "c1"): "PMI",
    ("C6", "c2"): "MHI",
}

# ---------------------------------------------------------------------------
# LABEL CLEANING
# ---------------------------------------------------------------------------

NOTE_PATTERNS = [
    r"\[Note:.*?\]",
    r"\[Interviewer note:.*?\]",
    r"Note:\s*\(Interviewer[^)]*\)",
    r"\(Interviewer[^)]*\)",
    r"Interviewer need to speak out the options",
]

TAG_PATTERNS = [
    r"\?\s*\((Single|Multiple)\s+Answer\)\s*$",   # "...? (Single Answer)" -> strip both
    r"\s*\((Single|Multiple)\s+Answer\)\s*$",     # "...(Single Answer)" without leading '?'
]

LEADING_CODE_PATTERN = re.compile(r"^\s*[A-Za-z0-9_]+\s*-\s*")


def resolve_piping(text):
    for pattern, replacement in PIPING_REPLACEMENTS.items():
        text = re.sub(pattern, replacement, text)
    return text


def clean_label(raw_label):
    """Returns (cleaned_text, needs_review_flag, review_reason)"""
    if not raw_label:
        return "", True, "empty raw label"

    text = raw_label

    # strip leading "VarCode - "
    text = LEADING_CODE_PATTERN.sub("", text)

    # resolve known piping placeholders
    text = resolve_piping(text)

    # flag any UNRESOLVED piping placeholder before stripping notes
    unresolved_piping = re.findall(r"\[%.*?%\]", text)

    # strip trailing (Single Answer)/(Multiple Answer) tags, incl. preceding '?'
    for pat in TAG_PATTERNS:
        text = re.sub(pat, "", text)

    # strip known interviewer-note phrasing
    review_reason = []
    for pat in NOTE_PATTERNS:
        if re.search(pat, text, flags=re.IGNORECASE):
            text = re.sub(pat, "", text, flags=re.IGNORECASE)

    text = re.sub(r"\s{2,}", " ", text).strip()

    needs_review = False
    if unresolved_piping:
        needs_review = True
        review_reason.append(f"unresolved piping: {unresolved_piping}")
    if "Interviewer" in text or "interviewer" in text:
        needs_review = True
        review_reason.append("possible leftover interviewer note")
    if "[" in text or "]" in text:
        needs_review = True
        review_reason.append("possible leftover bracket note")

    return text, needs_review, "; ".join(review_reason)


# ---------------------------------------------------------------------------
# VARIABLE RENAMING
# ---------------------------------------------------------------------------

GRID_RC_PATTERN = re.compile(r"^(?P<base>.+)_r(?P<row>\d+)_c(?P<col>\d+)$")
X_LOOP_PATTERN = re.compile(r"^(?P<base>.+?)x(?P<loop>\d+)(?P<suffix>_\d+.*)?$")
X_COUNTRY_PATTERN = re.compile(r"^(?P<base>.+?)x(?P<loop>\d+)(?P<country>[A-Z])(?P<suffix>_.*)?$")
COUNTRY_LETTER_SUFFIX_PATTERN = re.compile(r"^(?P<base>.+?)(?P<country>[A-Z])$")
COLUMN_CODE_PATTERN = re.compile(r"^(?P<base>[A-Za-z0-9]+?)(?P<country>[A-Z])_(?P<col>c\d+)$")


def is_excluded(name):
    return any(re.search(pat, name) for pat in EXCLUDE_PATTERNS)


def suggest_renames(raw_names):
    """
    Returns dict raw_name -> (suggested_new_name, needs_review, reason)
    Uses two passes: first collect grid siblings to decide collapse vs. suffix,
    then apply per-variable rules.
    """
    results = {}

    # --- pass 1: group "_r{n}_c1" siblings by base to decide singleton collapse ---
    rc_groups = defaultdict(list)
    for name in raw_names:
        m = GRID_RC_PATTERN.match(name)
        if m:
            rc_groups[m.group("base")].append(name)

    # --- pass 1b: pre-check whether collapsing "x{n}" segments would collide ---
    # count how many raw variables would map to each collapsed name
    collapse_targets = defaultdict(list)
    for name in raw_names:
        m = X_LOOP_PATTERN.match(name)
        if m and m.group("suffix") and not X_COUNTRY_PATTERN.match(name):
            collapsed = f"{m.group('base')}{m.group('suffix')}"
            collapse_targets[collapsed].append(name)
    x_loop_collapse_is_safe = {k: len(v) for k, v in collapse_targets.items()}

    for name in raw_names:
        # -1. excluded variables (dropped from final output entirely)
        if is_excluded(name):
            results[name] = ("[EXCLUDED]", False, "matches exclude pattern — dropped from output")
            continue

        # 0. exact overrides (system + manual) take priority
        if name in DEFAULT_SYSTEM_FIELD_OVERRIDES:
            new_name, _ = DEFAULT_SYSTEM_FIELD_OVERRIDES[name]
            results[name] = (new_name, False, "system field dictionary")
            continue
        if name in DEFAULT_MANUAL_OVERRIDES:
            new_name, _ = DEFAULT_MANUAL_OVERRIDES[name]
            results[name] = (new_name, False, "manual override dictionary")
            continue

        # 1. grid "_r{n}_c1" pattern
        m = GRID_RC_PATTERN.match(name)
        if m:
            base, row = m.group("base"), m.group("row")
            siblings = rc_groups[base]
            if len(siblings) == 1:
                new_name = base  # singleton grid -> collapse, no row suffix
            else:
                new_name = f"{base}_{row}"
            results[name] = (new_name, False, "grid _r{n}_c pattern")
            continue

        # 2. country-loop pattern e.g. A5x1M / A5x1B (letter = country code)
        m = X_COUNTRY_PATTERN.match(name)
        if m and m.group("country") in COUNTRY_CODE_MAP:
            base, loop, country = m.group("base"), m.group("loop"), m.group("country")
            suffix = m.group("suffix") or ""
            country_name = COUNTRY_CODE_MAP[country]
            new_name = f"{base}.{loop}{suffix}.{country_name}"
            results[name] = (new_name, True, "country-loop pattern — VERIFY dot-notation convention")
            continue

        # 3. plain x-loop pattern e.g. B4ax1_1 -> B4a_1
        #    SAFETY: only collapse the "x{n}" segment if doing so is unique across
        #    the whole file. If multiple x-pages repeat the same item number
        #    (e.g. B4ax1_1 AND B4ax2_1 both exist, with DIFFERENT piped content),
        #    collapsing would silently collide two different variables onto one
        #    name — so keep the loop number in the name instead and flag it.
        m = X_LOOP_PATTERN.match(name)
        if m and m.group("suffix"):
            base, loop = m.group("base"), m.group("loop")
            collapsed = f"{base}{m.group('suffix')}"
            if x_loop_collapse_is_safe.get(collapsed) == 1:
                results[name] = (collapsed, True, "x-loop pattern — VERIFY convention matches this question type")
            else:
                # not safe to collapse: keep loop segment to guarantee uniqueness
                safe_name = f"{base}_x{loop}{m.group('suffix')}"
                results[name] = (safe_name, True,
                                  "x-loop pattern COLLIDES across pages (same item # reused with different "
                                  "content) — kept loop number in name to avoid overwriting data. "
                                  "Confirm your team's real naming convention for this case.")
            continue

        # 4. column-coded country grid e.g. C6B_c1 -> C6.Bangladesh_PMI
        m = COLUMN_CODE_PATTERN.match(name)
        if m and m.group("country") in COUNTRY_CODE_MAP:
            base, country, col = m.group("base"), m.group("country"), m.group("col")
            topic = COLUMN_TOPIC_MAP.get((base, col))
            if topic:
                new_name = f"{base}.{COUNTRY_CODE_MAP[country]}_{topic}"
                results[name] = (new_name, False, "column-topic dictionary")
            else:
                new_name = f"{base}.{COUNTRY_CODE_MAP[country]}_{col}"
                results[name] = (new_name, True, f"column code '{col}' not in topic dictionary — add mapping")
            continue

        # 5. trailing bare country-letter suffix e.g. B6aB / B6aM
        m = COUNTRY_LETTER_SUFFIX_PATTERN.match(name)
        if m and m.group("country") in COUNTRY_CODE_MAP and len(name) > 2:
            base, country = m.group("base"), m.group("country")
            new_name = f"{base}.{COUNTRY_CODE_MAP[country]}"
            results[name] = (new_name, True, "trailing country-letter — VERIFY not a false positive")
            continue

        # default: unchanged
        results[name] = (name, False, "no rename rule matched — kept as-is")

    # --- FINAL SAFETY NET: never let two different raw variables map to the
    # same new name. If any rule produced a collision, fall back to keeping
    # the raw name for every variable in that collision group and flag it
    # loudly. This must never be silently skipped. ---
    new_name_owners = defaultdict(list)
    for raw, (new_name, _, _) in results.items():
        if new_name != "[EXCLUDED]":
            new_name_owners[new_name].append(raw)

    for new_name, owners in new_name_owners.items():
        if len(owners) > 1:
            for raw in owners:
                results[raw] = (
                    raw, True,
                    f"COLLISION BLOCKED: rule would have renamed {len(owners)} different variables "
                    f"({', '.join(owners[:5])}{'...' if len(owners) > 5 else ''}) to '{new_name}'. "
                    f"Reverted to raw name — resolve manually before running syntax."
                )

    return results


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def load_codebook_appends(path):
    """
    Optional supplementary file: two columns [Variable, Option Text].
    If provided, per-option text is appended to the cleaned stem label as
    ': {Option Text}', matching your team's existing convention for grid items.
    """
    if not path:
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            out[str(row[0]).strip()] = str(row[1]).strip() if row[1] else ""
    return out


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input_sav")
    parser.add_argument("output_xlsx")
    parser.add_argument("--codebook", default=None,
                         help="Optional xlsx with [Variable, Option Text] to append per-item text "
                              "that Lighthouse's SPSS export doesn't include for grid/multi-select items.")
    args = parser.parse_args()

    df, meta = pyreadstat.read_sav(args.input_sav)
    raw_names = list(meta.column_names)
    labels = meta.column_names_to_labels
    value_labels = meta.variable_value_labels
    codebook_appends = load_codebook_appends(args.codebook)

    renames = suggest_renames(raw_names)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mapping"

    headers = [
        "Raw Variable", "Suggested Rename", "Rename Needs Review", "Rename Reason",
        "Raw Variable Label", "Cleaned Variable Label", "Label Needs Review", "Label Reason",
        "Has Value Labels", "Value Labels (preview)"
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    review_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    n_review = 0
    n_excluded = 0
    for name in raw_names:
        new_name, rn_review, rn_reason = renames[name]

        if new_name == "[EXCLUDED]":
            n_excluded += 1
            ws.append([name, "[EXCLUDED]", "", rn_reason, "", "", "", "", "", ""])
            for col in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=col).fill = PatternFill(
                    start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            continue

        raw_label = labels.get(name, "")
        clean, lbl_review, lbl_reason = clean_label(raw_label)

        # append per-option text from optional codebook, if this variable is a
        # grid/multi-select item and we have supplementary option text for it
        if name in codebook_appends and codebook_appends[name]:
            clean = f"{clean} : {codebook_appends[name]}"
        elif lbl_reason == "" and re.search(r"_r\d+_c\d+$|_\d+$", name) and ":" not in clean:
            # looks like a grid/multi item but we have no option text source
            lbl_review = True
            lbl_reason = "grid item — per-option text not in .sav; provide --codebook or add manually"

        vlabels = value_labels.get(name)
        has_vlabels = "Yes" if vlabels else "No"
        vlabel_preview = "; ".join(f"{k}={v}" for k, v in list(vlabels.items())[:4]) if vlabels else ""
        if vlabels and len(vlabels) > 4:
            vlabel_preview += " ..."

        row_idx = ws.max_row + 1
        ws.append([
            name, new_name, "REVIEW" if rn_review else "", rn_reason,
            raw_label, clean, "REVIEW" if lbl_review else "", lbl_reason,
            has_vlabels, vlabel_preview
        ])
        if rn_review or lbl_review:
            n_review += 1
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = review_fill

    for i, w in enumerate([20, 24, 14, 30, 45, 45, 14, 30, 12, 40], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    wb.save(args.output_xlsx)
    kept = len(raw_names) - n_excluded
    print(f"Wrote {args.output_xlsx}")
    print(f"Total variables in .sav: {len(raw_names)}")
    print(f"Excluded (dropped, e.g. forB4a/sys_pagetime): {n_excluded}")
    print(f"Kept for mapping: {kept}")
    print(f"Flagged for review: {n_review} ({n_review/kept*100:.1f}% of kept)")
    print(f"Auto-resolved cleanly: {kept - n_review} ({(kept-n_review)/kept*100:.1f}% of kept)")


if __name__ == "__main__":
    main()
