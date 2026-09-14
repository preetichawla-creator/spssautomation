#!/usr/bin/env python3
"""
generate_mapping.py — Draft SPSS mapping file generator (Phase 1 of the
Word questionnaire + raw .sav -> mapping file -> SPSS syntax automation).

USAGE
    python generate_mapping.py --sav RawData.sav --qnr questionnaire.docx --out mapping.xlsx

OUTPUT
    An .xlsx with two sheets:
      - "Variable Label": Raw Variable | Renamed Variable | Label | Notes
      - "Value Label":    Variable (renamed) | Code | Label

WHAT THIS DOES AUTOMATICALLY (validated against two real studies)
    - Loop notation:      QM1x1_r1        -> QM1.1
    - Grid collapsing:    Q13_r1_c1..r8_c1 -> Q13_1..Q13_8   (constant axis dropped)
                           D1_r1_c1        -> D1_1_1          (both axes vary -> row_col, default)
    - Single-mention drop: QM1x1_r1 (sole member of its stem) -> QM1.1 (no trailing index)
    - Flat positional special-code lookup: QC4_10 -> QC4_98   (code read live from the
      questionnaire's own table, not hardcoded)
    - A small maintained SYS_VAR_MAP for platform system fields (sys_RespNum -> Resp_Num, etc.)

WHAT THIS DOES **NOT** DO, ON PURPOSE (flagged in the "Notes" column instead of guessed)
    - Variable SELECTION: nothing is dropped. Every raw variable is kept and renamed
      where a rule applies; your team removes what it doesn't want.
    - Study-specific semantic renames (e.g. RespDetails_r1_c1 -> Respondent_Name): these
      carry no information in the raw file or questionnaire that says what they mean,
      so they pass through unchanged and get flagged "no rule matched" in Notes.
    - Grid row/col ORDER exceptions: when both row and column vary, the default is
      "row-then-column" (Q13_r1_c2 -> Q13_1_2). Any question that uses the opposite
      convention (or flattens a catch-all row out of the grid entirely, as QM8 does in
      the McDonald's study) is flagged "ambiguous grid order" and needs a manual check.
    - Codeless option tables where the true code is NOT sequential: if a questionnaire
      table has no printed code column, this script assumes row order == code order.
      That is wrong for a table that both omits codes AND skips ahead to a
      reserved code (e.g. "Any Other" -> 99) without ever printing it. There's nothing
      in the document to detect this from, so it is a genuine blind spot, not a bug.

Everything above was learned and validated across the McDonald's bakery study and a
second seller-survey study — see chat history for the worked examples.
"""
import argparse
import re
from collections import defaultdict

import pyreadstat
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook

# ----------------------------------------------------------------------------
# Maintained lookup for platform system fields. These don't vary by study, so
# extend this dict once and reuse it rather than re-discovering it per study.
# ----------------------------------------------------------------------------
SYS_VAR_MAP = {
    'sys_respnum': 'Resp_Num',
    'sys_starttime': 'Start_Time',
    'sys_endtime': 'End_Time',
    'sys_elapsedtime': 'Elapsed_Time',
}

STEM_LINE_RE = re.compile(r'^\**\s*(Q?[A-Za-z]{1,4}\d{1,3}[a-zA-Z]*)\s*[\.\-–:/\s]*(.*)$')


# ----------------------------------------------------------------------------
# 1. Questionnaire parsing: build, per question stem, an ordered codebook
#    [(option_label, code), ...] plus the question's own prompt text.
# ----------------------------------------------------------------------------
def iter_block_items(parent):
    for child in parent.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)


def find_stem_and_text(text):
    for line in text.split('\n'):
        m = STEM_LINE_RE.match(line.strip())
        if m:
            return m.group(1), m.group(2).strip()
    return None, None


def is_intlike(s):
    return bool(re.fullmatch(r'-?\d+', s.strip()))


def cell_is_bold(cell):
    runs = [r for p in cell.paragraphs for r in p.runs if r.text.strip()]
    if not runs:
        return False
    return all(bool(r.bold) for r in runs)


def classify_row(row):
    cells_text = [c.text.strip() for c in row.cells]
    uniq = list(dict.fromkeys(c for c in cells_text if c != ''))
    ints = [c for c in uniq if is_intlike(c)]
    nonints = [c for c in uniq if c and not is_intlike(c)]

    for c in cells_text:
        stem, text = find_stem_and_text(c)
        if stem:
            return ('header', (stem, text))

    if ints and nonints:
        return ('coded', (nonints[0], int(ints[-1])))
    if nonints and not ints:
        all_bold = all(cell_is_bold(c) for c in row.cells if c.text.strip())
        return ('caption' if all_bold else 'textonly', nonints[0])
    return ('skip', None)


def flush(questions, stem, buf):
    if not stem or not buf:
        return
    kinds = set(k for k, _ in buf)
    if 'coded' in kinds:
        for k, v in buf:
            if k == 'coded':
                questions.setdefault(stem, []).append(v)
    elif kinds <= {'textonly', 'caption'} and any(k == 'textonly' for k, _ in buf):
        labels = [v for k, v in buf if k == 'textonly']
        for i, label in enumerate(labels, start=1):
            questions.setdefault(stem, []).append((label, i))


def parse_questionnaire(path):
    """Returns (questions, question_text):
       questions[stem]      = [(option_label, code), ...] in document order
       question_text[stem]  = the question's own prompt text (first line, cleaned)
    """
    doc = Document(path)
    items = list(iter_block_items(doc.element.body))
    questions, question_text = {}, {}
    current_stem, buf = None, []

    def set_stem(stem, text):
        nonlocal current_stem, buf
        if stem != current_stem:
            flush(questions, current_stem, buf)
            buf = []
            current_stem = stem
        if text and stem not in question_text:
            question_text[stem] = text

    for it in items:
        if isinstance(it, Paragraph):
            text = it.text.strip()
            if text:
                stem, rest = find_stem_and_text(text)
                if stem:
                    set_stem(stem, rest)
        else:
            for row in it.rows:
                kind, payload = classify_row(row)
                if kind == 'header':
                    set_stem(*payload)
                elif kind in ('coded', 'textonly', 'caption'):
                    buf.append((kind, payload))
    flush(questions, current_stem, buf)
    return questions, question_text


# ----------------------------------------------------------------------------
# 2. Raw variable name parsing: decompose into stem / loop / row / col / kind
# ----------------------------------------------------------------------------
def parse_raw(name):
    other = name.endswith('_other')
    base = name[:-6] if other else name

    m = re.match(r'^(.+)_r(\d+)_c(\d+)$', base)
    if m:
        stem_loop, row, col, kind = m.group(1), int(m.group(2)), int(m.group(3)), 'rc'
    else:
        m = re.match(r'^(.+)_r(\d+)$', base)
        if m:
            stem_loop, row, col, kind = m.group(1), int(m.group(2)), None, 'r'
        else:
            m = re.match(r'^(.+)_c(\d+)$', base)
            if m:
                stem_loop, row, col, kind = m.group(1), None, int(m.group(2)), 'c'
            else:
                m = re.match(r'^(.+)_(\d+)$', base)
                if m:
                    stem_loop, row, col, kind = m.group(1), None, int(m.group(2)), 'flat'
                else:
                    stem_loop, row, col, kind = base, None, None, 'plain'

    m2 = re.match(r'^(.+?)x(\d+)$', stem_loop)
    stem, loop = (m2.group(1), int(m2.group(2))) if m2 else (stem_loop, None)
    return dict(stem=stem, loop=loop, kind=kind, row=row, col=col, other=other, raw=name)


def build_family_rules(parsed_list):
    """Per (stem, loop) family: total size (to catch single-mention items whose
    index carries no information), which r/c/rc kinds appear together (a mix,
    e.g. a standalone 'c' plus an 'rc' grid under the same stem, signals a
    composite/multi-part question rather than one simple grid), and — for
    members with BOTH row and col — whether one axis is constant across the
    family and can be dropped."""
    fam_size = defaultdict(int)
    fam_kinds = defaultdict(set)
    rc_rows, rc_cols = defaultdict(set), defaultdict(set)
    for p in parsed_list:
        if p['kind'] in ('r', 'c', 'rc', 'flat'):
            key = (p['stem'], p['loop'])
            fam_size[key] += 1
            fam_kinds[key].add(p['kind'])
            if p['kind'] == 'rc':
                rc_rows[key].add(p['row'])
                rc_cols[key].add(p['col'])
    rc_rule = {}
    for key in set(rc_rows) | set(rc_cols):
        if len(rc_cols[key]) <= 1:
            rc_rule[key] = 'row_only'
        elif len(rc_rows[key]) <= 1:
            rc_rule[key] = 'col_only'
        else:
            rc_rule[key] = 'both_row_first'
    composite = {key for key, kinds in fam_kinds.items() if len(kinds) > 1}
    return dict(fam_size=fam_size, rc_rule=rc_rule, composite=composite)


# ----------------------------------------------------------------------------
# 3. Rename + label + notes for one variable
# ----------------------------------------------------------------------------
def item_position(p):
    """Which number identifies 'which item in the list' for a positional lookup —
    the flat suffix, or whichever of row/col varies for a grid item."""
    if p['kind'] == 'flat':
        return p['col']
    if p['kind'] == 'r':
        return p['row']
    if p['kind'] == 'c':
        return p['col']
    return None  # 'rc' with both varying has no single "position" -> no lookup


def rename_and_label(p, family_rules, questions, question_text, sav_label):
    notes = []

    if p['raw'].lower() in SYS_VAR_MAP:
        new_name = SYS_VAR_MAP[p['raw'].lower()]
        label = sav_label or new_name.replace('_', ' ')
        return new_name, label, ''

    new_stem = p['stem'] + (f".{p['loop']}" if p['loop'] else "")
    key = (p['stem'], p['loop'])
    qcodes_pairs = questions.get(p['stem'])
    parent_text = question_text.get(p['stem'], '')

    if p['kind'] == 'plain':
        new_name = new_stem + ('_other' if p['other'] else '')
        label = parent_text or sav_label or new_name
        if not qcodes_pairs and not parent_text:
            notes.append('no rule matched (kept as raw) — likely needs a study-specific name')
        return new_name, label, '; '.join(notes)

    fam_size = family_rules['fam_size'].get(key, 1)
    pos = item_position(p)

    if key in family_rules['composite']:
        notes.append('composite/multi-part question (mixed grid shapes under one stem) — '
                      'rename/label may not match your team\'s convention here, please verify')

    # --- suffix / rename ---
    if p['kind'] == 'flat':
        if qcodes_pairs and pos <= len(qcodes_pairs):
            code = qcodes_pairs[pos - 1][1]
            suffix = f"_{code}"
        else:
            suffix = f"_{pos}"
            notes.append('no questionnaire match for this position — code left as-is, please verify')
    elif fam_size == 1:
        suffix = ''  # sole member of its stem -> index carries no information
    elif p['kind'] == 'r':
        suffix = f"_{p['row']}"
    elif p['kind'] == 'c':
        suffix = f"_{p['col']}"
    else:  # 'rc'
        rule = family_rules['rc_rule'].get(key, 'both_row_first')
        if rule == 'row_only':
            suffix = f"_{p['row']}"
        elif rule == 'col_only':
            suffix = f"_{p['col']}"
        else:
            suffix = f"_{p['row']}_{p['col']}"
            notes.append('ambiguous grid order (both row & col vary) — used default row-then-column, please verify')

    new_name = new_stem + suffix + ('_other' if p['other'] else '')

    # --- label ---
    item_label = None
    if qcodes_pairs and pos is not None and pos <= len(qcodes_pairs):
        item_label = qcodes_pairs[pos - 1][0]
    elif qcodes_pairs is None and fam_size > 1:
        notes.append('question stem not found in questionnaire — label may be incomplete')

    if parent_text and item_label and fam_size > 1:
        label = f"{parent_text} :: {item_label}"
    elif parent_text:
        label = parent_text
    elif item_label:
        label = item_label
    else:
        label = sav_label or new_name

    if p['other']:
        label = f"{label} :: Others"

    return new_name, label, '; '.join(notes)


# ----------------------------------------------------------------------------
# 4. Main: read .sav, parse questionnaire, build both sheets
# ----------------------------------------------------------------------------
def clean_sav_label(varname, raw_label):
    if not raw_label:
        return ''
    lbl = re.sub(r'^' + re.escape(varname) + r'\s*-\s*', '', raw_label)
    lbl = re.sub(r'\[%.*?%\]', '', lbl).strip()
    return lbl


def build_workbook(sav_path, qnr_path):
    df, meta = pyreadstat.read_sav(sav_path, metadataonly=True)
    raw_cols = meta.column_names
    sav_labels = meta.column_names_to_labels
    sav_value_labels = meta.variable_value_labels

    questions, question_text = parse_questionnaire(qnr_path)

    parsed_all = [parse_raw(c) for c in raw_cols]
    family_rules = build_family_rules(parsed_all)

    rows = []
    rename_map = {}
    for p in parsed_all:
        sav_lbl = clean_sav_label(p['raw'], sav_labels.get(p['raw']))
        new_name, label, notes = rename_and_label(p, family_rules, questions, question_text, sav_lbl)
        rows.append((p['raw'], new_name, label, notes))
        rename_map[p['raw']] = new_name

    wb = Workbook()

    ws1 = wb.active
    ws1.title = 'Variable Label'
    ws1.append(['Variable Information'])
    ws1.append(['Raw Variable', 'Renamed Variable', 'Label', 'Notes'])
    for r in rows:
        ws1.append(list(r))

    ws2 = wb.create_sheet('Value Label')
    ws2.append(['Variable Values'])
    ws2.append(['Variable', 'Code', 'Label'])
    for raw_name in raw_cols:
        vl = sav_value_labels.get(raw_name)
        if not vl:
            continue
        new_name = rename_map[raw_name]
        first = True
        for code, label in vl.items():
            ws2.append([new_name if first else None, code, label])
            first = False

    return wb


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--sav', required=True, help='Path to the raw .sav data file')
    ap.add_argument('--qnr', required=True, help='Path to the Word questionnaire (.docx)')
    ap.add_argument('--out', required=True, help='Path to write the draft mapping .xlsx')
    args = ap.parse_args()

    wb = build_workbook(args.sav, args.qnr)
    wb.save(args.out)
    print(f'Wrote {args.out}')


if __name__ == '__main__':
    main()
