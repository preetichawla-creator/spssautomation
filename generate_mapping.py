#!/usr/bin/env python3
"""
generate_mapping.py — Draft SPSS mapping file generator (Phase 1 of the
Word questionnaire + raw .sav -> mapping file -> SPSS syntax automation).

USAGE
    python generate_mapping.py --sav RawData.sav --qnr questionnaire.docx --out mapping.xlsx

OUTPUT
    An .xlsx with one sheet:
      - "Variable Label": Raw Variable | Renamed Variable | Label | Notes
    (A "Value Label" sheet is intentionally not generated for now -- a different
    approach for value labels is being worked out separately.)

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
import zipfile
from collections import defaultdict

import pyreadstat
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import PatternFill

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

# Unicode blocks for non-Latin scripts commonly used for in-questionnaire translations
# (Hindi/Devanagari, Urdu/Arabic, and other major Indic scripts). Text in these ranges
# is stripped out of labels — SPSS variable labels should carry the English question
# text only, not a bundled translation.
NON_LATIN_RE = re.compile(
    r'[\u0900-\u097F\u0980-\u09FF\u0A00-\u0A7F\u0A80-\u0AFF\u0B00-\u0B7F'
    r'\u0B80-\u0BFF\u0C00-\u0C7F\u0C80-\u0CFF\u0D00-\u0D7F'
    r'\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]+'
)

LABEL_CHAR_LIMIT = 250

# A single item's own text is capped well under the label limit -- an item that's a full
# paragraph (e.g. "do you agree with this statement: <200 words>") would otherwise starve
# the parent question of any room at all once " : {item}" is appended. Leaves a comfortable
# floor for the parent text even in the worst case.
MAX_ITEM_TEXT_CHARS = 120

# Survey-programming/interviewer instructions that carry no respondent-facing meaning.
# Bracketed/parenthesized instructions ("[If A1=1]", "(Please select only one answer)",
# "(SR)"/"(MR)" tags) are removed wholesale by the [...] and (...) strips in clean_text
# below, per explicit instruction to drop all bracket/paren content from labels. This
# pattern only needs to catch the same kind of tag when it ISN'T bracket-wrapped.
INSTRUCTION_PHRASES_RE = re.compile(
    r'\b(single\s*[- ]?\s*(answer|code|coding|response)'
    r'|multi(ple)?\s*[- ]?\s*(option|response|coding|answer|code)s?'
    r'|check\s*box'
    r'|#?options?\s+[\d\-–\s]*(are\s+)?(randomly\s+)?(distributed|shown\s+in\s+(random|fixed)\s+order))\b',
    re.IGNORECASE
)


def clean_text(text):
    """Strip survey-programming markup (<...>, [...], (...)), interviewer/routing
    instructions, leading bullet/checkbox glyphs, blank-fill underscore runs, and any
    non-Latin-script translation text, then collapse whatever whitespace/punctuation
    that leaves behind. Deliberately keeps '?' intact -- callers comparing candidate
    question text (see maybe_update_question_text) rely on it; '?' -> '.' is applied
    once, at final label assembly, by finalize() in rename_and_label."""
    if not text:
        return ''
    # normalize fullwidth CJK punctuation to its ASCII equivalent -- this document mixes
    # both inconsistently (e.g. "Gender：" vs "Gender:"), and downstream regexes here are
    # all written against ASCII punctuation
    cleaned = text.translate(str.maketrans('（）：；，？', '():;,?'))
    cleaned = cleaned.translate(str.maketrans('', '', '"“”„‟'))   # drop all double-quote variants
    cleaned = re.sub(r'<[^>]*>', '', cleaned)
    cleaned = re.sub(r'\[[^\]]*\]', '', cleaned)   # covers [%...%] too, being a superset
    cleaned = re.sub(r'\([^)]*\)', '', cleaned)    # parenthetical content is never wanted in a label
    cleaned = INSTRUCTION_PHRASES_RE.sub('', cleaned)
    cleaned = NON_LATIN_RE.sub('', cleaned)
    cleaned = re.sub(r'_{3,}', '', cleaned)        # blank-fill-in underscore runs, e.g. "First____" -> "First"
    # leading UI marker glyphs (radio-button/checkbox bullets, a stray '#' programmer-note
    # marker, or a "(1)"/"(2)" style item-position marker) carry no content of their own.
    cleaned = re.sub(r'^[○□●■◯•]+\s*', '', cleaned)
    cleaned = re.sub(r'^\(\d+\)\s*', '', cleaned)
    cleaned = re.sub(r'^#+\s*', '', cleaned)
    cleaned = re.sub(r'[\(\[]\s*[\)\]]', '', cleaned)          # empty () or [] left behind
    cleaned = re.sub(r'\s{2,}', ' ', cleaned).strip()
    # drop trailing whitespace-separated tokens that are pure punctuation — these are
    # debris left behind where a translation used to sit (e.g. "gender? ?" -> "gender?")
    tokens = cleaned.split(' ')
    while tokens and re.fullmatch(r'[^\w]+', tokens[-1] or ''):
        tokens.pop()
    cleaned = ' '.join(tokens)
    return cleaned.strip(' /|-,')


def truncate_label(text, limit=LABEL_CHAR_LIMIT):
    """Hard cap at `limit` chars, cutting at the last full word so SPSS never
    receives a label over its variable-label limit. Returns (text, was_truncated)."""
    if not text or len(text) <= limit:
        return text, False
    truncated = text[:limit]
    if ' ' in truncated:
        truncated = truncated.rsplit(' ', 1)[0]
    return truncated.rstrip(' ,.;:-'), True


def compose_with_suffix(base, suffix, limit=LABEL_CHAR_LIMIT):
    """Join base+suffix, shortening `base` (not the suffix) if needed so the
    item-specific suffix — the part that actually distinguishes this variable
    from its siblings — always survives, rather than being the first thing lost
    to truncation. When the suffix is our " : "-style separator, a trailing "?"
    or "." on base is dropped first -- otherwise "...year? : First" turns into
    the doubled-up "...year. : First" once "?" becomes "." at final assembly."""
    if suffix.lstrip().startswith(':'):
        base = base.rstrip().rstrip('?.').rstrip()
    if len(base) + len(suffix) <= limit:
        return base + suffix
    room = limit - len(suffix)
    if room < 20:
        # suffix alone is nearly/over the limit -- nothing sensible to reserve for
        # base, so fall back to plain truncation of the combined string
        return base + suffix
    shortened, _ = truncate_label(base, room)
    return shortened + suffix


def load_numbering_map(docx_path):
    """Word can auto-number headings via a List style (numPr) instead of the number
    being literal text in the paragraph -- python-docx doesn't resolve these, so this
    reads word/numbering.xml directly. Returns {numId: (prefix, start)} for every
    simple single-level, decimal, "<prefix>%1<suffix>" numbering definition (e.g.
    lvlText "A%1." -> prefix "A"), which is what auto-numbered question stems like
    "A3.", "C5." turn out to be in practice. Anything more complex (multi-level,
    non-decimal, roman numerals, etc.) is left alone rather than guessed at."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            if 'word/numbering.xml' not in z.namelist():
                return {}
            content = z.read('word/numbering.xml').decode('utf-8', errors='ignore')
    except Exception:
        return {}

    abstract_fmt = {}
    for m in re.finditer(r'<w:abstractNum w:abstractNumId="(\d+)".*?</w:abstractNum>', content, re.S):
        aid, block = m.group(1), m.group(0)
        lvl0 = re.search(r'<w:lvl w:ilvl="0">.*?</w:lvl>', block, re.S)
        if not lvl0:
            continue
        lvltext_m = re.search(r'w:lvlText w:val="([^"]*)"', lvl0.group(0))
        start_m = re.search(r'w:start w:val="(\d+)"', lvl0.group(0))
        fmt_m = re.search(r'w:numFmt w:val="([^"]*)"', lvl0.group(0))
        if not lvltext_m or not fmt_m or fmt_m.group(1) != 'decimal':
            continue
        prefix_m = re.match(r'^([^%]*)%1[^%]*$', lvltext_m.group(1))
        if not prefix_m or not prefix_m.group(1):
            continue  # no literal prefix before the counter -> not a "letter+number" stem
        abstract_fmt[aid] = (prefix_m.group(1), int(start_m.group(1)) if start_m else 1)

    num_map = {}
    for m in re.finditer(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"/>', content):
        num_id, abs_id = m.group(1), m.group(2)
        if abs_id in abstract_fmt:
            num_map[num_id] = abstract_fmt[abs_id]
    return num_map


def resolve_auto_number(paragraph, numbering_map, counters):
    """If this paragraph is auto-numbered by one of the simple letter-series lists
    load_numbering_map found, return (stem, was_new) advancing that list's counter;
    otherwise None. ilvl must be 0 -- sub-levels of a multi-level list aren't a
    single letter+number stem and aren't handled here."""
    pPr = paragraph._p.pPr
    if pPr is None or pPr.numPr is None or pPr.numPr.numId is None:
        return None
    ilvl = pPr.numPr.ilvl.val if pPr.numPr.ilvl is not None else 0
    if ilvl != 0:
        return None
    num_id = str(pPr.numPr.numId.val)
    if num_id not in numbering_map:
        return None
    prefix, start = numbering_map[num_id]
    counters[num_id] = counters.get(num_id, start - 1) + 1
    return f"{prefix}{counters[num_id]}"


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


# A question numbered as a bare "1.", "2." with no letter prefix at all in the document
# text (some platforms number questions this way in the doc while the actual raw
# variable is "Q1", "Q2", etc. -- the "Q" only exists in the data, never in the doc).
# Tagged with a "#" sentinel here and resolved against the real raw variable stems in
# parse_questionnaire, rather than guessed blindly -- see resolve_sentinel_stem.
BARE_NUMBER_RE = re.compile(r'^\(?\s*(\d{1,3})\s*[\.\)](?:\s+(?=\S)|(?=[A-Z(]))(.+)$')


def find_bare_number_stem(text):
    for line in text.split('\n'):
        m = BARE_NUMBER_RE.match(line.strip())
        if m:
            return f"#{m.group(1)}", m.group(2).strip()
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

    # Word can merge a cell across many underlying grid columns, and python-docx then
    # reports that cell's text once per column it spans -- so a row can show up with
    # anywhere from 3 to 30+ raw cells for what is logically a 2- or 8-column row, and
    # even the lead "item" cell itself may repeat 1-3x before the real content starts.
    # Dedup (uniq, already order-preserving) is the only reliable read on row shape;
    # only whether the ORIGINAL first cell was blank still needs the raw list.
    raw_first_blank = not cells_text[0].strip() if cells_text else True
    if raw_first_blank:
        item_label, rest = None, uniq
    else:
        item_label = uniq[0] if uniq else ''
        rest = uniq[1:]

    # A bare programmer note as the row's lead cell (e.g. "#Randomize items") is never
    # a real item or a real header -- skip the row outright.
    if item_label and item_label.startswith('#'):
        return ('skip', None)

    # A rating-grid ITEM row: the lead cell is the item/statement text, and the rest of
    # the row is either the scale's real codes (printed once, typically on the row that
    # doubles as the legend -- several distinct ints) or an identical placeholder glyph
    # repeated per column (e.g. "○", collapsed by dedup to one remaining value). Either
    # way this is ONE item, positioned by row order -- not a single (label, code) pair
    # the way a plain coded option row ("Male | 1") is. Checked BEFORE the header/legend
    # skip below: an item + one placeholder also has zero ints and 2 "distinct nonints"
    # (the item text and the placeholder itself), which would otherwise look identical
    # to a real header row's shape.
    if item_label and not is_intlike(item_label) and rest:
        if len(rest) > 1 and all(is_intlike(c) for c in rest):
            return ('textonly', item_label)
        if len(rest) == 1 and not is_intlike(rest[0]):
            return ('textonly', item_label)

    # A pure column-header/legend row for a rating grid: several DIFFERENT scale-label
    # words, and NO numeric code anywhere in the row at all (e.g. "| | Very good | Good |
    # ... |", or "| LISTING & CATALOGING PROCESS | Very Poor | Poor | ... |" where the
    # lead cell is a section title rather than blank). Deliberately checks the row's
    # ints/nonints as a whole, not just "everything after the first cell" -- some tables
    # put the CODE first instead of the label ("| 1 | Advertising / Media | TERMINATE |"),
    # which must still fall through to the normal 'coded' handling below, not be mistaken
    # for a headerless legend row.
    if not ints and len(nonints) > 1:
        return ('skip', None)

    if ints and nonints:
        return ('coded', (nonints[0], int(ints[-1])))
    if nonints and not ints:
        # Bare-number stems are only checked for a genuine single-cell header row (the
        # whole row is one spanning cell) -- a real answer-option row like "(1)First___"
        # paired with a code is already caught by the 'coded' branch above, so this
        # can't be confused with one.
        if len(uniq) == 1:
            stem, text = find_bare_number_stem(uniq[0])
            if stem:
                return ('header', (stem, text))
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
                label, code = v
                questions.setdefault(stem, []).append((clean_text(label), code))
    elif kinds <= {'textonly', 'caption'} and any(k == 'textonly' for k, _ in buf):
        labels = [v for k, v in buf if k == 'textonly']
        for i, label in enumerate(labels, start=1):
            questions.setdefault(stem, []).append((clean_text(label), i))


def parse_questionnaire(path, raw_stems=frozenset()):
    """Returns (questions, question_text):
       questions[stem]      = [(option_label, code), ...] in document order
       question_text[stem]  = the question's own prompt text
    raw_stems is the set of actual variable stems seen in the raw .sav -- used to
    resolve bare-number headers ("1.", "2.") to their real stem ("Q1", "Q2") only
    when that stem is confirmed to exist, rather than guessed.
    """
    doc = Document(path)
    items = list(iter_block_items(doc.element.body))
    numbering_map = load_numbering_map(path)
    numbering_counters = {}
    questions, question_text = {}, {}
    current_stem, buf = None, []

    def resolve_sentinel(stem):
        if not stem or not stem.startswith('#'):
            return stem
        n = stem[1:]
        if f"Q{n}" in raw_stems:
            return f"Q{n}"
        if n in raw_stems:
            return n
        return None  # no confirmed match in the raw data -- don't guess

    def maybe_update_question_text(stem, raw_text):
        # Some questions open with an intro/transition sentence ("I will now ask you
        # some questions...") before the actual question appears in a later paragraph.
        # Prefer whichever candidate is actually phrased as a question (has a "?"),
        # and don't clobber a good one already found with a later non-question line.
        cleaned = clean_text(raw_text)
        if not cleaned:
            return
        existing = question_text.get(stem)
        if existing is None or ('?' not in existing and '?' in cleaned):
            question_text[stem] = cleaned

    def set_stem(stem, text):
        nonlocal current_stem, buf
        if stem != current_stem:
            flush(questions, current_stem, buf)
            buf = []
            current_stem = stem
        if text:
            maybe_update_question_text(stem, text)

    for it in items:
        if isinstance(it, Paragraph):
            text = it.text.strip()
            if text:
                stem, rest = find_stem_and_text(text)
                if not stem:
                    stem, rest = find_bare_number_stem(text)
                    stem = resolve_sentinel(stem)
                if not stem and numbering_map:
                    auto_stem = resolve_auto_number(it, numbering_map, numbering_counters)
                    if auto_stem:
                        stem, rest = auto_stem, text
                if stem:
                    set_stem(stem, rest)
                elif current_stem:
                    maybe_update_question_text(current_stem, text)
        else:
            for row in it.rows:
                kind, payload = classify_row(row)
                if kind == 'header':
                    stem, text = payload
                    stem = resolve_sentinel(stem)
                    if stem:
                        set_stem(stem, text)
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
def item_position(p, family_rules, key):
    """Which number identifies 'which item in the list' for a positional lookup —
    the flat suffix, whichever of row/col varies for a grid item, or — for an 'rc'
    item whose family collapses to a single varying axis (see build_family_rules) —
    that axis. A genuine two-axis grid (both vary) has no single "position", so no
    item-specific lookup is attempted there."""
    if p['kind'] == 'flat':
        return p['col']
    if p['kind'] == 'r':
        return p['row']
    if p['kind'] == 'c':
        return p['col']
    if p['kind'] == 'rc':
        rule = family_rules['rc_rule'].get(key)
        if rule == 'row_only':
            return p['row']
        if rule == 'col_only':
            return p['col']
    return None


def resolve_item_label(p, pos, key, fam_size, qcodes_pairs, stem_mismatch, sav_value_label_dict, other_positions):
    """Figure out the item-specific text (if any) for one grid/flat-family member,
    without composing or truncating anything -- shared by the family-wide reserve
    calculation (compute_family_reserve) and rename_and_label, so both agree on
    exactly what text a sibling will need.
    Returns (item_label_or_None, is_other_pair, note_or_None). is_other_pair means
    the caller should compose ' : Others' / ' : Others Specify' directly rather than
    use item_label (which is None in that case)."""
    if pos is not None and (key, pos) in other_positions:
        return None, True, None

    if p['kind'] == 'flat' and stem_mismatch:
        note = ("questionnaire item count for this question doesn't match the raw data "
                "— used the raw data's own value labels for the option text instead")
        item_label = None
        if sav_value_label_dict:
            raw_item = sav_value_label_dict.get(pos, sav_value_label_dict.get(float(pos)))
            item_label = clean_text(raw_item) if raw_item else None
        if not item_label:
            note += '; no matching option code in raw data value labels either — label may be incomplete'
    elif qcodes_pairs and pos is not None and pos <= len(qcodes_pairs):
        item_label, note = qcodes_pairs[pos - 1][0], None
    elif qcodes_pairs is None and fam_size > 1:
        item_label, note = None, 'question stem not found in questionnaire — label may be incomplete'
    else:
        item_label, note = None, None

    # An item's own text can itself be a full paragraph (e.g. a long statement to agree
    # or disagree with). Cap it well under the label limit so the parent question always
    # keeps a meaningful share of the budget too -- otherwise a single huge item starves
    # every sibling's parent text down to a near-meaningless fragment (see compose logic
    # in compute_family_reserve/rename_and_label, which size the parent around this cap).
    if item_label and len(item_label) > MAX_ITEM_TEXT_CHARS:
        item_label, _ = truncate_label(item_label, MAX_ITEM_TEXT_CHARS)
        note = (note + '; ' if note else '') + 'item text itself shortened to fit label limits'

    return item_label, False, note


def pick_meaningful_text(text, limit):
    """When the full text won't fit, prefer dropping whole LEADING sentences over a
    blind character cut -- the later sentence(s) in a questionnaire prompt are more
    often the actual question, with earlier ones being scene-setting context (e.g.
    "Up until <date>, X happened. Given that, would you...?" -- the second sentence
    alone is usually the meaningful, self-contained ask). Falls back to a plain
    character truncation only if no sentence-boundary cut fits."""
    if not text or len(text) <= limit:
        return text
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    for i in range(1, len(sentences)):
        candidate = ' '.join(sentences[i:]).strip()
        if candidate and len(candidate) <= limit:
            return candidate
    fallback, _ = truncate_label(text, limit)
    return fallback


def compute_family_reserve(parsed_all, family_rules, questions, flat_mismatch, other_positions, sav_value_labels):
    """For every stem where more than one raw variable shares it, find the longest
    ' : {item text}' (or ' : Others' / ' : Others Specify') suffix ANY sibling will
    need. Used to shorten that stem's shared parent text once, for the worst case --
    so every sibling in the same question keeps identical text before ' : ', rather
    than each independently truncating the parent to fit its own item's length."""
    reserve = defaultdict(int)
    for p in parsed_all:
        if p['kind'] == 'plain':
            continue
        key = (p['stem'], p['loop'])
        fam_size = family_rules['fam_size'].get(key, 1)
        if fam_size <= 1:
            continue
        pos = item_position(p, family_rules, key)
        qcodes_pairs = questions.get(p['stem'])
        stem_mismatch = p['kind'] == 'flat' and p['stem'] in flat_mismatch
        sav_value_label_dict = sav_value_labels.get(p['raw'])
        item_label, is_other_pair, _ = resolve_item_label(
            p, pos, key, fam_size, qcodes_pairs, stem_mismatch, sav_value_label_dict, other_positions)
        if is_other_pair:
            needed = len(" : Others Specify") if p['other'] else len(" : Others")
        elif item_label:
            needed = len(f" : {item_label}")
        else:
            continue
        reserve[p['stem']] = max(reserve[p['stem']], needed)
    return reserve


def rename_and_label(p, family_rules, questions, question_text, sav_label, flat_mismatch,
                      sav_value_label_dict, other_positions):
    notes = []
    recoded = False  # True when a flat position's raw suffix number was mapped to a
                      # DIFFERENT code (e.g. Q35_7 -> Q35_999) -- the team wants these
                      # highlighted since it's an actual value change, not just a rename

    def finalize(new_name, label):
        if label:
            label = label.replace('?', '.')
        label, was_truncated = truncate_label(label)
        if was_truncated:
            notes.append(f'label truncated to {LABEL_CHAR_LIMIT} characters — '
                          'original questionnaire text was longer, please review')
        return new_name, label, '; '.join(notes), recoded

    if p['raw'].lower() in SYS_VAR_MAP:
        new_name = SYS_VAR_MAP[p['raw'].lower()]
        label = sav_label or new_name.replace('_', ' ')
        return finalize(new_name, label)

    new_stem = p['stem'] + (f".{p['loop']}" if p['loop'] else "")
    key = (p['stem'], p['loop'])
    qcodes_pairs = questions.get(p['stem'])
    parent_text = question_text.get(p['stem'], '')

    if p['kind'] == 'plain':
        new_name = new_stem + ('_other' if p['other'] else '')
        label = parent_text or sav_label or new_name
        if not qcodes_pairs and not parent_text:
            notes.append('no rule matched (kept as raw) — likely needs a study-specific name')
        return finalize(new_name, label)

    fam_size = family_rules['fam_size'].get(key, 1)
    pos = item_position(p, family_rules, key)

    if key in family_rules['composite']:
        notes.append('composite/multi-part question (mixed grid shapes under one stem) — '
                      'rename/label may not match your team\'s convention here, please verify')

    stem_mismatch = p['kind'] == 'flat' and p['stem'] in flat_mismatch

    # --- suffix / rename ---
    if p['kind'] == 'flat':
        if qcodes_pairs and pos <= len(qcodes_pairs) and not stem_mismatch:
            code = qcodes_pairs[pos - 1][1]
            suffix = f"_{code}"
            recoded = (code != pos)
        else:
            # questionnaire count doesn't line up with the raw file for this stem (or the
            # stem wasn't found at all) -- the doc's row order can't be trusted to derive
            # the right code, so keep the raw suffix as-is rather than risk a wrong one
            suffix = f"_{pos}"
            if not stem_mismatch:
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

    item_label, is_other_pair, item_note = resolve_item_label(
        p, pos, key, fam_size, qcodes_pairs, stem_mismatch, sav_value_label_dict, other_positions)
    if item_note:
        notes.append(item_note)

    if is_other_pair:
        base_label = compose_with_suffix(parent_text, " : Others") if parent_text else "Others"
        label = compose_with_suffix(base_label, " Specify") if p['other'] else base_label
        return finalize(new_name, label)

    if parent_text and item_label and fam_size > 1:
        label = compose_with_suffix(parent_text, f" : {item_label}")
    elif parent_text:
        label = parent_text
    elif item_label:
        label = item_label
    else:
        label = sav_label or new_name

    if p['other']:
        # no detected non-other sibling at this position (rare -- e.g. a genuine two-axis
        # grid item) -- fall back to the old plain-append behavior rather than dropping it
        label = compose_with_suffix(label, " : Others")

    return finalize(new_name, label)


# ----------------------------------------------------------------------------
# 4. Main: read .sav, parse questionnaire, build both sheets
# ----------------------------------------------------------------------------
def clean_sav_label(varname, stem, raw_label):
    """Strip a leading 'varname - ' or 'stem - ' prefix off a raw .sav variable label.
    Multi-select family members (Q31_1, Q31_2, ...) commonly all share one label that's
    prefixed with the STEM ('Q31 - ...'), not the full variable name, so both are tried."""
    if not raw_label:
        return ''
    lbl = re.sub(r'^' + re.escape(varname) + r'\s*-\s*', '', raw_label)
    if lbl == raw_label and stem and stem != varname:
        lbl = re.sub(r'^' + re.escape(stem) + r'\s*-\s*', '', raw_label)
    return clean_text(lbl)


def build_workbook(sav_path, qnr_path):
    df, meta = pyreadstat.read_sav(sav_path, metadataonly=True)
    raw_cols = meta.column_names
    sav_labels = meta.column_names_to_labels
    sav_value_labels = meta.variable_value_labels

    parsed_all = [parse_raw(c) for c in raw_cols]
    raw_stems = {p['stem'] for p in parsed_all}
    questions, question_text = parse_questionnaire(qnr_path, raw_stems)

    family_rules = build_family_rules(parsed_all)

    # Detect stems where the questionnaire's option count doesn't match the number of
    # raw flat-family variables actually in the data -- positional matching can't be
    # trusted for these, so labeling (and renaming) falls back to the raw data's own
    # value labels instead (see rename_and_label).
    flat_positions = defaultdict(set)
    for p in parsed_all:
        if p['kind'] == 'flat':
            flat_positions[p['stem']].add(p['col'])
    # Mismatch = the questionnaire codebook can't cover every position actually used in
    # the raw file for this stem (missing entirely, or too short) -- NOT simply "different
    # counts", since a lone _other companion variable legitimately represents just one
    # high-numbered position out of a longer response list and that's not an error.
    flat_mismatch = {stem for stem, positions in flat_positions.items()
                      if len(questions.get(stem, [])) < max(positions)}

    # Positions that have a "please specify" write-in companion (X_N + X_N_other) --
    # these are always the catch-all "Others" option; see rename_and_label.
    other_positions = set()
    for p in parsed_all:
        if p['other']:
            key = (p['stem'], p['loop'])
            pos = item_position(p, family_rules, key)
            if pos is not None:
                other_positions.add((key, pos))

    # Resolve each multi-member stem's parent text ONCE, sized to fit alongside the
    # longest item any of its siblings will need -- so every sibling shares identical
    # text before " : ", instead of each independently truncating the parent based on
    # its own (shorter or longer) item text.
    reserve = compute_family_reserve(parsed_all, family_rules, questions, flat_mismatch,
                                      other_positions, sav_value_labels)
    for stem, needed in reserve.items():
        if stem in question_text:
            budget = max(LABEL_CHAR_LIMIT - needed, 20)
            question_text[stem] = pick_meaningful_text(question_text[stem], budget)

    rows = []
    rename_map = {}
    recoded_rows = set()   # row indices (0-based within `rows`) where a code was recoded
    for p in parsed_all:
        sav_lbl = clean_sav_label(p['raw'], p['stem'], sav_labels.get(p['raw']))
        sav_vl_dict = sav_value_labels.get(p['raw'])
        new_name, label, notes, recoded = rename_and_label(
            p, family_rules, questions, question_text, sav_lbl, flat_mismatch,
            sav_vl_dict, other_positions)
        if recoded:
            recoded_rows.add(len(rows))
        rows.append([p['raw'], new_name, label, notes])
        rename_map[p['raw']] = new_name

    _flag_duplicates(rows)

    wb = Workbook()

    ws1 = wb.active
    ws1.title = 'Variable Label'
    ws1.append(['Variable Information'])
    ws1.append(['Raw Variable', 'Renamed Variable', 'Label', 'Notes'])
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    header_rows = 2  # 'Variable Information' title row + column-header row, before data starts
    for i, r in enumerate(rows):
        ws1.append(r)
        if i in recoded_rows:
            excel_row = header_rows + i + 1  # openpyxl rows are 1-indexed
            for col in range(1, 5):
                ws1.cell(row=excel_row, column=col).fill = yellow_fill

    # Value Label sheet intentionally omitted for now -- a different approach for
    # value labels is coming; rename_map is kept above since that logic will need it.

    return wb


def _flag_duplicates(rows):
    """Post-pass over the built rows (raw, renamed, label, notes): flag any renamed
    variable name or label that isn't unique, so nothing silently collides once this
    hits SPSS. Mutates each row's notes (index 3) in place."""
    from collections import Counter

    name_counts = Counter(r[1] for r in rows)
    label_counts = Counter(r[2] for r in rows if r[2])

    for r in rows:
        extra = []
        if name_counts[r[1]] > 1:
            extra.append(f'DUPLICATE renamed variable name (shared by {name_counts[r[1]]} '
                          'variables) — must be resolved before use in SPSS')
        if r[2] and label_counts[r[2]] > 1:
            extra.append(f'duplicate label (same text as {label_counts[r[2]] - 1} other '
                          'variable(s)) — please verify')
        if extra:
            r[3] = '; '.join([r[3]] + extra) if r[3] else '; '.join(extra)


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
