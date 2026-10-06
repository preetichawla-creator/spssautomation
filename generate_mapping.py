#!/usr/bin/env python3
"""
generate_mapping.py — Draft SPSS mapping file generator (Phase 1 of the
Word questionnaire + raw .sav -> mapping file -> SPSS syntax automation).

USAGE
    python generate_mapping.py --sav RawData.sav --qnr questionnaire.docx --out mapping.xlsx

OUTPUT
    An .xlsx with one sheet:
      - "Variable Label": Raw Variable | Renamed Variable | Raw Label | Label | Notes
        (Raw Label = the .sav's own variable label, verbatim, for reference alongside
        the cleaned Label)
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
    - Two-axis grids declared with ROWS:/COLUMNS: markers (rows often shared by reference,
      e.g. "INSERT CODES ALLOCATED AT DPROJECT_B2B"): label = "{parent} : {row} :: {column}",
      with the parent taken from the raw .sav label cut at the row text

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
import os
import re
import unicodedata
import zipfile
from collections import Counter, defaultdict

import pyreadstat
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

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

# Some questionnaires are bilingual (e.g. Canadian EN/FR surveys), with English and its
# translation running together with no consistent separator -- sometimes a literal "/",
# sometimes just back-to-back sentences ("...to continue. / Veuillez...", "...belong
# to. À quel groupe d'âge..."). Unlike Hindi/Tamil/Arabic (see NON_LATIN_RE), French
# shares the Latin alphabet, so it can't be filtered by Unicode block -- detected
# instead by French-specific accented characters, which don't otherwise appear in
# English survey text.
FRENCH_CHAR_RE = re.compile(r'[àâäéèêëïîôöùûüÿçœæÀÂÄÉÈÊËÏÎÔÖÙÛÜŸÇŒÆ]')

# High-confidence French-only function words -- none of these are ordinary English
# words, so matching one (case-insensitively, as a whole word) is a safe signal that
# French has started, even with no accented character anywhere near the boundary
# (e.g. "...consistent approach Je n'utilise pas toujours la même approche.").
FRENCH_WORD_RE = re.compile(
    r'^(je|tu|il|elle|nous|vous|ils|elles|le|la|les|un|une|des|du|de|au|aux|et|est|sont|'
    r'votre|vos|notre|nos|quel|quelle|quels|quelles|avec|dans|sur|sans|veuillez|'
    r'r[ée]pondre|r[ée]ponses?|autre|banque|entreprise|ces|cette|cet|aucune?|toujours|'
    r'm[êe]me|pas|plus|aussi|donc|qui|que|quoi|combien|pourquoi|comment)$',
    re.IGNORECASE
)


def strip_french(text):
    """Scans left to right for the first word that's unambiguously French (an accented
    character, or a French-only function word) and drops everything from there on --
    handles every separator style seen in this bilingual EN/FR document uniformly
    (period, colon, dash, or no separator at all), since English always comes first."""
    if not text:
        return text
    words = text.split(' ')
    cut_idx = None
    for i, w in enumerate(words):
        w_clean = w.strip('.,;:!?()[]{}\u2013\u2014"\'')
        if not w_clean:
            continue
        if FRENCH_CHAR_RE.search(w) or FRENCH_WORD_RE.match(w_clean):
            cut_idx = i
            break
    if cut_idx is None:
        return text
    truncated = ' '.join(words[:cut_idx]).rstrip(' ,;:\u2013\u2014-/')
    return truncated if truncated.strip() else text


# A generic "how to answer this question" instruction, e.g. "Select one response.",
# "Please select only one answer.", "Select all that apply.", "Choose multiple
# responses." -- carries no respondent-facing content of its own. Deliberately requires
# the ENTIRE sentence/clause to be just this generic methodology pattern (anchored start
# to end, not a substring search), so a genuine question that happens to use "select" --
# e.g. "Please select the language of the questionnaire." -- is never touched: "the
# language of the questionnaire" doesn't match (one|only one|all|multiple), so the whole
# pattern fails and the sentence is left alone.
INSTRUCTION_SENTENCE_RE = re.compile(
    r'^('
    r'(please\s+)?(select|choose)\s+(one|only\s+one|all|multiple)'
    r'(\s+(response|responses|answer|answers|option|options))?'
    r'(\s+(that\s+)?appl(y|ies))?'
    r'|'
    r'(one|only\s+one|all|multiple)\s+(response|responses|answer|answers|option|options)(\s+only)?'
    r')'
    r'\s*\.?$',
    re.IGNORECASE
)


def strip_instruction_sentences(text):
    if not text:
        return text
    parts = re.split(r'(?<=[.!?])\s+', text)
    kept = [p for p in parts if p.strip() and not INSTRUCTION_SENTENCE_RE.match(p.strip())]
    return ' '.join(kept).strip()


# Piped/conditional question wording written inline, e.g.
#   "IF A9 = 4,5,6,10: At what point ...? IF A9 = 3: At what stage ...?"
# The "IF <var> = <values>:" markers are routing instructions, never label text -- and
# since a single label can only carry one wording, the first variant (the default, as
# listed first in the questionnaire) is kept rather than running both together.
# Case-sensitive uppercase IF with a variable name and "=" is a strong routing signature,
# so ordinary sentences beginning "If ..." are never touched.
COND_VARIANT_RE = re.compile(r'\bIF\s+[A-Za-z_]\w*\s*(?:=|<>|!=)\s*[\d,\s\-]+\s*:\s*')


def strip_conditional_variants(text):
    if not text or not COND_VARIANT_RE.search(text):
        return text
    parts = COND_VARIANT_RE.split(text)
    return (parts[0] + ' ' + parts[1]).strip() if len(parts) >= 2 else text


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
    cleaned = strip_conditional_variants(cleaned)
    cleaned = re.sub(r'<[^>]*>', '', cleaned)
    # a [%code%] placeholder (dynamic pipe-in content) should become a visible blank
    # rather than vanish silently -- swapped to a marker first since the literal
    # underscores would otherwise be eaten by the blank-fill-run stripping below, then
    # converted to real underscores after that step has already run
    cleaned = re.sub(r'\[%[^\]]*%\]', '\x00BLANK\x00', cleaned)
    cleaned = re.sub(r'\[[^\]]*\]', '', cleaned)
    cleaned = re.sub(r'\([^)]*\)', '', cleaned)    # parenthetical content is never wanted in a label
    cleaned = strip_instruction_sentences(cleaned)  # before substring-level stripping below, so a
                                                     # full instruction sentence is recognized intact
                                                     # rather than partially mangled first
    cleaned = INSTRUCTION_PHRASES_RE.sub('', cleaned)
    cleaned = NON_LATIN_RE.sub('', cleaned)
    cleaned = strip_french(cleaned)
    cleaned = re.sub(r'_{3,}', '', cleaned)        # blank-fill-in underscore runs, e.g. "First____" -> "First"
    cleaned = cleaned.replace('\x00BLANK\x00', '_______')
    # leading UI marker glyphs (radio-button/checkbox bullets, a stray '#' programmer-note
    # marker, or a "(1)"/"(2)" style item-position marker) carry no content of their own.
    cleaned = re.sub(r'^[○□●■◯•]+\s*', '', cleaned)
    cleaned = re.sub(r'^\(\d+\)\s*', '', cleaned)
    cleaned = re.sub(r'^#+\s*', '', cleaned)
    cleaned = re.sub(r'[\(\[]\s*[\)\]]', '', cleaned)          # empty () or [] left behind
    cleaned = re.sub(r'\s{2,}', ' ', cleaned).strip()
    # a bare (non-bracketed) trailing routing verb -- e.g. "Marketing/Market research
    # Terminate", with no parentheses at all around it. The bracketed/parenthesized form
    # ("(TERMINATE)") is already gone via the [...] and (...) stripping above; this
    # catches the same instruction word when nothing wraps it. "continue" is handled
    # separately below, case-SENSITIVE: unlike these other terms, it's also an ordinary
    # English word ("...language to continue."), so only the all-caps routing-note form
    # is safe to strip.
    cleaned = re.sub(r'\s+((thank\s+and\s+)?terminate|anchor|rotate|randomize)\.?\s*$', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+CONTINUE\.?\s*$', '', cleaned)
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


def load_multilevel_numbering(docx_path):
    """Questionnaires written as "1.1  question / 1.2  question ... 2.1 ..." usually have
    those numbers supplied by a two-level Word list (level 0 = section, level 1 =
    "%1.%2." question), not typed into the text -- so python-docx never sees them.
    Reads word/numbering.xml and returns {numId: {ilvl: (lvlText, start)}} for every
    decimal multi-level list that has at least one level combining two counters (a
    lvlText with %1 and %2 in it). Lists with only simple one-counter levels are left to
    load_numbering_map above."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            if 'word/numbering.xml' not in z.namelist():
                return {}
            content = z.read('word/numbering.xml').decode('utf-8', errors='ignore')
    except Exception:
        return {}

    abstract_levels = {}
    for m in re.finditer(r'<w:abstractNum [^>]*w:abstractNumId="(\d+)".*?</w:abstractNum>', content, re.S):
        aid, block = m.group(1), m.group(0)
        levels, multi = {}, False
        for lm in re.finditer(r'<w:lvl w:ilvl="(\d+)"[^>]*>(.*?)</w:lvl>', block, re.S):
            ilvl, body = int(lm.group(1)), lm.group(2)
            fmt = re.search(r'w:numFmt w:val="([^"]*)"', body)
            txt = re.search(r'w:lvlText w:val="([^"]*)"', body)
            start = re.search(r'w:start w:val="(\d+)"', body)
            if not fmt or fmt.group(1) != 'decimal' or not txt:
                continue
            levels[ilvl] = (txt.group(1), int(start.group(1)) if start else 1)
            if len(re.findall(r'%\d', txt.group(1))) >= 2:
                multi = True
        if multi:
            abstract_levels[aid] = levels

    out = {}
    for m in re.finditer(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"/>', content):
        if m.group(2) in abstract_levels:
            out[m.group(1)] = abstract_levels[m.group(2)]
    return out


class MultiLevelCounter:
    """Reproduces the numbers Word would display for a multi-level list, one paragraph
    at a time, in document order. Call next_label() for EVERY paragraph, including empty
    ones: an empty level-0 paragraph is how many questionnaires mark a section break
    (the 1.x -> 2.x jump), and skipping it would leave every later number wrong.
    A deeper level appearing before any shallower one is treated as if that shallower
    level were at its start value (Word shows "1.1", not "0.1")."""

    def __init__(self, defs):
        self.defs = defs
        self.state = {}

    def next_label(self, paragraph):
        pPr = paragraph._p.pPr
        if pPr is None or pPr.numPr is None or pPr.numPr.numId is None:
            return None
        num_id = str(pPr.numPr.numId.val)
        if num_id not in self.defs:
            return None
        levels = self.defs[num_id]
        ilvl = int(pPr.numPr.ilvl.val) if pPr.numPr.ilvl is not None else 0
        if ilvl not in levels:
            return None
        st = self.state.setdefault(num_id, {})
        for lv in range(ilvl):
            st.setdefault(lv, levels.get(lv, ('', 1))[1])
        st[ilvl] = st[ilvl] + 1 if ilvl in st else levels[ilvl][1]
        for lv in [k for k in st if k > ilvl]:
            del st[lv]
        text = levels[ilvl][0]
        label = re.sub(r'%(\d)', lambda m: str(st.get(int(m.group(1)) - 1,
                                                      levels.get(int(m.group(1)) - 1, ('', 1))[1])), text)
        return label.strip().rstrip('.')


# A routing/instruction paragraph that should never be mistaken for an answer option
# when a question lists its options as bare paragraphs with no table at all (see
# para_options in parse_questionnaire). Deliberately only excludes paragraphs that are
# ENTIRELY a routing note (the whole line matches), since a routing tag glued onto the
# END of a real option ("Marketing/Market research Terminate") is handled separately by
# clean_text's trailing-routing-word strip, not here.
PARA_ROUTING_RE = re.compile(
    r'^\(?\s*(programmer|ask\b|skip\b|note\s*:|if\s+\S+\s*=|please\s+select|select\s+one|'
    r'do\s+not\s+(randomize|rotate|read)|rotate\b|randomize\b|anchor\b|hide\s+option|'
    r'insert\s+option|thank\s+and\s+terminate|terminate\b)',
    re.IGNORECASE
)


def looks_like_option_candidate(text, is_bold=False):
    """Is this stray paragraph (one that didn't match any question stem) plausibly one
    line of an option list presented with no table at all? Deliberately conservative:
    excludes anything phrased as a question (more likely a continuation of the question
    text itself -- see maybe_update_question_text), whole-line routing instructions,
    anything implausibly long for a single answer option, and -- same signal as the
    table-row 'caption' check -- a paragraph that's entirely bold, since instruction
    notes ("DO NOT RANDOMIZE LIST.") are consistently bolded while real options aren't."""
    text = text.strip()
    if not text or len(text) > 200 or '?' in text:
        return False
    if is_bold:
        return False
    if PARA_ROUTING_RE.match(text):
        return False
    return True


# Some questionnaires lay out a question as a genuine two-axis grid using explicit
# "ROWS:" / "COLUMNS:" section markers rather than a table -- and the row definitions
# are often not even listed per-question, but declared ONCE and reused by reference
# across many questions in a section, e.g.:
#   PROGRAMMING NOTE: CREATE NEW VARIABLE = DPROJECT_B2B
#   IF DSEGMENT = 1: new home build
#   IF DSEGMENT = 2: new custom/architect-designed home build
#   ...
#   F3. Which of the following channels...?
#   ROWS:
#   INSERT CODES ALLOCATED AT DPROJECT_B2B
#   COLUMNS:
#   - Builder's standard supply arrangement
#   ...
CREATE_VAR_RE = re.compile(r'CREATE\s+NEW\s+VARIABLE\s*=\s*(\S+)', re.IGNORECASE)
IF_CODE_LABEL_RE = re.compile(r'^IF\s+\w+\s*=\s*(\d+)\s*:\s*(.+)$', re.IGNORECASE)
INSERT_CODES_RE = re.compile(r'INSERT\s+CODES\s+ALLOCATED\s+AT\s+(\S+)', re.IGNORECASE)
ROWS_MARKER_RE = re.compile(r'^ROWS?\s*:?\s*$', re.IGNORECASE)
COLUMNS_MARKER_RE = re.compile(r'^COLUMNS?\s*:?\s*$', re.IGNORECASE)


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
        if m and m.group(2).strip():
            # a stem token with NOTHING following it is never a real question header --
            # it's typically a bare column label (e.g. a cross-reference table literally
            # headed "Category | Q11a | Q11b"), which would otherwise switch current_stem
            # back mid-table and silently merge two questions' option lists together
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


def paragraph_is_bold(paragraph):
    runs = [r for r in paragraph.runs if r.text.strip()]
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
    # Returned as ('legend', rest) rather than a bare skip: for an ordinary single-axis
    # rating grid this is just an answer-scale legend and the caller ignores it, same as
    # before -- but for a genuine two-axis grid (raw variables named ..._rN_cM) this row
    # IS the column codebook (Purchased through Pharmacy / Purchased Online / ...), and
    # `rest` is already exactly that: the row's cells minus its leading corner cell
    # (a blank corner never entered `nonints` to begin with, so `rest` already excludes a
    # real corner label like "Items" the same way it would for a coded item row).
    if not ints and len(nonints) > 1:
        return ('legend', rest)

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


def parse_questionnaire(path, raw_stems=frozenset(), rc_stems=frozenset(), dotted_stems=None):
    """Returns (questions, question_text, row_col):
       questions[stem]      = [(option_label, code), ...] in document order
       question_text[stem]  = the question's own prompt text
       row_col[stem]        = {'rows': [(label, code), ...], 'cols': [(label, code), ...]}
                               for a genuine two-axis grid question (see ROWS_MARKER_RE)
    raw_stems is the set of actual variable stems seen in the raw .sav -- used to
    resolve bare-number headers ("1.", "2.") to their real stem ("Q1", "Q2") only
    when that stem is confirmed to exist, rather than guessed.
    rc_stems is the subset of those stems that are genuine two-axis grids in the raw
    data (at least one variable named ..._rN_cM) -- used to decide whether a table's
    legend/column-header row (see classify_row) should be captured as a real column
    codebook, or left alone as just an answer-scale legend for an ordinary rating grid.
    dotted_stems maps a two-level Word list number ("2.8") to the raw stem it belongs to
    ("Q2X8"), for questionnaires whose question numbers exist only as list numbering
    (see load_multilevel_numbering) -- used when a numbered question has no typed stem.
    """
    doc = Document(path)
    items = list(iter_block_items(doc.element.body))
    numbering_map = load_numbering_map(path)
    numbering_counters = {}
    ml_counter = None
    if dotted_stems:
        ml_defs = load_multilevel_numbering(path)
        ml_counter = MultiLevelCounter(ml_defs) if ml_defs else None
    questions, question_text = {}, {}
    para_options = defaultdict(list)
    named_codebooks = defaultdict(list)   # name (e.g. "DPROJECT_B2B") -> [(label, code), ...]
    row_col = {}                          # stem -> {'rows': [...] | ('ref', name), 'cols': [...]}
    pending_codebook_name = None
    capture_mode = None                   # None | 'rows' | 'cols' -- which axis is being
                                           # captured for the CURRENT stem right now
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
        nonlocal current_stem, buf, capture_mode
        if stem != current_stem:
            flush(questions, current_stem, buf)
            buf = []
            current_stem = stem
            capture_mode = None  # row/column capturing is per-question, not carried over
        if text:
            maybe_update_question_text(stem, text)

    for it in items:
        if isinstance(it, Paragraph):
            ml_label = ml_counter.next_label(it) if ml_counter else None  # every paragraph, even empty
            text = it.text.strip()
            if not text:
                continue

            m = CREATE_VAR_RE.search(text)
            if m:
                pending_codebook_name = m.group(1)
                continue
            m = IF_CODE_LABEL_RE.match(text) if pending_codebook_name else None
            if m:
                named_codebooks[pending_codebook_name].append((clean_text(m.group(2)), int(m.group(1))))
                continue

            if current_stem and ROWS_MARKER_RE.match(text):
                row_col.setdefault(current_stem, {'rows': [], 'cols': []})
                capture_mode = 'rows'
                continue
            if current_stem and COLUMNS_MARKER_RE.match(text):
                row_col.setdefault(current_stem, {'rows': [], 'cols': []})
                capture_mode = 'cols'
                continue
            if capture_mode == 'rows':
                m = INSERT_CODES_RE.search(text)
                if m:
                    row_col[current_stem]['rows'] = ('ref', m.group(1))
                    continue
                if looks_like_option_candidate(text, is_bold=paragraph_is_bold(it)):
                    target = row_col[current_stem]['rows']
                    if isinstance(target, list):
                        target.append(text)
                    continue
            if capture_mode == 'cols':
                if looks_like_option_candidate(text, is_bold=paragraph_is_bold(it)):
                    row_col[current_stem]['cols'].append(text)
                    continue

            stem, rest = find_stem_and_text(text)
            if not stem:
                stem, rest = find_bare_number_stem(text)
                stem = resolve_sentinel(stem)
            if not stem and numbering_map:
                auto_stem = resolve_auto_number(it, numbering_map, numbering_counters)
                if auto_stem:
                    stem, rest = auto_stem, text
            if not stem and ml_label and dotted_stems and ml_label in dotted_stems:
                stem, rest = dotted_stems[ml_label], text
            if stem:
                set_stem(stem, rest)
            elif current_stem:
                maybe_update_question_text(current_stem, text)
                if looks_like_option_candidate(text, is_bold=paragraph_is_bold(it)):
                    para_options[current_stem].append(text)
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
                elif kind == 'legend' and current_stem in rc_stems and len(payload) >= 2:
                    # A genuine two-axis grid laid out as an ordinary table (column
                    # headers across the top, one item per row below -- see classify_row)
                    # rather than the ROWS:/COLUMNS: paragraph-marker style. The row side
                    # needs no separate capture here: those item rows are simultaneously
                    # buffered as normal 'textonly' entries above and become
                    # questions[stem] via the usual flush() below, which the row_col
                    # post-processing step reuses as the row codebook once this loop ends.
                    row_col.setdefault(current_stem, {'rows': [], 'cols': []})
                    row_col[current_stem]['cols'] = list(payload)
    flush(questions, current_stem, buf)

    # Some questionnaires list a question's options as plain consecutive paragraphs --
    # no table at all, not even a codeless one (see classify_row's textonly path, which
    # only helps for genuine table rows). Used ONLY as a last resort: a stem whose table
    # parsing already produced a codebook is left alone, and this needs at least 2
    # candidates, since a single leftover paragraph is far more likely to be a missed
    # continuation of the question text (see maybe_update_question_text) than an option.
    for stem, opts in para_options.items():
        if not questions.get(stem) and len(opts) >= 2:
            questions[stem] = [(clean_text(opt), i) for i, opt in enumerate(opts, start=1)]

    # Resolve row references to their named codebook (declared earlier in the doc, see
    # CREATE_VAR_RE/IF_CODE_LABEL_RE), and assign sequential codes to columns, which are
    # always a plain option list with no explicit numbering of their own.
    for stem, rc in row_col.items():
        rows = rc['rows']
        if isinstance(rows, tuple) and rows[0] == 'ref':
            rc['rows'] = list(named_codebooks.get(rows[1], []))
        elif isinstance(rows, list) and rows:
            rc['rows'] = [(clean_text(t), i) for i, t in enumerate(rows, start=1)]
        elif not rows and questions.get(stem):
            # A table-based two-axis grid (see the 'legend' handling above): the row
            # side was never captured directly here -- it comes for free from the
            # item rows' ordinary 'textonly' processing into questions[stem] already.
            rc['rows'] = questions[stem]
        rc['cols'] = [(clean_text(t), i) for i, t in enumerate(rc['cols'], start=1)]

    return questions, question_text, row_col


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


# Some SPSS exports represent a multi-select item as a checked/unchecked flag rather
# than a labelled category code -- e.g. Q1_1's own value labels are just {0: 'Unchecked',
# 1: 'Checked'}, describing whether THIS box was ticked, not which item it is. That
# carries no usable item text (and using it naively is actively wrong, not just
# unhelpful: position 1 of a family would silently collide with value 1 = 'Checked' and
# other low positions with 0 = 'Unchecked', since Python treats int 1 and float 1.0 as
# equal dict keys). Detected once per variable and excluded from the item-text fallback.
CHECKBOX_FLAG_LABELS = {'checked', 'unchecked', 'selected', 'not selected', 'yes', 'no', 'true', 'false'}


def is_checkbox_flag_labels(vl_dict):
    if not vl_dict or len(vl_dict) > 2:
        return False
    texts = {str(v).strip().lower() for v in vl_dict.values()}
    return bool(texts) and texts <= CHECKBOX_FLAG_LABELS


def find_flat_match(pos, qcodes_pairs):
    """For a flat-family raw position, find its matching questionnaire entry by treating
    `pos` as a POSITIONAL index into the parsed list -- the raw file's Nth flat-family
    variable corresponds to the Nth option listed in the questionnaire, and that entry's
    own code (read from the doc) becomes the rename target, which may differ from `pos`
    itself (e.g. position 10 of a 9-item list -> a reserved code like 98 for "Any Other").
    Returns (label, code, recoded) or None if `pos` is out of the list's range.
    (An earlier version tried a direct code-equals-pos match first, for platforms that
    export the intended code as the suffix directly rather than a sequential position --
    reverted per instruction: that convention is a one-off, and positional matching is
    what every other study, including checked/unchecked-format ones, actually needs.)"""
    if not qcodes_pairs or pos > len(qcodes_pairs):
        return None
    label, code = qcodes_pairs[pos - 1]
    return label, code, (code != pos)


def resolve_item_label(p, pos, key, fam_size, qcodes_pairs, sav_value_label_dict, other_positions,
                        row_col_entry=None):
    """Figure out the item-specific text (if any) for one grid/flat-family member,
    without composing or truncating anything -- shared by the family-wide reserve
    calculation (compute_family_reserve) and rename_and_label, so both agree on
    exactly what text a sibling will need.
    Returns (item_label_or_None, is_other_pair, note_or_None). is_other_pair means
    the caller should compose ' : Others' / ' : Others Specify' directly rather than
    use item_label (which is None in that case).
    Checked PER POSITION, not per stem: a stem's codebook can legitimately cover only
    some of the raw positions in its family (e.g. a superset of reserved codes in the
    raw schema that this particular questionnaire wave never actually asked) -- that
    doesn't make the codebook untrustworthy for the positions it DOES cover."""
    if pos is not None and (key, pos) in other_positions:
        return None, True, None

    # A genuine two-axis grid (both row AND column vary) with its own separate row and
    # column codebooks (see ROWS_MARKER_RE in parse_questionnaire) -- e.g. F3_r1_c1 with
    # row 1 = "new home build" and column 1 = "Builder's standard supply arrangement".
    # Composed here as "row :: col" (double colon, to read as a second-level split from
    # the single colon already separating the parent question from this whole suffix)
    # and returned as a normal item_label, so the usual " : {item_label}" composition
    # in rename_and_label naturally produces "{parent} : {row} :: {col}". A single-axis
    # question that shares the same row codebook (F1_r1 .. F1_r5 -- one variable per row,
    # the column choice living in the answer value itself) just gets the row as its item.
    # Deliberately NOT an early return: this flows through the same item-length cap below
    # as every other item source, since an early return would silently skip that cap.
    axis_label = None
    rows = (row_col_entry or {}).get('rows') or []
    cols = (row_col_entry or {}).get('cols') or []
    if rows:
        if (p['kind'] == 'rc' and cols and p['row'] is not None and p['col'] is not None
                and p['row'] <= len(rows) and p['col'] <= len(cols)):
            axis_label = f"{rows[p['row'] - 1][0]} :: {cols[p['col'] - 1][0]}"
        elif p['kind'] == 'r' and p['row'] is not None and p['row'] <= len(rows):
            axis_label = rows[p['row'] - 1][0]

    note = None
    match = find_flat_match(pos, qcodes_pairs) if p['kind'] == 'flat' else None
    if axis_label is not None:
        item_label = axis_label
    elif match:
        item_label = match[0]
    elif qcodes_pairs and pos is not None and pos <= len(qcodes_pairs) and p['kind'] != 'flat':
        item_label = qcodes_pairs[pos - 1][0]
    elif p['kind'] == 'flat' and pos is not None:
        # this position isn't covered by the questionnaire codebook (missing entirely,
        # or just doesn't reach this far) -- try the raw data's own value labels as a
        # last resort, skipped entirely for checked/unchecked-style labels since those
        # describe the box's state, not which item it is
        note = ('no questionnaire match for this position'
                + ('' if qcodes_pairs else ' (question stem not found in questionnaire)'))
        item_label = None
        if sav_value_label_dict and not is_checkbox_flag_labels(sav_value_label_dict):
            raw_item = sav_value_label_dict.get(pos, sav_value_label_dict.get(float(pos)))
            item_label = clean_text(raw_item) if raw_item else None
        if item_label:
            note += ' — used the raw data\'s own value labels for the option text instead'
        else:
            note += ' — no matching option in raw data value labels either, label may be incomplete'
    elif qcodes_pairs is None and fam_size > 1:
        item_label, note = None, 'question stem not found in questionnaire — label may be incomplete'
    else:
        item_label = None

    # An item's own text can itself be a full paragraph (e.g. a long statement to agree
    # or disagree with). Cap it well under the label limit so the parent question always
    # keeps a meaningful share of the budget too -- otherwise a single huge item starves
    # every sibling's parent text down to a near-meaningless fragment (see compose logic
    # in compute_family_reserve/rename_and_label, which size the parent around this cap).
    if item_label and len(item_label) > MAX_ITEM_TEXT_CHARS:
        item_label, _ = truncate_label(item_label, MAX_ITEM_TEXT_CHARS)
        note = (note + '; ' if note else '') + 'item text itself shortened to fit label limits'

    return item_label, False, note


def _has_content(fragment):
    """True if the fragment has real words in it -- at least two consecutive letters or
    digits. A blank-fill placeholder ("_______", what a [% ... %] pipe becomes) or bare
    punctuation has none, so it is never mistaken for a sentence worth keeping."""
    return re.search(r'[^\W_]{2,}', fragment) is not None


def pick_meaningful_text(text, limit):
    """When the full text won't fit, prefer dropping whole LEADING sentences over a
    blind character cut -- the later sentence(s) in a questionnaire prompt are more
    often the actual question, with earlier ones being scene-setting context (e.g.
    "Up until <date>, X happened. Given that, would you...?" -- the second sentence
    alone is usually the meaningful, self-contained ask). Falls back to a plain
    character truncation only if no sentence-boundary cut fits.
    Two guards, both found on a real question ("As far as you know, which areas of
    cooperation ... China ? _______") whose only text after the question mark was a
    blank-fill placeholder: that placeholder alone was being picked as "the later,
    more specific sentence", throwing away the whole actual question. So (1) a
    trailing placeholder is dropped first, and if the real text then fits, that is
    the result; (2) a candidate with no real words in it is never chosen."""
    if not text or len(text) <= limit:
        return text
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    core = list(sentences)
    while len(core) > 1 and not _has_content(core[-1]):
        core.pop()
    trimmed = ' '.join(core).strip()
    if trimmed and len(trimmed) <= limit:
        return trimmed
    for i in range(1, len(sentences)):
        candidate = ' '.join(sentences[i:]).strip()
        if candidate and _has_content(candidate) and len(candidate) <= limit:
            return candidate
    fallback, _ = truncate_label(trimmed or text, limit)
    return fallback


def compute_family_reserve(parsed_all, family_rules, questions, other_positions, sav_value_labels, row_col):
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
        sav_value_label_dict = sav_value_labels.get(p['raw'])
        item_label, is_other_pair, _ = resolve_item_label(
            p, pos, key, fam_size, qcodes_pairs, sav_value_label_dict, other_positions,
            row_col.get(p['stem']))
        if is_other_pair:
            needed = len(" : Others Specify") if p['other'] else len(" : Others")
        elif item_label:
            needed = len(f" : {item_label}")
        else:
            continue
        reserve[p['stem']] = max(reserve[p['stem']], needed)
    return reserve


# A question-level name written "<section>X<question>" (Q1X1, Q2x11, Q2X10_r12.23, Q2X12a)
# is the same section-dot-question notation as QM1x1 -> QM1.1 elsewhere in this file, so
# the question part is written with a dot in the renamed variable: Q1.1, Q2.11, ...
# Only the question part is rewritten and everything after it is left exactly as it is,
# so names stay unique and traceable. (The lowercase-x loop names that parse_raw already
# splits into stem + loop never reach this with an "x" left in them.)
X_NOTATION_RE = re.compile(r'^([A-Za-z]+\d+)[Xx](\d+)(?=$|[A-Za-z_.])')


def to_dot_notation(stem):
    return X_NOTATION_RE.sub(r'\1.\2', stem, count=1)


def dotted_loop_name(m, shape):
    """Renamed form of a loop-style variable with a trailing ".N" iteration number
    (Q2X10_r12.23 -> Q2.10_12.23), built the way every other grid in this file is:
    the question part in dot notation, the r / c letters dropped, an axis that never
    varies within this question dropped too (Q13_r1_c1 .. r8_c1 -> Q13_1 .. Q13_8),
    and the loop number last, after a dot. `m` is a DOTTED_VAR_RE match; `shape` is
    that question's {'r': {...}, 'c': {...}, 'f': {...}} of distinct indexes seen.
    Names stay unique because the only axes dropped are ones that cannot tell two
    variables apart. Returns (name, both_axes_vary)."""
    shape = shape or {}
    parts = []
    for key in 'rcf':
        if m.group(key) and len(shape.get(key, ())) > 1:
            parts.append(m.group(key))
    both_vary = bool(m.group('r') and m.group('c')
                     and len(shape.get('r', ())) > 1 and len(shape.get('c', ())) > 1)
    name = to_dot_notation(m.group('base')) + ''.join('_' + x for x in parts)
    return f"{name}.{m.group('d')}", both_vary


def rename_and_label(p, family_rules, questions, base_text, sav_label,
                      sav_value_label_dict, other_positions, row_col, ctx=None):
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

    new_stem = to_dot_notation(p['stem'] + (f".{p['loop']}" if p['loop'] else ""))
    dotted_m = DOTTED_VAR_RE.match(p['raw']) if (ctx and p['kind'] == 'plain') else None
    if dotted_m:
        new_stem, both_vary = dotted_loop_name(dotted_m, ctx['shape'].get(dotted_m.group('base')))
        if both_vary:
            notes.append('ambiguous grid order (both row & col vary) — used default row-then-column, please verify')
    key = (p['stem'], p['loop'])
    qcodes_pairs = questions.get(p['stem'])
    # base_text already prefers the raw .sav's own label over the questionnaire's (see
    # build_workbook) -- this is just whichever of the two survived for this stem.
    parent_text = base_text.get(p['stem'], '')

    if p['kind'] == 'plain':
        new_name = new_stem + ('_other' if p['other'] else '')
        dm = DOTTED_VAR_RE.match(p['raw']) if (ctx and not sav_label) else None
        if dm and dm.group('base') in ctx['linked']:
            base = dm.group('base')
            qtext = base_text.get(base, '')
            if qtext:
                # A loop-style variable (Q2X10_r12.23) whose .sav label is blank. The
                # questionnaire supplies the question text; the ".23" is the loop
                # iteration (which celebrity, brand, etc. -- the questionnaire doesn't
                # list the loop items, they are piped in at runtime), so it is shown as
                # "Loop 23 :" the same way the S2a.N loop variables are.
                shape = ctx['shape'].get(base, {})
                cb = questions.get(base)
                bits = []
                if dm.group('r'):
                    r = int(dm.group('r'))
                    if cb and r <= len(cb):
                        bits.append(cb[r - 1][0])
                    else:
                        bits.append(f'Row {r}')
                        notes.append('row text not available from the questionnaire for this '
                                     'question, so the row number is used')
                if dm.group('c') and len(shape.get('c', ())) > 1:
                    bits.append(f"Column {int(dm.group('c'))}")
                if dm.group('f') and len(shape.get('f', ())) > 1:
                    bits.append(f"Item {int(dm.group('f'))}")
                label = f"Loop {int(dm.group('d'))} : {qtext}"
                for bit in bits:
                    label = compose_with_suffix(label, f' : {bit}')
                notes.append('no label text in the .sav: built from the questionnaire question '
                             'text; the loop number is shown because the questionnaire does not '
                             'list the loop items')
                return finalize(new_name, label)
        label = sav_label or parent_text or new_name
        if not qcodes_pairs and not parent_text and not sav_label:
            notes.append('no rule matched (kept as raw) — likely needs a study-specific name')
        return finalize(new_name, label)

    fam_size = family_rules['fam_size'].get(key, 1)
    pos = item_position(p, family_rules, key)

    if key in family_rules['composite']:
        notes.append('composite/multi-part question (mixed grid shapes under one stem) — '
                      'rename/label may not match your team\'s convention here, please verify')

    # --- suffix / rename ---
    if p['kind'] == 'flat':
        match = find_flat_match(pos, qcodes_pairs)
        if match:
            _, code, recoded = match
            suffix = f"_{code}"
        else:
            # this position isn't covered by the questionnaire codebook (missing
            # entirely, or just doesn't reach this far) -- the doc can't tell us the
            # right code here, so keep the raw suffix as-is rather than risk a wrong one.
            # resolve_item_label (below) adds the note explaining why.
            suffix = f"_{pos}"
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
        p, pos, key, fam_size, qcodes_pairs, sav_value_label_dict, other_positions,
        row_col.get(p['stem']))
    if item_note:
        if ctx and p['stem'] in ctx['guarded']:
            # the stem WAS found; it is its option list that didn't line up (see load_study)
            item_note = item_note.replace('question stem not found in questionnaire',
                                          "questionnaire's option list does not line up with the raw variables")
        notes.append(item_note)

    if is_other_pair:
        base_label = compose_with_suffix(parent_text, " : Others") if parent_text else "Others"
        label = compose_with_suffix(base_label, " Specify") if p['other'] else base_label
        return finalize(new_name, label)

    if parent_text and item_label and fam_size > 1:
        label = compose_with_suffix(parent_text, f" : {item_label}")
    elif parent_text and ctx and p['stem'] in ctx['guarded'] and fam_size > 1:
        # The questionnaire's option list for this question doesn't line up with its raw
        # variables (see load_study), so no option text is used. Without anything to tell
        # the variables apart they would all share one label, so the position is shown.
        if p['kind'] == 'rc':
            where = f"Row {p['row']} Column {p['col']}"
        elif p['kind'] == 'r':
            where = f"Row {p['row']}"
        elif p['kind'] == 'c':
            where = f"Column {p['col']}"
        else:
            where = f"Item {p['col']}"
        label = compose_with_suffix(parent_text, f" : {where}")
        if not any('does not line up' in n for n in notes):
            notes.append("questionnaire's option list does not line up with the raw variables, "
                         "so no option text is used and only the position is shown")
        else:
            notes.append('only the position is shown')
    elif parent_text:
        label = parent_text
    elif item_label:
        label = item_label
    else:
        label = sav_label or new_name
        if not sav_label and not notes:
            # only when nothing else has already flagged this row (a flat-family variable
            # with no questionnaire match is already marked "label may be incomplete")
            notes.append('no label text in the .sav and none found in the questionnaire: '
                         'the variable name is used as a placeholder label')

    if p['other']:
        # no detected non-other sibling at this position (rare -- e.g. a genuine two-axis
        # grid item) -- fall back to the old plain-append behavior rather than dropping it
        label = compose_with_suffix(label, " : Others")

    return finalize(new_name, label)


# ----------------------------------------------------------------------------
# 4. Main: read .sav, parse questionnaire, build both sheets
# ----------------------------------------------------------------------------
# A loop-iteration variable named "stem.N" (e.g. S2a.1 .. S2a.24, one variable per idea
# a respondent evaluated) -- as opposed to the "stemxN_..." notation elsewhere in this
# file, this is a bare dot before the loop number with nothing else attached.
LOOP_DOT_RE = re.compile(r'^(.+?)\.(\d+)$')


def clean_sav_label(varname, stem, raw_label):
    """Strip a leading 'varname - ' or 'stem - ' prefix off a raw .sav variable label.
    Multi-select family members (Q31_1, Q31_2, ...) commonly all share one label that's
    prefixed with the STEM ('Q31 - ...'), not the full variable name, so both are tried."""
    if not raw_label:
        return ''
    lbl = re.sub(r'^' + re.escape(varname) + r'\s*-\s*', '', raw_label)
    if lbl == raw_label and stem and stem != varname:
        lbl = re.sub(r'^' + re.escape(stem) + r'\s*-\s*', '', raw_label)
    if lbl == raw_label:
        # Some exports give a grid's variables NO label text at all, just the QUESTION's
        # code and a dash ("Q2X8 -" on Q2X8_r1_c1.1). That is a blank label, not a label
        # reading "Q2X8" -- left as-is it would be passed along as if it were real text and
        # hide that there is nothing there. Deliberately only the label-is-nothing-but-the-
        # code case (the token must contain a digit and be a prefix of this variable's
        # name): a code prefix in front of REAL text is left exactly as before.
        pm = re.fullmatch(r'\s*(\S*\d\S*)\s*-\s*', raw_label)
        if pm and len(pm.group(1)) >= 2 and varname.lower().startswith(pm.group(1).lower()):
            lbl = ''

    # For a loop-iteration variable, the raw label's own [%...%] placeholder (a dynamic
    # "which loop item" pipe-in the platform never resolved when exporting to .sav) is
    # replaced with this variable's own loop number instead of becoming a blank -- e.g.
    # S2a.1's "[% LoopLabel(S2aLoop) %] You have said..." becomes "Loop 1 : You have
    # said...". Only the first placeholder is touched (there's normally just one, right
    # at the start); anything else still becomes a blank via clean_text as usual.
    loop_m = LOOP_DOT_RE.match(varname)
    if loop_m:
        lbl = re.sub(r'\[%[^\]]*%\]', f'Loop {loop_m.group(2)} :', lbl, count=1)

    return clean_text(lbl)


def looks_truncated_by_source(raw_label):
    """SPSS (or the platform exporting to it) can itself cut a variable label off
    mid-sentence at its own length limit (commonly ~250 chars) -- when that's happened,
    the raw label is genuinely missing content that no cleanup on our end can recover
    (the rest was never exported at all), so it should be treated as unusable rather
    than kept as a broken fragment. Signal: suspiciously close to that length AND not
    ending on real sentence-final punctuation."""
    if not raw_label or len(raw_label) < 245:
        return False
    return not re.search(r'[.?!][”"\')]*\s*$', raw_label.strip())


# Question-level stem of a raw variable named "<letters><section>X<question>[a-z]", e.g.
# Q2X8 (question 2.8) or Q2x11 -- followed by end of name, "_" or "." so that something
# like Q1X3PC is NOT mistaken for a stem.
QSTEM_BASE_RE = re.compile(r'^([A-Za-z]+(\d+)[Xx](\d+)([a-z]?))(?=$|_|\.)')
# The same, for a loop-style name with a trailing ".N" iteration number, e.g.
# Q2X8_r1_c1.1 or Q2X10_r12.23.
DOTTED_VAR_RE = re.compile(
    r'^(?P<base>[A-Za-z]+\d+[Xx]\d+[a-z]?)(?:_r(?P<r>\d+))?(?:_c(?P<c>\d+))?(?:_(?P<f>\d+))?\.(?P<d>\d+)$')


def link_numbered_stems(raw_cols):
    """For questionnaires whose question numbers are only Word list numbering ("2.8"),
    work out which raw stem each number belongs to, from the raw variable names alone.
    Returns (dotted_stems, expected_len, shape):
      dotted_stems  {"2.8": "Q2X8"} -- only where exactly one raw stem fits that number
                    (a stem with a trailing letter, like Q2X12a, is a sub-part that has
                    no number of its own, so is never linked)
      expected_len  {stem: highest row/column/item index its variables use} -- how long
                    that question's option list has to be to line up with the data
      shape         {stem: {'r': {...}, 'c': {...}, 'f': {...}}} -- distinct row, column
                    and item indexes seen for loop-style (".N") variables of that stem"""
    bases, expected, shape = {}, {}, {}
    for c in raw_cols:
        m = QSTEM_BASE_RE.match(c)
        if not m:
            continue
        base = m.group(1)
        bases[base] = (int(m.group(2)), int(m.group(3)), m.group(4))
        pm = re.match(r'(?:_r(\d+))?(?:_c(\d+))?(?:_(\d+))?', c[len(base):])
        idx = next((int(g) for g in pm.groups() if g), None)
        if idx:
            expected[base] = max(expected.get(base, 0), idx)
        dm = DOTTED_VAR_RE.match(c)
        if dm:
            sh = shape.setdefault(base, {'r': set(), 'c': set(), 'f': set()})
            for k in 'rcf':
                if dm.group(k):
                    sh[k].add(int(dm.group(k)))
    by_num = defaultdict(set)
    for base, (sec, q, letter) in bases.items():
        if not letter:
            by_num[f"{sec}.{q}"].add(base)
    # Platforms often add helper variables built from a question's code ("forQ1x1_1" ...
    # "forQ1x1_23"), which would make the number 1.1 ambiguous. The real question
    # variables share one prefix across the whole study ("Q"), so when a number has more
    # than one candidate, keep only those with that dominant prefix -- and still refuse
    # to link unless exactly one is left.
    prefix_of = lambda b: re.match(r'[A-Za-z]+', b).group(0).lower()
    counts = defaultdict(int)
    for b in bases:
        counts[prefix_of(b)] += 1
    dominant = max(counts, key=counts.get) if counts else None
    dotted = {}
    for k, v in by_num.items():
        if len(v) > 1:
            v = {b for b in v if prefix_of(b) == dominant}
        if len(v) == 1:
            dotted[k] = next(iter(v))
    return dotted, expected, shape


def load_study(qnr_path, raw_cols, parsed_all):
    """Everything derived from the questionnaire that both the mapping and the recode
    generators need, computed once and identically for both.
    Returns (questions, question_text, row_col, ctx). ctx['guarded'] lists numbered
    questions whose option list does NOT line up with their raw variables (a table
    shared by two questions, or a grid laid out the other way round, e.g. celebrity
    names across the top): their option text is discarded rather than risk attaching the
    wrong option to a variable -- only the question text is used for them."""
    raw_stems = {p['stem'] for p in parsed_all}
    rc_stems = {p['stem'] for p in parsed_all if p['kind'] == 'rc'}
    dotted_stems, expected_len, shape = link_numbered_stems(raw_cols)
    questions, question_text, row_col = parse_questionnaire(
        qnr_path, raw_stems, rc_stems, dotted_stems=dotted_stems or None)
    guarded = set()
    for base in dotted_stems.values():
        exp, cb = expected_len.get(base), questions.get(base)
        if exp and exp > 1 and (cb is None or len(cb) != exp):
            guarded.add(base)
            questions.pop(base, None)
    ctx = {'guarded': guarded, 'shape': shape, 'linked': set(dotted_stems.values())}
    return questions, question_text, row_col, ctx


def split_sav_label_at_row(raw_label, stem, rows):
    """For a question laid out with ROWS:/COLUMNS: (see ROWS_MARKER_RE), the raw .sav
    label is the whole thing run together: "{stem} - {intro} {question} {Row} {Column}".
    Finds which known row label appears in it (case-insensitively -- the questionnaire's
    own row text is often lowercase where the .sav's is sentence-case) and cuts there.
    Returns (parent_text, row_code, row_text_as_written_in_sav) or None if no row matches.
    Cutting at the ROW rather than looking for the column means this still works when
    the label was truncated by the source partway through the column text."""
    lbl = re.sub(r'^' + re.escape(stem) + r'\s*-\s*', '', raw_label.strip())
    low = lbl.lower()
    best = None
    for row_label, code in rows:
        idx = low.find(row_label.lower())
        if idx > 0 and (best is None or idx < best[0]):
            best = (idx, code, lbl[idx: idx + len(row_label)])
    if best is None:
        return None
    idx, code, cased = best
    return lbl[:idx].rstrip(), code, cased


def build_workbook(sav_path, qnr_path):
    df, meta = pyreadstat.read_sav(sav_path, metadataonly=True)
    raw_cols = meta.column_names
    sav_labels = meta.column_names_to_labels
    sav_value_labels = meta.variable_value_labels

    parsed_all = [parse_raw(c) for c in raw_cols]
    questions, question_text, row_col, ctx = load_study(qnr_path, raw_cols, parsed_all)

    family_rules = build_family_rules(parsed_all)

    # Positions that have a "please specify" write-in companion (X_N + X_N_other) --
    # these are always the catch-all "Others" option; see rename_and_label.
    other_positions = set()
    for p in parsed_all:
        if p['other']:
            key = (p['stem'], p['loop'])
            pos = item_position(p, family_rules, key)
            if pos is not None:
                other_positions.add((key, pos))

    # Each stem's "parent" text for label composition: prefer the raw .sav's own
    # variable label, and only fall back to the questionnaire's text when the .sav
    # genuinely has nothing usable there. "Usable" specifically means shared across the
    # whole family: some platforms bake each item's OWN specific text directly into its
    # individual sav label (e.g. B5_r1's label ends "...I like my home to have a
    # coordinated, cohesive design throughout" -- that row's own statement, not a
    # generic parent), rather than repeating one generic label across every sibling like
    # most platforms do. Reusing just one member's label as the shared "parent" for all
    # of them would durably bake that ONE item's text into every sibling's label. Detected
    # by collecting every distinct cleaned sav label across a stem's family: if they
    # don't all agree, the sav label isn't safe to use as shared parent text at all, and
    # the questionnaire (which has no such per-item variation) is used instead.
    base_text = {}
    sav_label_variants = defaultdict(set)
    for p in parsed_all:
        stem = p['stem']
        raw_lbl = sav_labels.get(p['raw'])
        if raw_lbl and looks_truncated_by_source(raw_lbl):
            continue  # cut off at the source -- not usable, let questionnaire fall through
        sav_lbl = clean_sav_label(p['raw'], stem, raw_lbl)
        if sav_lbl:
            sav_label_variants[stem].add(sav_lbl)
    for stem, variants in sav_label_variants.items():
        if len(variants) == 1:
            base_text[stem] = next(iter(variants))
    # Questions laid out with ROWS:/COLUMNS: (row_col): each variable's raw label is the
    # shared parent text with that variable's own row (and column) text appended, so the
    # generic "do the siblings' labels agree?" check above always fails for them even
    # though the parent portion IS identical. Cut every sibling's label at its row text
    # instead, and use the parent if they all agree (else fall through to the
    # questionnaire, as before). Also recovers each row's casing as written in the .sav.
    global_casing = {}   # row text (lowercased) -> casing as written in any .sav label
    for stem, rc in row_col.items():
        rows = rc['rows']
        if not rows:
            continue
        parents = set()
        for p in parsed_all:
            if p['stem'] != stem:
                continue
            raw_lbl = sav_labels.get(p['raw'])
            res = split_sav_label_at_row(raw_lbl, stem, rows) if raw_lbl else None
            if res:
                parent, code, cased = res
                parents.add(parent)
                global_casing.setdefault(cased.lower(), cased)
        if len(parents) == 1:
            base_text[stem] = clean_text(next(iter(parents)))
    # Questions sharing one row codebook should share its casing too -- a question whose
    # own raw label was truncated before the row text (so it couldn't recover the casing
    # itself) borrows it from a sibling question that could.
    for rc in row_col.values():
        rc['rows'] = [(global_casing.get(label.lower(), label), code) for label, code in rc['rows']]

    for stem, qtext in question_text.items():
        if not base_text.get(stem):
            base_text[stem] = qtext

    # Resolve each multi-member stem's parent text ONCE, sized to fit alongside the
    # longest item any of its siblings will need -- so every sibling shares identical
    # text before " : ", instead of each independently truncating the parent based on
    # its own (shorter or longer) item text.
    reserve = compute_family_reserve(parsed_all, family_rules, questions,
                                      other_positions, sav_value_labels, row_col)
    for stem, needed in reserve.items():
        if stem in base_text:
            budget = max(LABEL_CHAR_LIMIT - needed, 20)
            base_text[stem] = pick_meaningful_text(base_text[stem], budget)

    rows = []
    raw_labels = []   # parallel to rows: the .sav's own label, verbatim
    rename_map = {}
    recoded_rows = set()   # row indices (0-based within `rows`) where a code was recoded
    for p in parsed_all:
        sav_lbl = clean_sav_label(p['raw'], p['stem'], sav_labels.get(p['raw']))
        sav_vl_dict = sav_value_labels.get(p['raw'])
        new_name, label, notes, recoded = rename_and_label(
            p, family_rules, questions, base_text, sav_lbl,
            sav_vl_dict, other_positions, row_col, ctx)
        if recoded:
            recoded_rows.add(len(rows))
        # Written verbatim. The one unavoidable exception: characters Excel itself cannot
        # store in a cell (control characters) would make the save fail, so those alone are
        # dropped -- and the row is noted, so "exact" is never silently untrue.
        raw_lbl = sav_labels.get(p['raw'])
        if raw_lbl and ILLEGAL_CHARACTERS_RE.search(raw_lbl):
            raw_lbl = ILLEGAL_CHARACTERS_RE.sub('', raw_lbl)
            notes = (notes + '; ' if notes else '') + \
                'raw label contained characters Excel cannot store — those were removed from Raw Label'
        raw_labels.append(raw_lbl)
        rows.append([p['raw'], new_name, label, notes])
        rename_map[p['raw']] = new_name

    _flag_duplicates(rows)

    wb = Workbook()

    ws1 = wb.active
    ws1.title = 'Variable Label'
    ws1.append(['Variable Information'])
    ws1.append(['Raw Variable', 'Renamed Variable', 'Raw Label', 'Label', 'Notes'])
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    header_rows = 2  # 'Variable Information' title row + column-header row, before data starts
    for i, r in enumerate(rows):
        # r = [raw variable, renamed, label, notes]; Raw Label slots in as column C
        ws1.append([r[0], r[1], raw_labels[i], r[2], r[3]])
        if i in recoded_rows:
            excel_row = header_rows + i + 1  # openpyxl rows are 1-indexed
            for col in range(1, 6):
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


# ----------------------------------------------------------------------------
# Phase 2: SPSS syntax generation from a (possibly hand-edited) mapping file
# ----------------------------------------------------------------------------
SPSS_RESERVED_WORDS = {
    'ALL', 'AND', 'BY', 'EQ', 'GE', 'GT', 'LE', 'LT', 'NE', 'NOT', 'OR', 'TO', 'WITH'
}
SPSS_VALID_NAME_RE = re.compile(r'^[A-Za-z@#$][A-Za-z0-9_.@#$]*$')


def validate_spss_name(name):
    """Returns a short problem description if `name` isn't a legal SPSS variable name,
    or None if it's fine. Checked so a problem surfaces here, in a warning, rather than
    as a cryptic error when the generated syntax is actually run in SPSS."""
    if not name:
        return 'empty name'
    if len(name) > 64:
        return f'{len(name)} characters, over the 64-character SPSS limit'
    if not SPSS_VALID_NAME_RE.match(name):
        return 'contains characters not valid in an SPSS variable name'
    if name.upper() in SPSS_RESERVED_WORDS:
        return 'is an SPSS reserved word'
    if name.upper().startswith('SYS_') or name.upper().startswith('SYSTEM_'):
        return None  # sys_ names are conventional here, not actually reserved
    return None


def read_mapping_rows(mapping_path):
    """Reads the final, human-finalized mapping file: exactly 3 columns, always in this
    position regardless of what their header text says (per the team's own workflow --
    the header wording can vary run to run, the position doesn't):
      column A = Raw Variable, column B = Renamed Variable, column C = Variable Label
    This is deliberately position-based, not header-name-based: the header ROW is still
    skipped (whatever it says), but which column is which is never inferred from that
    text. Works whether or not there's a leading title row above the headers (our own
    tool's draft output has one, a hand-built 3-column file might not) -- detected by
    checking how many of the first row's first three cells are populated.
    A file left with more than 3 columns is very likely the 5-column DRAFT mapping,
    not the finalized one -- there, column C is Raw Label (the raw .sav text), not the
    finalized Variable Label, and using it as one would silently mislabel every
    variable. That specific case is caught explicitly; a generic extra-columns file
    still gets a warning instead of failing, in case it's simply an extra note column
    someone added that isn't yet accounted for above."""
    wb = load_workbook(mapping_path, read_only=True)
    ws = wb['Variable Label'] if 'Variable Label' in wb.sheetnames else wb.worksheets[0]
    rows = [r for r in ws.iter_rows(values_only=True) if any(c not in (None, '') for c in r)]
    if not rows:
        raise ValueError('This file has no data in it')

    first_row_filled = sum(1 for c in rows[0][:3] if c not in (None, ''))
    header_idx = 0 if first_row_filled >= 2 else 1
    if len(rows) <= header_idx:
        raise ValueError('This file has a header row but no data rows below it')
    header = rows[header_idx]

    if len(header) > 3 and str(header[2] or '').strip().lower() == 'raw label':
        raise ValueError(
            "Column C here is \"Raw Label\", the raw .sav text -- this looks like the "
            "5-column draft mapping from step 1, not the finalized 3-column file "
            "(Raw Variable, Renamed Variable, Variable Label). Please upload the "
            "finalized version, with the Variable Label itself in column C.")

    warnings = []
    if sum(1 for c in header if c not in (None, '')) > 3:
        warnings.append('This file has more than 3 columns. Only columns A, B and C were '
                         'used (as Raw Variable, Renamed Variable, Variable Label) -- '
                         'please confirm that matches what you intended.')

    out = []
    for r in rows[header_idx + 1:]:
        out.append({
            'raw': str(r[0]).strip() if len(r) > 0 and r[0] not in (None, '') else '',
            'renamed': str(r[1]).strip() if len(r) > 1 and r[1] not in (None, '') else '',
            'label': str(r[2]).strip() if len(r) > 2 and r[2] not in (None, '') else '',
        })
    return out, warnings


# Scripts that mark a switch to another LANGUAGE when they appear after text that started in
# Latin script (Hindi, Tamil, Arabic, Chinese, ... after English).
_FOREIGN_SCRIPTS = {
    'DEVANAGARI', 'TAMIL', 'TELUGU', 'KANNADA', 'MALAYALAM', 'BENGALI', 'GUJARATI', 'GURMUKHI',
    'ORIYA', 'SINHALA', 'ARABIC', 'HEBREW', 'CYRILLIC', 'GREEK', 'THAI', 'LAO', 'KHMER', 'MYANMAR',
    'TIBETAN', 'GEORGIAN', 'ARMENIAN', 'HANGUL', 'HIRAGANA', 'KATAKANA', 'CJK', 'BOPOMOFO', 'ETHIOPIC',
}
_ENGLISH_STOPWORDS = {
    'the', 'of', 'and', 'to', 'in', 'for', 'or', 'with', 'not', 'your', 'you', 'are', 'is', 'that',
    'this', 'from', 'have', 'how', 'what', 'which', 'about', 'than', 'more', 'most', 'very', 'all',
}


def _foreign_script(ch):
    """Name of the non-Latin language script this character belongs to, or None. Combining
    vowel signs are included (a Devanagari word is not made of letters alone)."""
    if not (unicodedata.category(ch)[0] in 'LM'):
        return None
    first = unicodedata.name(ch, '').split(' ')[0]
    return first if first in _FOREIGN_SCRIPTS else None


def _is_french_word(w):
    return bool(FRENCH_WORD_RE.match(w.strip('.,;:!?()[]{}"\'')))


def french_bilingual_study(sav_value_labels, threshold=0.20, minimum_labels=20):
    """Is this study's data English AND French (same script, so the languages can't be told
    apart by script)? Judged across all its value labels: a real bilingual export has
    French function words AND English function words in a large share of them (Canada:
    48% and most; every other study checked: 1% or less French). Requiring English too
    matters: a study written ONLY in French must be left alone, and the French cutter,
    which assumes English comes first, would otherwise damage it ("Pas du tout
    satisfait" would become "Pas")."""
    labels = {str(t).strip() for vl in sav_value_labels.values() for t in vl.values() if isinstance(t, str)}
    labels = {t for t in labels if re.search(r'[A-Za-z]{3}', t)}
    if len(labels) < minimum_labels:
        return False
    fr = sum(1 for t in labels if any(_is_french_word(w) for w in t.split()))
    en = sum(1 for t in labels if any(w.lower().strip('.,;:!?()[]{}"\'') in _ENGLISH_STOPWORDS for w in t.split()))
    return fr / len(labels) >= threshold and en / len(labels) >= threshold


_NUMBER_TOKEN_RE = re.compile(r'^[\d.,\-\u2013+%$\u20ac\u00a3/:\u00a0]+$')
_NUMBER_UNIT_WORDS = {'million', 'millions', 'billion', 'billions', 'milliard', 'milliards',
                      'thousand', 'thousands'}


def _drop_translated_numbers(head):
    """After the second language is cut off, its translation of any NUMBERS in the label
    is still sitting at the end of what is left, because a number comes before the first
    word of that language ("18-29 years old 18-29", "100,001 - 150,000 INR 100,001 -
    150,000", "$1,000 to $2,499 1 000 $"). Those trailing number-only words are removed
    when their digits already appear earlier in the label. Only ever run on a label that
    has already been identified as carrying a second language, never on a normal one."""
    toks = head.split(' ')
    n = len(toks)
    best = 0
    for k in range(1, n):                                  # always keep at least one word
        suffix = toks[n - k:]
        if not all(_NUMBER_TOKEN_RE.match(t) or t.lower() in _NUMBER_UNIT_WORDS for t in suffix):
            break
        # what is removed must START with a number: "... $3 million 1 million $" can lose
        # "1 million $" (a translated "$1 million") but never reach back into "$3 million"
        if not re.search(r'\d', suffix[0]):
            continue
        digits = re.sub(r'\D', '', ' '.join(suffix))
        if digits and digits in re.sub(r'\D', '', ' '.join(toks[:n - k])):
            best = k
    if not best:
        return head
    return re.sub(r'[\s\-\u2013\u2014/|:;,\u00b7\u2022(\[]+$', '', ' '.join(toks[:n - best])) or head


# French-only words and elisions beyond the variable-label detector's list, used ONLY for
# value labels (and only inside a study recognised as English/French). Words that are also
# ordinary English ("plus", "an", "non", "a") are deliberately not here.
_FRENCH_EXTRA_WORDS = {
    'ou', 'ni', 'mais', 'pour', 'par', 'ans', 'mois', 'jours', 'fois', 'chez', 'leur', 'leurs',
    'mon', 'mes', 'ses', 'tr\u00e8s', 'tous', 'toutes', 'toute', 'oui', 'peu', 'trop', 'assez',
}
_FRENCH_ELISION_RE = re.compile(r"^(?:[jlmnstcd]|qu)['\u2019][a-z\u00e0-\u00ff]", re.I)
_PUNCT = '.,;:!?()[]{}"\''


def _french_evidence(word):
    wc = word.strip(_PUNCT)
    if not wc:
        return False
    return bool(FRENCH_CHAR_RE.search(word) or FRENCH_WORD_RE.match(wc)
                or wc.lower() in _FRENCH_EXTRA_WORDS or _FRENCH_ELISION_RE.match(wc))


def _french_cut(s):
    """Cut English-then-French text at the start of the French half, in a study already
    known to be English/French. Starts from the first word that is unmistakably French,
    then corrects the two ways that goes wrong on real labels:
      - the French half usually starts BEFORE its first French-only word (a translated
        noun like "Voyages", "Commerce"). When the English half ended with a closed
        parenthetical -- "Travel (flights, hotels) Voyages (vols, ...)" -- and what follows
        it is 1-3 capitalised words with no English function word in them, the French half
        starts right after that parenthetical.
      - a label must never END inside an unclosed parenthesis: that opener belongs to the
        French half, so it and the word before it are dropped.
    Returns the label unchanged if the French starts at the very first word (nothing in
    English to keep)."""
    words = s.split(' ')
    cut = next((i for i, w in enumerate(words) if _french_evidence(w)), None)
    if not cut:
        return s
    end = cut
    depth, closed_at = 0, None
    for j, w in enumerate(words[:cut]):
        for ch in w:
            if ch == '(':
                depth += 1
            elif ch == ')' and depth:
                depth -= 1
                if depth == 0 and closed_at is None:
                    closed_at = j
    if closed_at is not None and closed_at + 1 < cut:
        between = words[closed_at + 1:cut]
        if (len(between) <= 3 and between[0][:1].isupper()
                and not any(w.lower().strip(_PUNCT) in _ENGLISH_STOPWORDS for w in between)):
            end = closed_at + 1
    stack = []
    for j, w in enumerate(words[:end]):
        for ch in w:
            if ch == '(':
                stack.append(j)
            elif ch == ')' and stack:
                stack.pop()
    if stack:
        end = max(stack[0] - 1, 1)
    head = re.sub(r'[\s,;:\u2013\u2014\-/]+$', '', ' '.join(words[:end]))
    return head or s


def keep_first_language(text, french_study=False):
    """For a value label that carries more than one language, keep only the FIRST
    (English, in every export seen) and drop the rest. Returns (text, what_happened) with
    what_happened one of None, 'script', 'french', 'starts_foreign'.
    A label in a single language is never changed, whatever the language: German,
    Spanish, French, Chinese ... all pass through untouched. Only changes that can be
    made reliably are made:
      - text that starts in Latin script and later switches to another script is cut at
        the switch ("Prefer not to answer उत्तर देना ..." -> "Prefer not to answer").
        A lone Greek letter (µ in "µg") is a symbol, not a language switch.
      - English followed by French in the SAME script can't be told apart by script, so
        that is only attempted in a study that french_bilingual_study() recognises as
        English/French, using the same detector as the variable labels. It is best
        effort: a few French words can be left behind.
      - a label that STARTS in a non-Latin script is left exactly as it is."""
    s = str(text).strip()
    first_alpha = next((c for c in s if c.isalpha()), None)
    if first_alpha is None:
        return s, None
    if _foreign_script(first_alpha):
        return s, 'starts_foreign'
    for i, ch in enumerate(s):
        sc = _foreign_script(ch)
        if not sc:
            continue
        if sc == 'GREEK' and not (i + 1 < len(s) and _foreign_script(s[i + 1]) == 'GREEK'):
            continue
        head = re.sub(r'[\s\-\u2013\u2014/|:;,\u00b7\u2022(\[]+$', '', s[:i])
        return _drop_translated_numbers(head or s), 'script'
    if french_study:
        spaced = s.replace('\u00a0', ' ')   # French typesetting joins "70\xa0ans"; split so words are visible
        cut = _french_cut(spaced).strip()
        if cut and cut != spaced:
            return _drop_translated_numbers(cut), 'french'
    return s, None


def _clean_code(code):
    """SPSS value-label keys come back from pyreadstat as floats even for whole-number
    codes (7.0, not 7) -- formatted as a plain int here so generated syntax reads "7",
    matching every hand-written example seen, not "7.0". A genuinely non-integer code
    (rare, but SPSS does allow decimal values) is left untouched."""
    try:
        f = float(code)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return code


def _normalize_for_match(text):
    """Lowercase, drop parenthetical content, and strip everything but letters/digits --
    a forgiving but still meaningful text match between a raw .sav value label and a
    questionnaire option (e.g. "Any Other (Please Specify)" and "Any Other" normalize
    to the same thing). Used only to decide whether two options are the same thing;
    never changes what's actually written to the output."""
    t = re.sub(r'\([^)]*\)', '', str(text))
    return re.sub(r'[^a-z0-9]+', '', t.lower())


def compute_value_recode_plan(parsed_all, family_rules, questions, sav_value_labels, rename_map,
                              skipped=None, lang_stats=None):
    """For every variable where the raw value codes don't match the questionnaire's
    intended codes, builds the (old=new) recode mapping and the value-label text to
    use -- taken directly from the raw .sav (per the team's convention: only the CODE
    changes, never the label wording).
    Two kinds of variable qualify:
    (1) A 'flat'-kind multi-select family member whose own stored value equals its raw
        position when selected (an "item number" convention -- confirmed against real
        study data, not a generic 0/1 flag) where find_flat_match already found
        position != intended code.
    (2) A single variable holding one of several named codes (a single-select
        question), where the same kind of positional mismatch can occur between the
        raw codes (sorted ascending) and the questionnaire's own option list, checked
        position by position the same way find_flat_match does for (1).
    A variable using the "Checked"/"Unchecked" convention is always excluded from
    both: a binary flag has no code of its own to recode, only a variable name to be
    attached to -- already handled by the rename step.
    Label text is kept VERBATIM from the raw .sav (only whitespace-trimmed, since the
    team's instruction is explicit: value labels are taken directly from the source,
    unlike variable labels which go through the full cleanup pipeline elsewhere).
    Returns a list of dicts: {'variable': renamed_name, 'recodes': [(old, new), ...],
    'value_labels': [(code, label_text), ...]}, one per variable that needs anything,
    in raw-column order.
    `skipped`, if given, is a list that gets one (variable, reason, [(old, new, label)])
    entry appended for every single-select variable where a label matched the
    questionnaire under a DIFFERENT code but a safety check refused to act on it -- so
    a legitimate recode is never lost silently, only held back for a human to decide."""
    plan = []
    skipped = skipped if skipped is not None else []
    # Value labels that carry more than one language keep only the first (see
    # keep_first_language); `lang_stats`, if given, is a Counter of what was done.
    french_study = french_bilingual_study(sav_value_labels)

    def V(text, tally=None):
        # `tally` collects what was done for the variable being built; it is only merged
        # into lang_stats if that variable ends up in the plan, so the totals describe
        # labels that are actually written, not every label that was looked at.
        out, how = keep_first_language(str(text).strip(), french_study)
        if tally is not None and how:
            tally[how] += 1
        return out

    def text_diffs(vl, qcodes_pairs):
        qt = {}
        for lab, code in qcodes_pairs:
            qt.setdefault(_normalize_for_match(lab), code)
        return [(_clean_code(k), qt[_normalize_for_match(v)], str(v).strip())
                for k, v in vl.items()
                if _normalize_for_match(v) in qt and qt[_normalize_for_match(v)] != _clean_code(k)]

    for p in parsed_all:
        raw = p['raw']
        vl = sav_value_labels.get(raw)
        if not vl or is_checkbox_flag_labels(vl):
            continue
        new_name = rename_map.get(raw, raw)

        if p['kind'] == 'flat':
            key = (p['stem'], p['loop'])
            pos = item_position(p, family_rules, key)
            qcodes_pairs = questions.get(p['stem'])
            match = find_flat_match(pos, qcodes_pairs) if pos is not None else None
            if not match or not match[2]:
                continue  # not recoded -- its own stored value already matches its name
            _, new_code, _ = match
            label_text = vl.get(pos, vl.get(float(pos)))
            if not label_text:
                continue  # raw .sav has no label for the value actually being recoded
            tally = Counter()
            plan.append({
                'raw': raw,
                'variable': new_name,
                'recodes': [(pos, new_code)],
                'value_labels': [(new_code, V(label_text, tally))],
                '_lang': tally,
            })

        elif p['kind'] == 'plain':
            qcodes_pairs = questions.get(p['stem'])
            # Require the raw value-labels count to exactly match the questionnaire's
            # own option count for this stem before trusting ANY text match for it.
            # Confirmed against real, working examples: both the reference file's
            # Q1X1 (23 raw codes, 23 questionnaire options) and this tool's own
            # McDonald's A16 (7 and 7) match exactly. A real counterexample found
            # while testing this: a detailed age-in-years variable (30 raw codes)
            # happened to text-match one entry ("46+") against an UNRELATED banded
            # age-group question with only 3 entries in the questionnaire -- same
            # wording, completely different question. An exact count match is what
            # distinguishes "this IS the same option list" from "one label happens to
            # coincide" -- a looser size check (within some margin) would still have
            # let that case through.
            if not qcodes_pairs or len(vl) < 2:
                continue
            if len(vl) != len(qcodes_pairs):
                diffs = text_diffs(vl, qcodes_pairs)
                if diffs:
                    skipped.append((new_name,
                                    f'its raw value labels have {len(vl)} codes but the questionnaire '
                                    f'lists {len(qcodes_pairs)} options for this question',
                                    diffs))
                continue
            # Matched by TEXT, not position: positional matching (safe for a flat
            # multi-select family, where each raw variable genuinely IS the Nth listed
            # option, one variable per option in document order) is NOT safe for a
            # single-select variable's own value set -- its raw codes don't reliably
            # line up ordinally with the questionnaire's list. Matching each raw code's
            # own label text against the questionnaire's option text only acts where we
            # can actually tell which option a code represents, and says nothing where
            # we can't, rather than guessing from position or count alone.
            q_by_text = {}
            for label, code in qcodes_pairs:
                q_by_text.setdefault(_normalize_for_match(label), code)
            # Work out the code every option WOULD end up with, then check the result as a
            # whole: after recoding, every option must still have its own distinct code. That
            # one rule covers every way a recode can silently merge categories -- a target
            # code already held by an option that is staying put, OR two options both being
            # sent to the same new code (seen on real data: a questionnaire that gave two
            # different industries the code 0). A shift where every option moves to a free
            # code is fine in a single SPSS RECODE, which evaluates each case once against
            # its original value, so e.g. (11=9)(9=0) cannot chain.
            proposed = []
            tally = Counter()
            for orig_code, label_text in vl.items():
                raw_code = _clean_code(orig_code)
                lab = V(label_text, tally)
                intended = q_by_text.get(_normalize_for_match(lab))
                proposed.append((raw_code, raw_code if intended is None else intended, lab))
            changed = [(r, f, lab) for r, f, lab in proposed if f != r]
            if not changed:
                continue
            by_final = defaultdict(list)
            for r, f, lab in proposed:
                by_final[f].append(lab)
            clashes = {c: labs for c, labs in by_final.items() if len(labs) > 1}
            if clashes:
                detail = '; '.join(f"code {c}: " + ', '.join(repr(lab[:30]) for lab in labs[:3])
                                   for c, labs in list(clashes.items())[:2])
                skipped.append((new_name,
                                f'recoding it would give more than one option the same code ({detail})',
                                changed))
                continue
            plan.append({
                'raw': raw,
                'variable': new_name,
                'recodes': [(r, f) for r, f, _ in changed],
                'value_labels': sorted(((f, lab) for _, f, lab in proposed), key=lambda t: float(t[0])),
                '_lang': tally,
            })

    # ------------------------------------------------------------------------------------
    # Second pass: clean the value labels of EVERY mapped variable to the first language,
    # not just the ones that needed a recode. A variable already in the plan keeps its recode;
    # any other variable whose labels carry a second language gets its own Val Lab block with
    # its raw codes (right as they are: nothing is recoded for it). Labels always come from
    # the raw .sav; the questionnaire only ever decides which CODE a recoded option gets.
    # ------------------------------------------------------------------------------------
    by_raw = {e['raw']: e for e in plan}
    numeric_order = lambda t: float(t[0])
    final = []
    for p in parsed_all:
        raw = p['raw']
        entry = by_raw.get(raw)
        vl = sav_value_labels.get(raw)
        if vl and raw in rename_map:
            trimmed, tally, changed_any = {}, Counter(), False
            for k, t in vl.items():
                out, how = keep_first_language(str(t).strip(), french_study)
                trimmed[k] = out
                if how:
                    tally[how] += 1
                changed_any = changed_any or out != str(t).strip()
            if changed_any and any(isinstance(k, str) for k in vl):
                # string-coded variable: its Val Lab needs quoted string values, which this
                # does not write -- reported, not guessed at
                if lang_stats is not None:
                    lang_stats['string_coded'] += 1
            elif changed_any:
                if entry is None:
                    entry = {'raw': raw, 'variable': rename_map[raw], 'recodes': [],
                             'value_labels': sorted(((_clean_code(k), lab) for k, lab in trimmed.items()),
                                                    key=numeric_order),
                             '_lang': tally}
                elif p['kind'] == 'flat':
                    # A recoded multi-select item: its entry carries only its OWN label, but its
                    # dictionary holds the whole shared option list, still with translations.
                    # Write all of it, with the recoded code swapped in -- unless that would
                    # make two options share a code, in which case the single label stays.
                    swap = {_clean_code(o): n for o, n in entry['recodes']}
                    full = [(swap.get(_clean_code(k), _clean_code(k)), lab) for k, lab in trimmed.items()]
                    if len({c for c, _ in full}) == len(full):
                        entry = dict(entry, value_labels=sorted(full, key=numeric_order), _lang=tally)
        if entry is not None:
            final.append(entry)

    if lang_stats is not None:
        for e in final:
            lang_stats.update(e['_lang'])
            if e['_lang']['script'] or e['_lang']['french']:
                lang_stats['variables'] += 1
    return final


def format_recode_value_syntax(plan):
    """Formats a compute_value_recode_plan() result as SPSS syntax, matching the team's
    own convention: command name, variable name on its own line, each entry with no
    indentation, the closing period attached to the final entry, double-quoted label
    text (apostrophes dropped, an embedded double quote escaped by doubling -- same
    treatment as the rename/label syntax, for the same reasons -- see
    generate_spss_syntax), and EXECUTE. after every block.
    Grouped PER VARIABLE: each variable's RECODE is immediately followed by its own
    Val Lab, then the next variable's pair, so a question's recode and the labels that
    go with it read together instead of being split into an all-RECODE section and an
    all-Val-Lab section."""
    lines = []
    for entry in plan:
        var, recodes, value_labels = entry['variable'], entry['recodes'], entry['value_labels']

        if recodes:
            lines.append('RECODE')
            lines.append(var)
            for i, (old, new) in enumerate(recodes):
                terminator = '.' if i == len(recodes) - 1 else ''
                lines.append(f'({old}={new}){terminator}')
            lines.append('EXECUTE.')
            lines.append('')

        if not value_labels:
            continue
        lines.append('Val Lab')
        lines.append(var)
        for i, (code, label) in enumerate(value_labels):
            clean_label = label.replace("'", '').replace('\u2019', '').replace('\u2018', '')
            clean_label = clean_label.replace('"', '""')
            terminator = '.' if i == len(value_labels) - 1 else ''
            lines.append(f'{code}"{clean_label}"{terminator}')
        lines.append('EXECUTE.')
        lines.append('')

    return '\n'.join(lines)


def generate_recode_value_syntax(sav_path, qnr_path, mapping_path):
    """Top-level entry point: builds RECODE + Val Lab syntax for a study, using the
    finalized mapping file's own Raw Variable -> Renamed Variable pairs as the
    authoritative variable names (respecting any manual renaming the team did), and the
    raw .sav + questionnaire to find which values actually need recoding and what their
    raw labels say. Returns (syntax_text, warnings)."""
    df, meta = pyreadstat.read_sav(sav_path, metadataonly=True)
    raw_cols = meta.column_names
    sav_value_labels = meta.variable_value_labels

    parsed_all = [parse_raw(c) for c in raw_cols]
    questions, _, _, _ = load_study(qnr_path, raw_cols, parsed_all)
    family_rules = build_family_rules(parsed_all)

    mapping_rows, warnings = read_mapping_rows(mapping_path)
    rename_map = {r['raw']: r['renamed'] for r in mapping_rows if r['raw'] and r['renamed']}

    skipped = []
    lang_stats = Counter()
    plan = compute_value_recode_plan(parsed_all, family_rules, questions, sav_value_labels, rename_map,
                                     skipped=skipped, lang_stats=lang_stats)
    if lang_stats['script'] or lang_stats['french']:
        parts = []
        if lang_stats['script']:
            parts.append(f"{lang_stats['script']} with a second language in another script")
        if lang_stats['french']:
            parts.append(f"{lang_stats['french']} with French after the English (best effort, "
                         "a few French words can remain)")
        warnings.append(f"Value labels with more than one language were cut to the first language "
                        f"in {lang_stats['variables']} variable(s): " + ' and '.join(parts) + '.')
    if lang_stats['string_coded']:
        warnings.append(f"{lang_stats['string_coded']} string-coded variable(s) have value labels with more "
                        'than one language and were NOT changed: string values need different syntax, '
                        'so these need to be done by hand.')
    if lang_stats['starts_foreign']:
        warnings.append(f"{lang_stats['starts_foreign']} value label(s) start in a non-Latin script "
                        'and were left exactly as they are.')
    # SPSS stores a value label in 120 bytes, so a long label (especially one that also carried
    # translations) can already be cut off IN THE RAW DATA. Labels are taken from the raw
    # SPSS as asked, so those are written as they are -- but said out loud, not left to be
    # noticed as a word that stops halfway.
    cut_off = sorted({lab for e in plan for _, lab in e['value_labels']
                      if len(lab) >= 118 and re.search(r'[A-Za-z0-9,]$', lab)})
    if cut_off:
        warnings.append(f'{len(cut_off)} distinct value label(s) look cut off at SPSS\'s 120-character '
                        'limit in the raw data itself (for example ending "...' + cut_off[0][-25:] +
                        '") and are written exactly as the raw SPSS has them; complete them by hand if '
                        'you want the full wording.')
    for var, reason, diffs in skipped:
        shown = ', '.join(f'{o}->{n} ({lab[:40]!r})' for o, n, lab in diffs[:3])
        more = f' and {len(diffs) - 3} more' if len(diffs) > 3 else ''
        warnings.append(f'{var}: NOT recoded automatically, because {reason}, so the match is not safe '
                        f'to trust. Label text suggests: {shown}{more}. Check by hand before recoding.')
    if not any(e['recodes'] for e in plan):
        if skipped:
            warnings.append('No recode was generated automatically. See the held-back '
                             'variable(s) listed above, which need a manual decision.')
        elif plan:
            warnings.append('No variable needed a recode: the file only contains value labels '
                             'cleaned to the first language.')
        else:
            warnings.append('Nothing to recode: no variable both has non-checkbox value '
                             'labels and a raw code that differs from its intended code.')
    return format_recode_value_syntax(plan), warnings


def generate_spss_syntax(mapping_path):
    """Builds SPSS syntax (RENAME VARIABLES, then VARIABLE LABELS, each followed by its
    own EXECUTE.) from the finalized 3-column mapping file (see read_mapping_rows for
    the exact column contract). Returns (syntax_text, warnings).
    Formatted to match the team's own SPSS convention (see the reference syntax they
    supplied): one entry per line with no indentation, the closing period attached
    directly to the final entry rather than on its own line, and double-quoted label
    text (an apostrophe in the label -- "Builder's" -- is dropped entirely rather than
    escaped, becoming "Builders", and an actual embedded double quote is escaped by
    doubling it instead, same as SPSS's own convention for an embedded single quote). Every mapped
    variable is listed in RENAME VARIABLES, including a "no-op" rename where the raw
    and renamed name are identical -- SPSS accepts this fine, and it matches how their
    own reference syntax lists every variable rather than only the ones that changed.
    RECODE and value labels are a separate file -- see generate_recode_value_syntax,
    which needs the raw .sav and questionnaire as well, since this mapping file
    carries no value-level data."""
    rows, warnings = read_mapping_rows(mapping_path)

    renames = []
    labels = []
    seen_names = {}

    for r in rows:
        raw, new, label = r['raw'], r['renamed'], r['label']
        if not raw or not new:
            continue

        problem = validate_spss_name(new)
        if problem:
            warnings.append(f'{new!r}: {problem} -- included below, but SPSS will likely reject it')

        if new in seen_names:
            warnings.append(f'{new!r}: renamed variable name used more than once '
                             f'(also for {seen_names[new]!r}) -- only the first entry '
                             'was kept, since SPSS would error on a repeat')
            continue
        seen_names[new] = raw

        renames.append((raw, new))
        if label:
            if len(label) > 251:
                warnings.append(f'{new!r}: label is {len(label)} characters, truncated to 251 '
                                 'for SPSS\'s variable-label limit')
                label = label[:251]
            # Apostrophes are dropped entirely here (not escaped, not kept) --
            # "Builder's" becomes "Builders" -- per the team's convention for this
            # syntax output specifically; the mapping file's own Label column is
            # untouched. Covers the straight apostrophe and the curly/smart variants
            # ' and ' that text pasted from Word or a survey platform often carries.
            clean_label = label.replace("'", '').replace('\u2019', '').replace('\u2018', '')
            labels.append((new, clean_label.replace('"', '""')))

    def block(command, entries, formatter):
        out = [command]
        for i, e in enumerate(entries):
            terminator = '.' if i == len(entries) - 1 else ''
            out.append(formatter(e) + terminator)
        out.append('EXECUTE.')
        out.append('')
        return out

    lines = []
    if renames:
        lines += block('RENAME VARIABLES', renames, lambda e: f'({e[0]}={e[1]})')
    if labels:
        lines += block('VARIABLE LABELS', labels, lambda e: f'{e[0]}"{e[1]}"')

    if not renames and not labels:
        warnings.append('Nothing to generate -- no rows had both a Raw Variable and a '
                         'Renamed Variable, or none had a Label.')

    return '\n'.join(lines), warnings


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    subparsers = ap.add_subparsers(dest='command')

    map_ap = subparsers.add_parser('map', help='Generate a draft mapping .xlsx (default if no command given)')
    map_ap.add_argument('--sav', required=True, help='Path to the raw .sav data file')
    map_ap.add_argument('--qnr', required=True, help='Path to the Word questionnaire (.docx)')
    map_ap.add_argument('--out', required=True, help='Path to write the draft mapping .xlsx')

    syntax_ap = subparsers.add_parser(
        'syntax', help='Generate SPSS syntax (.txt) from the finalized mapping .xlsx: '
                       'rename + variable labels, and (with --sav and --qnr) recode + value labels')
    syntax_ap.add_argument('--mapping', required=True, help='Path to the finalized 3-column mapping .xlsx')
    syntax_ap.add_argument('--out', required=True, help='Path to write the rename + variable labels .txt')
    syntax_ap.add_argument('--sav', help='Raw .sav (enables the recode + value labels file)')
    syntax_ap.add_argument('--qnr', help='Questionnaire .docx (needed together with --sav)')
    syntax_ap.add_argument('--recode-out', help='Path to write the recode + value labels .txt '
                                                '(default: <out> with "_recode" before the extension)')

    # Back-compat: `--sav X --qnr Y --out Z` with no subcommand still works as `map`.
    ap.add_argument('--sav', help=argparse.SUPPRESS)
    ap.add_argument('--qnr', help=argparse.SUPPRESS)
    ap.add_argument('--out', help=argparse.SUPPRESS)
    args = ap.parse_args()

    if args.command == 'syntax':
        text, warnings = generate_spss_syntax(args.mapping)
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f'Wrote {args.out}')
        for w in warnings:
            print(f'Warning: {w}')
        if bool(args.sav) != bool(args.qnr):
            ap.error('--sav and --qnr must be given together')
        if args.sav and args.qnr:
            recode_out = args.recode_out
            if not recode_out:
                base, ext = os.path.splitext(args.out)
                recode_out = f'{base}_recode{ext or ".txt"}'
            rtext, rwarnings = generate_recode_value_syntax(args.sav, args.qnr, args.mapping)
            if rtext.strip():
                with open(recode_out, 'w', encoding='utf-8') as f:
                    f.write(rtext)
                print(f'Wrote {recode_out}')
            else:
                # same as the app: no empty file, just say there was nothing to write
                print(f'No recode file written (nothing to recode).')
            for w in rwarnings:
                print(f'Warning: {w}')
        return

    sav, qnr, out = args.sav, args.qnr, args.out
    if not (sav and qnr and out):
        ap.error('--sav, --qnr and --out are required')
    wb = build_workbook(sav, qnr)
    wb.save(out)
    print(f'Wrote {out}')


if __name__ == '__main__':
    main()
