"""
utils/normalizer.py

Transcript normalization (V1 / V3) — shared by analysis-pipeline.py.

Reusable API:
    make_v1(turns)              — written-form cleanup of spk_NNNN turns
    make_v3(turns_v1, table)    — spoken-form expansion of V1 turns
    load_abbrev_table(tsv_path) — load abbrev-expansions.tsv
    parse_turns(path) / write_turns(turns, path)

Version meanings:
  V1 — basic cleanup:
         • whitespace / non-ASCII / dash normalisation
         • dotted initials spaced        (A.K. → A K)
         • broken contractions restored  (I ve → I've)
         • [UNCLEAR] → <noise>
         • parentheses spaced            (7(2) → 7 ( 2 ))
  V3 — v1 + abbreviation/acronym expansion (abbrev-expansions.tsv)
         + parentheses stripped  (7 ( 2 ) → 7 2)
         + numbers → spoken-form words   (7 2 → seven two)
         + fully lower-cased  (<tags> preserved)

Pure library module — all functions take explicit paths/arguments. Invoked by
analysis-pipeline.py's --get-norm-trans flag.
"""

import re
import sys
from pathlib import Path

try:
    import num2words as nw
except ImportError:
    sys.exit("num2words not installed — run: pip install num2words")


# ── Tag preservation ──────────────────────────────────────────────────────────
# <noise> and any future <tag> markers are kept verbatim through all transforms.

_TAG_RE = re.compile(r'<[^>]+>')


def apply_outside_tags(text: str, fn) -> str:
    """Apply fn to each text segment outside <tag> sequences; tags pass through unchanged."""
    parts, last = [], 0
    for m in _TAG_RE.finditer(text):
        parts.append(fn(text[last:m.start()]))
        parts.append(m.group(0))
        last = m.end()
    parts.append(fn(text[last:]))
    return ''.join(parts)


# ── Number conversion ─────────────────────────────────────────────────────────
# Chunk style for 3-/4-digit numbers:
#   370  → three seventy      1950 → nineteen fifty
#   300  → three hundred      1000 → one thousand
#   1100 → eleven hundred

_ORDINAL_RE = re.compile(r'\b(\d{1,4})(st|nd|rd|th)\b', re.IGNORECASE)
_NUMBER_RE  = re.compile(r'\b(\d+)\b')


_DIGIT_WORDS = {
    '0': 'zero', '1': 'one', '2': 'two', '3': 'three', '4': 'four',
    '5': 'five', '6': 'six', '7': 'seven', '8': 'eight', '9': 'nine',
}


def _digit_by_digit(n_str: str) -> str:
    """Speak each digit individually: '23729' -> 'two three seven two nine'."""
    return ' '.join(_DIGIT_WORDS[d] for d in n_str)


def _w(n: int) -> str:
    """num2words output with hyphens removed (forty-two → forty two)."""
    return nw.num2words(n).replace('-', ' ')


def _int_to_spoken(n: int) -> str:
    if n < 100:
        return _w(n)
    if n < 1000:
        h, r = divmod(n, 100)
        return f"{_w(h)} hundred" if r == 0 else f"{_w(h)} {_w(r)}"
    if n < 10000:
        first, last = divmod(n, 100)
        if last == 0:
            return _w(n) if first % 10 == 0 else f"{_w(first)} hundred"
        return f"{_w(first)} {_w(last)}"
    return _w(n)


def expand_numbers(text: str) -> str:
    def _do(seg: str) -> str:
        seg = _ORDINAL_RE.sub(
            lambda m: nw.num2words(int(m.group(1)), to="ordinal").replace('-', ' '), seg)

        def _replace(m):
            n_str = m.group(1)
            # Rule 6: 5+ digit sequences -> digit by digit (e.g. case/petition numbers)
            if len(n_str) >= 5:
                return _digit_by_digit(n_str)
            # Rule 7: 4-digit number preceded by "page" or "para" -> digit by digit
            if len(n_str) == 4:
                words_before = seg[:m.start()].split()
                if words_before and words_before[-1].lower() in ('page', 'para'):
                    return _digit_by_digit(n_str)
            return _int_to_spoken(int(n_str))

        return _NUMBER_RE.sub(_replace, seg)
    return apply_outside_tags(text, _do)


# ── Abbreviation / acronym expansion ─────────────────────────────────────────

def _build_pattern(short: str) -> re.Pattern:
    if short.endswith('.'):
        base = re.escape(short[:-1])
        # \s? handles the space inserted by V1 punct spacing ("Dr ." still matches "Dr.")
        return re.compile(rf'\b{base}\s?\.(?=[\s,;:!?)]|$)', re.IGNORECASE)
    return re.compile(rf'\b{re.escape(short)}\b', re.IGNORECASE)


def load_abbrev_table(tsv_path: Path) -> list[tuple[re.Pattern, str, str]]:
    """Return list of (pattern, expansion, context) sorted longest-short-form first."""
    rows: list[tuple[int, re.Pattern, str, str]] = []
    for line in tsv_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        short     = parts[0].strip()
        expansion = parts[1].strip()
        context   = parts[2].strip().lower() if len(parts) > 2 else "body"
        rows.append((len(short), _build_pattern(short), expansion, context))
    rows.sort(key=lambda x: -x[0])
    return [(pat, exp, ctx) for _, pat, exp, ctx in rows]


def expand_abbreviations(text: str, table: list, is_speaker: bool) -> str:
    for pattern, expansion, context in table:
        if context == "body" and is_speaker:
            continue
        text = apply_outside_tags(text, lambda seg, p=pattern, e=expansion: p.sub(e, seg))
    return text


# ── V1 cleanup helpers ────────────────────────────────────────────────────────

_INITIALS_RE  = re.compile(r'\b(?:[A-Z]\.){2,}')        # A.K. → A K
_LPAREN_RE    = re.compile(r'\(')
_RPAREN_RE    = re.compile(r'\)')
_UNCLEAR_RE   = re.compile(r'\[UNCLEAR\]', re.IGNORECASE)
# Rule 0: non-speech transcript annotations to omit entirely
_OMIT_PHRASES_RE = re.compile(
    r'\[NO\s+AUDIO(?:\s*/\s*VIDEO)?\]'         # [NO AUDIO], [NO AUDIO /VIDEO]
    r'|<<[^>]*>>'                               # <<LUNCH BREAK>>, <<...>>
    r'|END\s+OF\s+THIS\s+PROCEEDINGS?\*?'       # END OF THIS PROCEEDING(S)(*)
    r'|END\s+OF\s+PROCEEDINGS?\*?',             # END OF PROCEEDING(S)(*)
    re.IGNORECASE
)
_REPEAT_PUNCT = re.compile(r'([!?,;])\1+')
_NON_ASCII    = re.compile(r"[^\x00-\x7F‘’\-]")  # keep smart quotes for next step
_SMART_APOS   = re.compile(r"[‘’]")
_MULTI_SPACE  = re.compile(r' {2,}')
_DASH_TABLE   = str.maketrans({'—': '-', '–': '-'})  # em-dash, en-dash


# ── Ellipsis / repeated-dot collapse ─────────────────────────────────────────
_DOTS_BEFORE_PUNCT = re.compile(r'\.{2,}\s*(?=[.?!;:,])')   # I..? → I ?  ; word... . → word .
_DOTS_BEFORE_WORD  = re.compile(r'\.{2,}\s*(?=\S)')         # I... I'm → I , I'm
_DOTS_TRAILING     = re.compile(r'\.{2,}')                  # remaining → remove


def _collapse_dots(text: str) -> str:
    text = _DOTS_BEFORE_PUNCT.sub(' ', text)
    text = _DOTS_BEFORE_WORD.sub(' , ', text)
    text = _DOTS_TRAILING.sub('', text)
    return text


# ── Inline punctuation spacing ────────────────────────────────────────────────
# Comma rules: preserve digit,digit (thousands separators like 32,381 or 3,34,966)
_COMMA_AFTER_ALPHA = re.compile(r'([a-zA-Z]),')       # word,  → word ,
_COMMA_AFTER_DIGIT = re.compile(r'(\d),(?!\d)')        # 100,   → 100 ,  (not 32,000)
_COMMA_BEFORE_ALPHA = re.compile(r',([a-zA-Z])')       # ,word  → , word (not ,5 thousands)
# Period rules
_DOT_ALPHA    = re.compile(r'([a-zA-Z])\.')            # alpha. → alpha .
_DOT_DIGIT_NB = re.compile(r'(\d)\.(?!\d)')            # digit. not before digit (decimal safe)
_DOT_AFTER_RE = re.compile(r'\.([a-zA-Z])')            # .alpha → . alpha


def _space_inline_punct(text: str) -> str:
    """Space commas/fullstops away from adjacent words.
    Preserves: digit.digit decimals (3.14), digit,digit thousands separators (32,381).
    """
    text = _COMMA_AFTER_ALPHA.sub(r'\1 ,', text)
    text = _COMMA_AFTER_DIGIT.sub(r'\1 ,', text)
    text = _COMMA_BEFORE_ALPHA.sub(r', \1', text)
    text = _DOT_ALPHA.sub(r'\1 .', text)
    text = _DOT_DIGIT_NB.sub(r'\1 .', text)
    text = _DOT_AFTER_RE.sub(r'. \1', text)
    return text


def _space_parens(text: str) -> str:
    text = _LPAREN_RE.sub(' ( ', text)
    text = _RPAREN_RE.sub(' ) ', text)
    return text


def _strip_parens(text: str) -> str:
    """Remove parentheses entirely (used in V3)."""
    return text.replace('(', '').replace(')', '')


# ── New V1 normalisation helpers ──────────────────────────────────────────────

# [inaudible] → space (like [UNCLEAR] but silently removed rather than tagged)
_INAUDIBLE_RE = re.compile(r'\[inaudible\]', re.IGNORECASE)

# Rule 0a: any all-caps phrase in square brackets → omit
# e.g. [BREAK], [OFF THE RECORD], [NO VIDEO] etc.
# [UNCLEAR] is converted to <noise> BEFORE this fires, so it is not affected.
_ALL_CAPS_BRACKET_RE = re.compile(r'\[[A-Z][A-Z\s/*-]*\]')

# Number fused with alpha → split: 6a → 6 a, 31c → 31 c
# Ordinal suffixes (st/nd/rd/th) are preserved intact.
_NUM_ALPHA_FUSED_RE = re.compile(r'\b(\d+)([A-Za-z]+)\b')

# Alpha fused with number → split: nng2 → nng 2, under11 → under 11, p1 → p 1
_ALPHA_NUM_FUSED_RE = re.compile(r'\b([A-Za-z]+)(\d+)\b')


def _split_num_alpha(text: str) -> str:
    def _repl(m):
        digits, alpha = m.group(1), m.group(2)
        if alpha.lower() in ('st', 'nd', 'rd', 'th'):
            return m.group(0)  # keep ordinals
        # keep decade/year+s for V3 rule 17: 4-digit+s or 2-digit round-decade+s (90s, 60s)
        if alpha.lower() == 's' and (len(digits) == 4 or (len(digits) == 2 and int(digits) % 10 == 0)):
            return m.group(0)
        return f"{digits} {alpha}"
    return _NUM_ALPHA_FUSED_RE.sub(_repl, text)


def _split_alpha_num(text: str) -> str:
    """Split words that start with letters and contain digits: nng2 → nng 2, p1 → p 1."""
    return _ALPHA_NUM_FUSED_RE.sub(r'\1 \2', text)


# Alpha/alpha slash → space: sc/st → sc st, r/w → r w
_ALPHA_SLASH_RE = re.compile(r'([A-Za-z]+)/([A-Za-z]+)')


def _slash_alpha_to_space(text: str) -> str:
    return _ALPHA_SLASH_RE.sub(r'\1 \2', text)


# Alphabetic word immediately followed by colon → space: Article: → Article
# Negative lookahead on / to avoid breaking http:// style tokens.
_WORD_COLON_RE = re.compile(r'([A-Za-z]+):(?!/)')


def _colon_after_word_to_space(text: str) -> str:
    return _WORD_COLON_RE.sub(r'\1 ', text)


# Words with 2+ hyphens → split all: sadr-e-riyasat → sadr e riyasat
_MULTI_HYPHEN_ALPHA_RE = re.compile(r'\b[A-Za-z]+(?:-[A-Za-z]+){2,}\b', re.IGNORECASE)


def _split_multi_hyphen(text: str) -> str:
    return _MULTI_HYPHEN_ALPHA_RE.sub(lambda m: m.group(0).replace('-', ' '), text)


# Single alpha-alpha hyphen:
#   first part <= 3 chars → join (prefix): non-obstante → nonobstante
#   first part >  3 chars → split (space): north-west   → north west
_SINGLE_HYPHEN_ALPHA_RE = re.compile(r'\b([A-Za-z]+)-([A-Za-z]+)\b', re.IGNORECASE)


def _normalize_single_hyphen(text: str) -> str:
    def _repl(m):
        first, second = m.group(1), m.group(2)
        return first + second if len(first) <= 3 else first + ' ' + second
    return _SINGLE_HYPHEN_ALPHA_RE.sub(_repl, text)


# ── Contraction restoration ───────────────────────────────────────────────────
# Restores apostrophes dropped by PDF extraction (e.g. "I ve" → "I've").
# Applied after space-collapsing so a single space is guaranteed between tokens.
# Uses capturing group to preserve original capitalisation of the main word.

def _csub(suffix: str):
    return lambda m: m.group(1) + suffix


_CONTRACTION_PATTERNS: list[tuple[re.Pattern, object]] = [
    # 've
    (re.compile(r'\b(I)\s+ve\b'),                  _csub("'ve")),
    (re.compile(r'\b(you)\s+ve\b',      re.I),     _csub("'ve")),
    (re.compile(r'\b(we)\s+ve\b',       re.I),     _csub("'ve")),
    (re.compile(r'\b(they)\s+ve\b',     re.I),     _csub("'ve")),
    (re.compile(r'\b(who)\s+ve\b',      re.I),     _csub("'ve")),
    (re.compile(r'\b(would)\s+ve\b',    re.I),     _csub("'ve")),
    (re.compile(r'\b(could)\s+ve\b',    re.I),     _csub("'ve")),
    (re.compile(r'\b(should)\s+ve\b',   re.I),     _csub("'ve")),
    (re.compile(r'\b(might)\s+ve\b',    re.I),     _csub("'ve")),
    (re.compile(r'\b(must)\s+ve\b',     re.I),     _csub("'ve")),
    # 'm
    (re.compile(r'\b(I)\s+m\b'),                   _csub("'m")),
    # 're
    (re.compile(r'\b(you)\s+re\b',      re.I),     _csub("'re")),
    (re.compile(r'\b(we)\s+re\b',       re.I),     _csub("'re")),
    (re.compile(r'\b(they)\s+re\b',     re.I),     _csub("'re")),
    (re.compile(r'\b(who)\s+re\b',      re.I),     _csub("'re")),
    (re.compile(r'\b(there)\s+re\b',    re.I),     _csub("'re")),
    # 'll
    (re.compile(r'\b(I)\s+ll\b'),                  _csub("'ll")),
    (re.compile(r'\b(you)\s+ll\b',      re.I),     _csub("'ll")),
    (re.compile(r'\b(we)\s+ll\b',       re.I),     _csub("'ll")),
    (re.compile(r'\b(they)\s+ll\b',     re.I),     _csub("'ll")),
    (re.compile(r'\b(he)\s+ll\b',       re.I),     _csub("'ll")),
    (re.compile(r'\b(she)\s+ll\b',      re.I),     _csub("'ll")),
    (re.compile(r'\b(it)\s+ll\b',       re.I),     _csub("'ll")),
    (re.compile(r'\b(that)\s+ll\b',     re.I),     _csub("'ll")),
    (re.compile(r'\b(who)\s+ll\b',      re.I),     _csub("'ll")),
    (re.compile(r'\b(there)\s+ll\b',    re.I),     _csub("'ll")),
    # 'd
    (re.compile(r'\b(I)\s+d\b'),                   _csub("'d")),
    (re.compile(r'\b(you)\s+d\b',       re.I),     _csub("'d")),
    (re.compile(r'\b(we)\s+d\b',        re.I),     _csub("'d")),
    (re.compile(r'\b(they)\s+d\b',      re.I),     _csub("'d")),
    (re.compile(r'\b(he)\s+d\b',        re.I),     _csub("'d")),
    (re.compile(r'\b(she)\s+d\b',       re.I),     _csub("'d")),
    (re.compile(r'\b(it)\s+d\b',        re.I),     _csub("'d")),
    (re.compile(r'\b(that)\s+d\b',      re.I),     _csub("'d")),
    (re.compile(r'\b(who)\s+d\b',       re.I),     _csub("'d")),
    # 's
    (re.compile(r'\b(he)\s+s\b',        re.I),     _csub("'s")),
    (re.compile(r'\b(she)\s+s\b',       re.I),     _csub("'s")),
    (re.compile(r'\b(it)\s+s\b',        re.I),     _csub("'s")),
    (re.compile(r'\b(that)\s+s\b',      re.I),     _csub("'s")),
    (re.compile(r'\b(there)\s+s\b',     re.I),     _csub("'s")),
    (re.compile(r'\b(here)\s+s\b',      re.I),     _csub("'s")),
    (re.compile(r'\b(what)\s+s\b',      re.I),     _csub("'s")),
    (re.compile(r'\b(who)\s+s\b',       re.I),     _csub("'s")),
    # n't  (negations)
    (re.compile(r'\b(isn)\s+t\b',       re.I),     _csub("'t")),
    (re.compile(r'\b(aren)\s+t\b',      re.I),     _csub("'t")),
    (re.compile(r'\b(wasn)\s+t\b',      re.I),     _csub("'t")),
    (re.compile(r'\b(weren)\s+t\b',     re.I),     _csub("'t")),
    (re.compile(r'\b(don)\s+t\b',       re.I),     _csub("'t")),
    (re.compile(r'\b(doesn)\s+t\b',     re.I),     _csub("'t")),
    (re.compile(r'\b(didn)\s+t\b',      re.I),     _csub("'t")),
    (re.compile(r'\b(won)\s+t\b',       re.I),     _csub("'t")),
    (re.compile(r'\b(can)\s+t\b',       re.I),     _csub("'t")),
    (re.compile(r'\b(couldn)\s+t\b',    re.I),     _csub("'t")),
    (re.compile(r'\b(wouldn)\s+t\b',    re.I),     _csub("'t")),
    (re.compile(r'\b(shouldn)\s+t\b',   re.I),     _csub("'t")),
    (re.compile(r'\b(haven)\s+t\b',     re.I),     _csub("'t")),
    (re.compile(r'\b(hasn)\s+t\b',      re.I),     _csub("'t")),
    (re.compile(r'\b(hadn)\s+t\b',      re.I),     _csub("'t")),
    (re.compile(r'\b(needn)\s+t\b',     re.I),     _csub("'t")),
    (re.compile(r'\b(mustn)\s+t\b',     re.I),     _csub("'t")),
    (re.compile(r'\b(shan)\s+t\b',      re.I),     _csub("'t")),
]


def _restore_contractions(text: str) -> str:
    for pattern, repl in _CONTRACTION_PATTERNS:
        text = pattern.sub(repl, text)
    return text


# ── V1 cleanup (base layer for all versions) ──────────────────────────────────

def clean_v1(text: str) -> str:
    text = text.translate(_DASH_TABLE)                  # em/en dash → hyphen
    text = _NON_ASCII.sub(' ', text)                    # remove non-ASCII (keep smart quotes)
    text = _SMART_APOS.sub("'", text)                   # smart quotes → straight apostrophe
    text = _INITIALS_RE.sub(                            # A.K. → A K
        lambda m: ' '.join(m.group(0).replace('.', '')), text)
    text = _collapse_dots(text)                         # I... I'm → I , I'm; I..? → I ?
    text = _space_parens(text)                          # 7(2) → 7 ( 2 )
    text = _space_inline_punct(text)                    # Hello, → Hello , ; word.word → word . word
    text = _split_multi_hyphen(text)                    # sadr-e-riyasat → sadr e riyasat (2+ hyphens first)
    text = _normalize_single_hyphen(text)               # non-obstante → nonobstante; north-west → north west
    text = _slash_alpha_to_space(text)                  # sc/st → sc st
    text = _colon_after_word_to_space(text)             # Article: → Article
    text = _split_num_alpha(text)                       # 6a → 6 a, 31c → 31 c
    text = _split_alpha_num(text)                       # nng2 → nng 2, under11 → under 11
    text = _UNCLEAR_RE.sub('<noise>', text)             # [UNCLEAR] → <noise>  (before all-caps bracket)
    text = _OMIT_PHRASES_RE.sub(' ', text)              # rule 0: remove specific non-speech annotations
    text = _ALL_CAPS_BRACKET_RE.sub(' ', text)          # rule 0a: any remaining [ALL CAPS] bracket phrase
    text = _INAUDIBLE_RE.sub(' ', text)                 # [inaudible] → space
    text = _MULTI_SPACE.sub(' ', text).strip()          # collapse spaces before contractions
    text = _restore_contractions(text)                  # I ve → I've
    text = _REPEAT_PUNCT.sub(r'\1', text)               # !! → !
    text = _MULTI_SPACE.sub(' ', text).strip()
    return text


# ── Write ─────────────────────────────────────────────────────────────────────

def write_turns(turns: list[tuple[str, str]], out_path: Path) -> None:
    with out_path.open("w", encoding="utf-8") as f:
        for spk, body in turns:
            f.write(f"{spk}:\n{body}\n\n")


# ── V3-specific expansions ────────────────────────────────────────────────────
# These run inside make_v3 after parens are stripped and before lowercasing.
# They are ordered so that more-specific patterns fire before more-general ones.

# Rule 9: percentage -> "X percent" or "X point Y percent"  (before decimal expansion)
_PERCENT_RE = re.compile(r'\b(\d+)(?:\.(\d+))?%')


def _expand_percent(text: str) -> str:
    def _repl(m):
        int_part = _int_to_spoken(int(m.group(1)))
        if m.group(2):
            frac = ' '.join(_DIGIT_WORDS[d] for d in m.group(2))
            return f"{int_part} point {frac} percent"
        return f"{int_part} percent"
    return apply_outside_tags(text, lambda seg: _PERCENT_RE.sub(_repl, seg))


# Rules 10-11: time with colon -> "X o'clock" (zero mins) or "X Y" (non-zero mins)
_TIME_COLON_RE = re.compile(r'\b(\d{1,2}):(\d{2})\b')


def _expand_time(text: str) -> str:
    def _repl(m):
        hour = int(m.group(1))
        mins = int(m.group(2))
        hw = _int_to_spoken(hour)
        if mins == 0:
            return f"{hw} o'clock"
        return f"{hw} {_int_to_spoken(mins)}"
    return apply_outside_tags(text, lambda seg: _TIME_COLON_RE.sub(_repl, seg))


# Rule 12: slash-separated dates -> spoken components without slash (25/03/71 -> twenty five three seventy one)
_DATE_SLASH_RE = re.compile(r'\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b')


def _expand_date_slash(text: str) -> str:
    def _repl(m):
        return ' '.join(_int_to_spoken(int(g)) for g in m.groups())
    return apply_outside_tags(text, lambda seg: _DATE_SLASH_RE.sub(_repl, seg))


# Rule 15: Roman numerals after "volume", "chapter", "part", or "section"
# e.g. "Volume V" -> "volume five", "Chapter VIII-A" -> "chapter eight a"
_ROMAN_CONTEXT_RE = re.compile(
    r'\b(volume|chapter|part|section)\s+([IVXLCDM]+)(?:-([A-Za-z0-9]+))?\b',
    re.IGNORECASE
)


def _roman_to_int(s: str) -> int:
    vals = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
    n, prev = 0, 0
    for ch in reversed(s.upper()):
        v = vals.get(ch, 0)
        n += v if v >= prev else -v
        prev = v
    return n


def _expand_roman(text: str) -> str:
    def _repl(m):
        prefix    = m.group(1)
        roman_str = m.group(2).upper()
        suffix    = m.group(3)
        n = _roman_to_int(roman_str)
        if n <= 0 or n > 3999:
            return m.group(0)
        result = f"{prefix} {_int_to_spoken(n)}"
        if suffix:
            result += f" {suffix}"
        return result
    return apply_outside_tags(text, lambda seg: _ROMAN_CONTEXT_RE.sub(_repl, seg))


# Rule 16: first 10 Roman numerals always expand (II–X; "I" excluded — too ambiguous with pronoun)
# Ordered longest-first in alternation to prevent partial matches (VIII before VII before VI etc.)
_ROMAN_ALWAYS_RE = re.compile(
    r'\b(VIII|VII|VI|IV|IX|III|II|V|X)\b',
    re.IGNORECASE
)
_ROMAN_ALWAYS_MAP = {
    'II': 2, 'III': 3, 'IV': 4, 'V': 5,
    'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10,
}


def _expand_roman_standalone(text: str) -> str:
    def _repl(m):
        n = _ROMAN_ALWAYS_MAP.get(m.group(1).upper(), 0)
        return _int_to_spoken(n) if n else m.group(0)
    return apply_outside_tags(text, lambda seg: _ROMAN_ALWAYS_RE.sub(_repl, seg))


# Rule 21: hyphenated digit-alpha -> separated by space (2-b -> 2 b; digit expanded later)
_HYPHEN_ALPHANUM_RE = re.compile(r'\b(\d+)-([A-Za-z]+)\b')


def _expand_hyphen_alphanum(text: str) -> str:
    return apply_outside_tags(text, lambda seg: _HYPHEN_ALPHANUM_RE.sub(r'\1 \2', seg))


# Rule 19: strip apostrophes at word edges ('private' -> private; preserves don't / I've)
_EDGE_APOS_RE = re.compile(r"(?<!\w)'|'(?!\w)")


def _strip_edge_apostrophes(text: str) -> str:
    return _EDGE_APOS_RE.sub('', text)


# Rules 13-14: fused digit+unit distance expressions (e.g. 310km, 310.12km -> three hundred and ten kilometre)
# Standalone km/kms without preceding digits handled via TSV (rule 13).
_DIST_FUSED_RE = re.compile(r'\b(\d+)(?:\.(\d+))?\s*(km|kms)\b', re.IGNORECASE)


def _expand_distance(text: str) -> str:
    def _repl(m):
        int_part = _w(int(m.group(1)))   # rule 14: "normal number" 310 -> three hundred and ten
        unit = 'kilometres' if m.group(3).lower() == 'kms' else 'kilometre'
        if m.group(2):
            frac = ' '.join(_DIGIT_WORDS[d] for d in m.group(2))
            return f"{int_part} point {frac} {unit}"
        return f"{int_part} {unit}"
    return apply_outside_tags(text, lambda seg: _DIST_FUSED_RE.sub(_repl, seg))


# Rule 8: decimal numbers -> "X point Y"  (after percent; before general number expansion)
_DECIMAL_NUM_RE = re.compile(r'\b(\d+)\.(\d+)\b')


def _expand_decimal(text: str) -> str:
    def _repl(m):
        int_part = _int_to_spoken(int(m.group(1)))
        frac = ' '.join(_DIGIT_WORDS[d] for d in m.group(2))
        return f"{int_part} point {frac}"
    return apply_outside_tags(text, lambda seg: _DECIMAL_NUM_RE.sub(_repl, seg))


# Rules 1, 4, 5: context-sensitive abbreviations (moved from TSV for digit lookahead)
_CTX_ABBREV = [
    # Art. before digit -> article  (not before alphabetic like Roman numerals)
    (re.compile(r'\bArt\s?\.(?=\s+\d)', re.I), 'article'),
    # No. before digit -> number  (not before alphabetic)
    (re.compile(r'\bNo\s?\.(?=\s+\d)',  re.I), 'number'),
    # para NOT before digit -> paragraph  (para 5 stays as "para"; "in that para" -> "paragraph")
    (re.compile(r'\bpara\b(?!\s*\d)',   re.I), 'paragraph'),
]


def _expand_ctx_abbrev(text: str) -> str:
    for pat, exp in _CTX_ABBREV:
        text = apply_outside_tags(text, lambda seg, p=pat, e=exp: p.sub(e, seg))
    return text


# Rule 17: year/decade followed by 's' → pluralised spoken form
# 1960s → nineteen sixties   90s → nineties   1950s → nineteen fifties
# 4-digit: any year ending in 's'.  2-digit: only round decades (20, 30 … 90).
_YEAR_S_RE = re.compile(r'\b(\d{4}|\d{2})s\b', re.IGNORECASE)


def _pluralize_decade(spoken: str) -> str:
    words = spoken.split()
    last = words[-1]
    if last.endswith('ty'):
        last = last[:-1] + 'ies'   # sixty → sixties, ninety → nineties
    else:
        last += 's'                  # hundred → hundreds, thousand → thousands
    words[-1] = last
    return ' '.join(words)


def _expand_year_s(text: str) -> str:
    def _repl(m):
        digits = m.group(1)
        n = int(digits)
        if len(digits) == 2 and n % 10 != 0:
            return m.group(0)  # not a round decade (e.g. 45s) — leave as-is
        return _pluralize_decade(_int_to_spoken(n))
    return apply_outside_tags(text, lambda seg: _YEAR_S_RE.sub(_repl, seg))


# f2: ensure space before sentence-end punctuation (. ? !)
# After all expansions, any non-space immediately before . ? ! gets a space inserted.
_SENT_END_RE = re.compile(r'([^\s])([.?!])')


def _space_before_sent_punct(text: str) -> str:
    return _SENT_END_RE.sub(r'\1 \2', text)


# Rule 22: singleton & → and (not inside fused tokens like J&K which TSV handles)
_AMP_RE = re.compile(r'(?<!\w)&(?!\w)')


def _expand_ampersand(text: str) -> str:
    return apply_outside_tags(text, lambda seg: _AMP_RE.sub('and', seg))


# Rule 23: replace any remaining hyphens with space (catch-all after specific hyphen rules)
def _replace_remaining_hyphens(text: str) -> str:
    return apply_outside_tags(text, lambda s: s.replace('-', ' '))


# Rule 24: remaining fullstop directly followed by alpha → insert space after it
_REMAINING_DOT_ALPHA_RE = re.compile(r'\.([A-Za-z])')


def _space_dot_alpha(text: str) -> str:
    return apply_outside_tags(text, lambda seg: _REMAINING_DOT_ALPHA_RE.sub(r'. \1', seg))


# f1 (updated): remove commas, semicolons, colons, square brackets, slashes → space
_FINAL_PUNC_RE = re.compile(r'[,;:\[\]/]')


def _remove_final_punct(text: str) -> str:
    return _FINAL_PUNC_RE.sub(' ', text)


# ── Per-version transforms ────────────────────────────────────────────────────

def make_v1(turns: list[tuple[str, str]]) -> list[tuple[str, str]]:
    return [(spk, clean_v1(body)) for spk, body in turns]


def make_v3(turns_v1: list[tuple[str, str]], table: list) -> list[tuple[str, str]]:
    result = []
    for spk, body in turns_v1:
        body = apply_outside_tags(body, _strip_parens)        # rule 20: remove ( )
        body = _expand_ampersand(body)                        # rule 22: & -> and
        body = _expand_ctx_abbrev(body)                       # rules 1,4,5: Art./No./para digit context
        body = _expand_distance(body)                         # rules 13-14: before TSV so "310 km" is caught before km->kilometre
        body = expand_abbreviations(body, table, is_speaker=False)
        body = _expand_percent(body)                          # rule 9: 32% -> thirty two percent
        body = _expand_time(body)                             # rules 10-11: 04:00 -> four o'clock
        body = _expand_date_slash(body)                       # rule 12: 25/03/71 -> twenty five three seventy one
        body = _expand_decimal(body)                          # rule 8: 1.2 -> one point two
        body = _expand_roman(body)                            # rule 15: Chapter/Part/Section VIII-A -> spoken
        body = _expand_roman_standalone(body)                 # rule 16: II-X always -> two-ten
        body = _expand_year_s(body)                           # rule 17: 1960s -> nineteen sixties
        body = _expand_hyphen_alphanum(body)                  # rule 21: 2-b -> 2 b (digit expanded next)
        body = expand_numbers(body)                           # rules 6,7 + base cardinal/ordinal
        body = _replace_remaining_hyphens(body)               # rule 23: any remaining hyphens -> space
        body = apply_outside_tags(body, _strip_edge_apostrophes)  # rule 19: 'word' -> word
        body = apply_outside_tags(body, str.lower)
        body = _space_dot_alpha(body)                         # rule 24: remaining .alpha -> . alpha
        body = apply_outside_tags(body, _remove_final_punct)  # f1: remove , ; : [ ] /
        body = apply_outside_tags(body, _space_before_sent_punct)       # f2: space before . ? !
        body = _TAG_RE.sub(' ', body)                                  # remove <noise> and any tags
        body = _MULTI_SPACE.sub(' ', body).strip()
        result.append((spk, body))
    return result
