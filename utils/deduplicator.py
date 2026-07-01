"""
Speaker deduplication module.

SpeakerDeduplicator ingests a PdfExtractor and applies:
  1. Title / honorific stripping
  2. Phonetic consonant-skeleton comparison (vowel-tolerant)
  3. Union-Find clustering by surname + first name
  4. Wildcard merging (surname-only entries)
  5. Sequential speaker-ID assignment (spk_YXXX)

Module-level functions (strip_titles, parse_name, etc.) are also
importable by reporters and excel_builder for display purposes.
"""

import re
from collections import defaultdict
from difflib import SequenceMatcher
from typing import Optional

from .extractor import PdfExtractor


# ── Title / honorific list ────────────────────────────────────────────────────
TITLES: list[str] = sorted([
    "HON'BLE THE CHIEF JUSTICE OF INDIA",
    "HON'BLE THE CHIEF JUSTICE",
    "HON'BLE MR. JUSTICE", "HON'BLE MS. JUSTICE", "HON'BLE MRS. JUSTICE",
    "HON'BLE JUSTICE",
    "CHIEF JUSTICE OF INDIA", "CHIEF JUSTICE",
    "ATTORNEY GENERAL", "SOLICITOR GENERAL", "ADDITIONAL SOLICITOR GENERAL",
    "JUSTICE", "CJI",
    "MR.", "MRS.", "MS.", "DR.", "MR", "MRS", "MS", "DR", "PROF.", "PROF",
], key=len, reverse=True)

# Generic role labels — not individual persons
ROLE_RE = re.compile(
    r"^(ADVOCATE|COUNSEL|PETITIONER'S COUNSEL \d*|RESPONDENT'S COUNSEL \d*|"
    r"PETITIONER'S COUNSEL|RESPONDENT'S COUNSEL|CLAIMANT'S COUNSEL|NODAL COUNSEL|"
    r"COURT MASTER|BENCH|COURT|INTERPRETER|REGISTRAR|AMICUS|JUDGE|JUSTICE|SPEAKER)$"
)

# Consonant mapping for phonetic skeleton
# V→W, C→K, Q→K, Z→S, P→B  (T→D intentionally excluded — too aggressive)
_CMAP = str.maketrans("VCQZP", "WKKSB")

SUR_THRESHOLD   = 0.80   # surname consonant-key similarity floor
FIRST_THRESHOLD = 0.75   # first-name consonant-key similarity floor
MIN_KEY_LEN     = 3      # keys shorter than this require exact match only


# ── Module-level pure functions (usable from other modules) ───────────────────

def strip_titles(raw: str) -> str:
    """Remove all leading title tokens. Returns uppercase, no dots, collapsed spaces."""
    n = raw.strip().upper()
    changed = True
    while changed:
        changed = False
        for t in TITLES:
            pat = re.compile(r"^" + re.escape(t) + r"\b\.?\s*")
            m = pat.match(n)
            if m and n[m.end():].strip():
                n = n[m.end():].strip()
                changed = True
                break
    n = n.replace(".", " ")
    return re.sub(r"\s+", " ", n).strip()


def is_initial(tok: str) -> bool:
    clean = tok.replace(" ", "")
    return len(clean) <= 2 and clean.isalpha()


def parse_name(raw: str) -> tuple[Optional[str], list[str], Optional[str]]:
    """
    Return (first, middles, surname) after stripping titles.
    Returns (None, [], None) for generic role labels or unresolvable names.
    """
    if ROLE_RE.match(raw.strip().upper()):
        return None, [], None
    norm = strip_titles(raw)
    if not norm:
        return None, [], None
    parts = norm.split()
    sur_idx: Optional[int] = None
    for i in range(len(parts) - 1, -1, -1):
        if not is_initial(parts[i]):
            sur_idx = i
            break
    if sur_idx is None:
        sur_idx = len(parts) - 1
    surname  = parts[sur_idx]
    prefix   = parts[:sur_idx]
    first    = prefix[0] if prefix else None
    middles  = prefix[1:] if len(prefix) > 1 else []
    return first, middles, surname


def ckey(word: str) -> str:
    """Consonant skeleton: uppercase, no dots/spaces, mapped, vowels removed, deduped."""
    if not word:
        return ""
    w = word.upper().replace(".", "").replace(" ", "").replace("'", "")
    w = w.translate(_CMAP)
    w = re.sub(r"[AEIOU]", "", w)
    return re.sub(r"(.)\1+", r"\1", w)


def _sim(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def surname_match(s1: str, s2: str) -> bool:
    if not s1 or not s2:           return False
    if s1 == s2:                   return True
    if s1[0].upper() != s2[0].upper(): return False
    k1, k2 = ckey(s1), ckey(s2)
    if k1 == k2:                   return True
    if len(k1) < MIN_KEY_LEN or len(k2) < MIN_KEY_LEN: return False
    return _sim(k1, k2) >= SUR_THRESHOLD


def first_match(f1: str, f2: str) -> bool:
    if not f1 or not f2: return False
    a = f1.replace(".", "").replace(" ", "").upper()
    b = f2.replace(".", "").replace(" ", "").upper()
    if a == b:                            return True
    if len(a) <= 2 and b.startswith(a):  return True
    if len(b) <= 2 and a.startswith(b):  return True
    if a[0] != b[0]:                      return False
    ka, kb = ckey(a), ckey(b)
    if ka and kb and ka == kb:            return True
    if len(ka) < MIN_KEY_LEN or len(kb) < MIN_KEY_LEN: return False
    return _sim(ka, kb) >= FIRST_THRESHOLD


def speaker_type(name: str) -> str:
    f, _, s = parse_name(name)
    return "Role / Generic" if (f is None and s is None) else "Named Speaker"


def extract_title_and_name(raw: str) -> tuple[str, str, str, str]:
    """Return (title, first_name, middle_names, surname) for display."""
    if ROLE_RE.match(raw.strip().upper()):
        return "", "", "", raw.strip()
    n = raw.strip().upper()
    matched: list[str] = []
    changed = True
    while changed:
        changed = False
        for t in TITLES:
            pat = re.compile(r"^" + re.escape(t) + r"\b\.?\s*")
            m = pat.match(n)
            if m and n[m.end():].strip():
                matched.append(t)
                n = n[m.end():].strip()
                changed = True
                break
    title = " ".join(matched)
    n = re.sub(r"\s+", " ", n.replace(".", " ")).strip()
    if not n:
        return title, "", "", ""
    parts = n.split()
    sur_idx: Optional[int] = None
    for i in range(len(parts) - 1, -1, -1):
        if not is_initial(parts[i]):
            sur_idx = i
            break
    if sur_idx is None:
        sur_idx = len(parts) - 1
    surname = parts[sur_idx]
    prefix  = parts[:sur_idx]
    first   = prefix[0] if prefix else ""
    middles = " ".join(prefix[1:]) if len(prefix) > 1 else ""
    return title, first, middles, surname


# ── Clustering ────────────────────────────────────────────────────────────────

def _cluster(
    parsed_map: dict,
    word_counts: dict,
) -> tuple[list[list[str]], list[tuple[str, list]], list[str]]:
    """
    Union-Find clustering of speaker names.
    Returns (groups, unresolved_wildcards, roles).
      groups    — list of lists of raw names (same canonical speaker)
      wildcards — [(raw_name, [matching_groups])] for ambiguous surname-only entries
      roles     — list of raw names that are generic labels
    """
    roles, wildcards_raw, regulars = [], [], []
    for n, (f, _, s) in parsed_map.items():
        if f is None and s is None:
            roles.append(n)
        elif f is None:
            wildcards_raw.append(n)
        else:
            regulars.append(n)

    parent = {n: n for n in regulars}

    def find(x: str) -> str:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(x: str, y: str) -> None:
        px, py = find(x), find(y)
        if px == py:
            return
        if word_counts.get(px, 0) >= word_counts.get(py, 0):
            parent[py] = px
        else:
            parent[px] = py

    for i, a in enumerate(regulars):
        fa, _, sa = parsed_map[a]
        for b in regulars[i + 1:]:
            fb, _, sb = parsed_map[b]
            if surname_match(sa, sb) and first_match(fa, fb):
                union(a, b)

    clusters: dict[str, list[str]] = defaultdict(list)
    for n in regulars:
        clusters[find(n)].append(n)
    groups = list(clusters.values())

    unresolved = []
    for w in wildcards_raw:
        _, _, ws = parsed_map[w]
        hits = [g for g in groups
                if any(surname_match(ws, parsed_map[n][2]) for n in g)]
        if len(hits) == 1:
            hits[0].append(w)
        else:
            unresolved.append((w, hits))

    return groups, unresolved, roles


# ── Main class ────────────────────────────────────────────────────────────────

class SpeakerDeduplicator:
    """
    Runs the full deduplication pipeline on a PdfExtractor's data.

    Usage:
        dedup = SpeakerDeduplicator(extractor).run()
        print(dedup.canonical_speakers)
    """

    def __init__(self, extractor: PdfExtractor):
        self._extractor = extractor
        self._canonical_speakers: list[str] = []
        self._canon_map:     dict[str, str] = {}
        self._canon_words:   dict[str, int] = {}
        self._canon_files:   dict[str, set] = {}
        self._canon_per_pdf: dict[int, dict[str, int]] = {}
        self._canon_seg_min: dict[str, int] = {}
        self._canon_seg_max: dict[str, int] = {}
        self._canon_spk_id:  dict[str, str] = {}
        self._raw_spk_id:    dict[str, str] = {}
        self._groups:        list[list[str]] = []
        self._wildcards:     list            = []
        self._roles:         list[str]       = []

    def run(self) -> "SpeakerDeduplicator":
        ext   = self._extractor
        gwt   = ext.global_word_totals    # {raw_name: word_count}
        gsl   = ext.seg_lengths           # {raw_name: [seg_wc, ...]}

        # Cluster names
        parsed_map = {n: parse_name(n) for n in gwt}
        groups, wildcards, roles = _cluster(parsed_map, gwt)

        # Sort: groups by total words desc; within each group by member word count desc
        groups.sort(key=lambda g: -sum(gwt.get(n, 0) for n in g))
        for g in groups:
            g.sort(key=lambda n: -gwt.get(n, 0))

        self._groups    = groups
        self._wildcards = wildcards
        self._roles     = roles

        # Build variant → canonical map
        canon_map: dict[str, str] = {}
        for g in groups:
            canonical = g[0]
            for n in g:
                canon_map[n] = canonical
        self._canon_map = canon_map

        # Aggregate stats by canonical name
        canon_words:   dict[str, int]        = defaultdict(int)
        canon_files:   dict[str, set]        = defaultdict(set)
        canon_per_pdf: dict[int, dict]       = defaultdict(lambda: defaultdict(int))
        canon_seg_min: dict[str, int]        = {}
        canon_seg_max: dict[str, int]        = {}

        for pdf_idx, counts in ext.raw_data.items():
            for raw_name, wc in counts.items():
                canon = canon_map.get(raw_name, raw_name)
                canon_words[canon]              += wc
                canon_files[canon].add(pdf_idx)
                canon_per_pdf[pdf_idx][canon]   += wc

        for raw_name, lengths in gsl.items():
            canon = canon_map.get(raw_name, raw_name)
            lo, hi = min(lengths), max(lengths)
            if canon not in canon_seg_min:
                canon_seg_min[canon] = lo
                canon_seg_max[canon] = hi
            else:
                canon_seg_min[canon] = min(canon_seg_min[canon], lo)
                canon_seg_max[canon] = max(canon_seg_max[canon], hi)

        self._canon_words   = dict(canon_words)
        self._canon_files   = dict(canon_files)
        self._canon_per_pdf = {k: dict(v) for k, v in canon_per_pdf.items()}
        self._canon_seg_min = canon_seg_min
        self._canon_seg_max = canon_seg_max

        # Order canonical speakers by total word count desc
        canonical_speakers = sorted(canon_words.keys(), key=lambda n: -canon_words[n])
        self._canonical_speakers = canonical_speakers

        # Assign speaker IDs  spk_YXXX
        canon_spk_id: dict[str, str] = {}
        for seq, canon in enumerate(canonical_speakers, start=1):
            y = "1" if speaker_type(canon) == "Named Speaker" else "0"
            canon_spk_id[canon] = f"spk_{y}{seq:03d}"
        self._canon_spk_id = canon_spk_id

        # Map every raw name to its canonical's spk_id
        raw_spk_id: dict[str, str] = {}
        for raw_name in gwt:
            canon = canon_map.get(raw_name, raw_name)
            raw_spk_id[raw_name] = canon_spk_id[canon]
        self._raw_spk_id = raw_spk_id

        print(f"Deduplication complete: {len(canonical_speakers)} canonical speakers "
              f"({len(groups)} groups, {len(roles)} roles, "
              f"{len(wildcards)} unresolved wildcards)")
        return self

    # ── Properties ────────────────────────────────────────────────────────────

    @property
    def canonical_speakers(self) -> list[str]:   return self._canonical_speakers
    @property
    def canon_map(self) -> dict[str, str]:        return self._canon_map
    @property
    def canon_words(self) -> dict[str, int]:      return self._canon_words
    @property
    def canon_files(self) -> dict[str, set]:      return self._canon_files
    @property
    def canon_per_pdf(self) -> dict:              return self._canon_per_pdf
    @property
    def canon_seg_min(self) -> dict[str, int]:    return self._canon_seg_min
    @property
    def canon_seg_max(self) -> dict[str, int]:    return self._canon_seg_max
    @property
    def canon_spk_id(self) -> dict[str, str]:     return self._canon_spk_id
    @property
    def raw_spk_id(self) -> dict[str, str]:       return self._raw_spk_id
    @property
    def groups(self) -> list[list[str]]:          return self._groups
    @property
    def wildcards(self):                           return self._wildcards
    @property
    def roles(self) -> list[str]:                 return self._roles
