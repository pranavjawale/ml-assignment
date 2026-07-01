"""
PDF extraction module.

PdfExtractor reads all transcript PDFs and produces:
  - per-PDF word counts and turn counts per speaker
  - per-speaker segment length lists (for min/max stats)
  - per-PDF structural metadata (fonts, noise counts, page count)
"""

import fitz
import re
import os
from collections import defaultdict

from .aliases import RAW_ALIASES


class PdfExtractor:
    RE_SPEAKER   = re.compile(r"^([A-Z][A-Z0-9 .\-']+?):\s*(.*)")
    RE_TERES     = re.compile(r"^\s*Transcribed by TERES\s*$")
    RE_INTEGER   = re.compile(r"^\s*\d+\s*$")
    RE_TIMESTAMP = re.compile(r"^\s*\d{1,2}:\d{2}\s*[AP]M\s*IST\s*$")
    RE_EMAIL     = re.compile(r"^\s*\S+@\S+\.\S+\s*$")
    # Speaker names containing characters other than A-Z, space, apostrophe, period
    RE_UNUSUAL   = re.compile(r"[^A-Z '.]")

    def __init__(self, pdf_dir: str):
        self.pdf_dir = pdf_dir
        # {pdf_idx: {speaker_name: word_count}}
        self._raw_data: dict[int, dict[str, int]] = {}
        # {speaker_name: [segment_word_counts]}
        self._seg_lengths: dict[str, list[int]] = defaultdict(list)
        # {pdf_idx: metadata_dict}
        self._pdf_meta: dict[int, dict] = {}
        # raw speaker names (pre-alias) that contain chars outside A-Z, space, period
        self._unusual_speakers: set[str] = set()
        # character frequency across all spoken text (not speaker labels)
        self._char_freq: dict[str, int] = defaultdict(int)

    # ── Public API ────────────────────────────────────────────────────────────

    def extract_all(self) -> "PdfExtractor":
        pdfs = sorted(
            f for f in os.listdir(self.pdf_dir)
            if f.startswith("transcript_") and f.endswith(".pdf")
        )
        print(f"Extracting {len(pdfs)} PDFs...")
        for fname in pdfs:
            idx = int(fname.replace("transcript_", "").replace(".pdf", ""))
            path = os.path.join(self.pdf_dir, fname)
            words, _turns, seg_lens, meta, unusual, chars = self._extract_one(path)
            self._raw_data[idx]  = words
            self._pdf_meta[idx]  = meta
            for name, lengths in seg_lens.items():
                self._seg_lengths[name].extend(lengths)
            self._unusual_speakers.update(unusual)
            for ch, cnt in chars.items():
                self._char_freq[ch] += cnt
            print(f"  [{idx:02d}] {fname}  — {len(words)} speakers, "
                  f"{meta['word_count']:,} words, {meta['turn_count']} turns")
        return self

    @property
    def raw_data(self) -> dict[int, dict[str, int]]:
        """Word counts: {pdf_idx: {speaker: word_count}}"""
        return self._raw_data

    @property
    def seg_lengths(self) -> dict[str, list[int]]:
        """Segment word-count lists: {speaker: [wc1, wc2, ...]}"""
        return self._seg_lengths

    @property
    def pdf_meta(self) -> dict[int, dict]:
        return self._pdf_meta

    @property
    def global_word_totals(self) -> dict[str, int]:
        totals: dict[str, int] = defaultdict(int)
        for counts in self._raw_data.values():
            for name, wc in counts.items():
                totals[name] += wc
        return dict(totals)

    @property
    def global_file_sets(self) -> dict[str, set[int]]:
        """Files (pdf indices) in which each raw speaker appears."""
        result: dict[str, set[int]] = defaultdict(set)
        for idx, counts in self._raw_data.items():
            for name in counts:
                result[name].add(idx)
        return dict(result)

    @property
    def unusual_speakers(self) -> set[str]:
        """Raw speaker names (pre-alias) containing chars outside A-Z, space, period."""
        return self._unusual_speakers

    @property
    def char_freq(self) -> dict[str, int]:
        """Character frequency across all spoken text (not speaker labels)."""
        return dict(self._char_freq)

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _is_noise(self, line: str) -> tuple[bool, str]:
        """Return (is_noise, noise_type). noise_type is '' if not noise."""
        if self.RE_TERES.match(line):     return True, "teres"
        if self.RE_INTEGER.match(line):   return True, "line_num"
        if self.RE_TIMESTAMP.match(line): return True, "timestamp"
        if self.RE_EMAIL.match(line):     return True, "email"
        return False, ""

    def _extract_one(self, pdf_path: str) -> tuple[dict, dict, dict, dict, set, dict]:
        """
        Process one PDF.
        Returns (word_counts, turn_counts, seg_lengths, meta, unusual_names, char_freq).
        """
        doc = fitz.open(pdf_path)
        word_counts  = defaultdict(int)
        turn_counts  = defaultdict(int)
        seg_lengths  = defaultdict(list)
        noise_counts = {"teres": 0, "line_num": 0, "timestamp": 0, "email": 0}
        unusual_names: set[str] = set()
        char_freq: dict[str, int] = defaultdict(int)
        cur, cur_seg_wc = None, 0

        for pi in range(1, len(doc)):
            for raw_line in doc[pi].get_text("text").splitlines():
                line = raw_line.strip()
                if not line:
                    continue
                noisy, ntype = self._is_noise(line)
                if noisy:
                    noise_counts[ntype] += 1
                    continue
                m = self.RE_SPEAKER.match(line)
                if m:
                    if cur and cur_seg_wc > 0:
                        seg_lengths[cur].append(cur_seg_wc)
                    raw_name = m.group(1).strip()
                    if self.RE_UNUSUAL.search(raw_name):
                        unusual_names.add(raw_name)
                    cur = RAW_ALIASES.get(raw_name, raw_name)
                    cur_seg_wc = 0
                    turn_counts[cur] += 1
                    rest = m.group(2).strip()
                    if rest:
                        wc = len(rest.split())
                        word_counts[cur] += wc
                        cur_seg_wc += wc
                        for ch in rest:
                            char_freq[ch] += 1
                elif cur:
                    wc = len(line.split())
                    word_counts[cur] += wc
                    cur_seg_wc += wc
                    for ch in line:
                        char_freq[ch] += 1

        if cur and cur_seg_wc > 0:
            seg_lengths[cur].append(cur_seg_wc)

        meta = {
            "filename":      os.path.basename(pdf_path),
            "page_count":    len(doc),
            "speaker_count": len(word_counts),
            "word_count":    sum(word_counts.values()),
            "turn_count":    sum(turn_counts.values()),
            "noise_counts":  noise_counts,
            "cover_fonts":   self._get_page_fonts(doc[0]) if len(doc) > 0 else set(),
            "body_fonts":    self._get_body_fonts(doc),
        }
        doc.close()
        return dict(word_counts), dict(turn_counts), dict(seg_lengths), meta, unusual_names, dict(char_freq)

    def _get_page_fonts(self, page) -> set[str]:
        fonts: set[str] = set()
        try:
            for block in page.get_text("rawdict").get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        fname = span.get("font", "")
                        fsize = round(span.get("size", 0))
                        if fname:
                            fonts.add(f"{fname} ({fsize}pt)")
        except Exception:
            pass
        return fonts

    def _get_body_fonts(self, doc) -> set[str]:
        fonts: set[str] = set()
        for pi in range(1, min(4, len(doc))):   # sample first 3 body pages
            fonts.update(self._get_page_fonts(doc[pi]))
        return fonts
