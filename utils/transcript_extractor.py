"""
Transcript-based speaker extractor.

Reads transcript_NNN.txt files (the raw speaker-name transcripts produced by
--extract-trans) and exposes the same read interface that SpeakerDeduplicator and
build_speaker_excel expect from PdfExtractor:

    raw_data            {file_idx: {raw_name: word_count}}
    seg_lengths         {raw_name: [segment_word_count, ...]}
    global_word_totals  {raw_name: total_word_count}
    global_file_sets    {raw_name: {file_idx, ...}}

Speaker names are alias-normalised via RAW_ALIASES at read time, mirroring
PdfExtractor (extractor.py), so the deduplication input is identical whether it
is sourced from the PDFs directly or from the extracted transcripts.
"""

import os
import re
from collections import defaultdict

from .aliases import RAW_ALIASES

# Speaker-label line in an extracted transcript:  "JUSTICE D Y CHANDRACHUD:"
_SPEAKER_RE = re.compile(r"^([A-Z][A-Z0-9 .\-']+):\s*$")


class TranscriptExtractor:
    """Duck-typed stand-in for PdfExtractor, sourced from transcript TXT files."""

    def __init__(self, transcript_dir: str):
        self.transcript_dir = transcript_dir
        self._raw_data:    dict[int, dict[str, int]] = {}
        self._seg_lengths: dict[str, list[int]]      = defaultdict(list)

    def extract_all(self) -> "TranscriptExtractor":
        files = sorted(
            f for f in os.listdir(self.transcript_dir)
            if f.startswith("transcript_") and f.endswith(".txt")
        )
        if not files:
            raise SystemExit(
                f"ERROR: no transcript_*.txt files found in {self.transcript_dir}"
            )

        print(f"\nReading {len(files)} transcript files from {self.transcript_dir}")
        print("-" * 55)

        for fname in files:
            idx  = int(fname.replace("transcript_", "").replace(".txt", ""))
            path = os.path.join(self.transcript_dir, fname)

            counts:  dict[str, int] = defaultdict(int)
            cur:     str | None     = None
            cur_seg: int            = 0

            with open(path, encoding="utf-8") as fh:
                for raw in fh:
                    line = raw.strip()
                    m = _SPEAKER_RE.match(line)
                    if m:
                        # close previous turn's segment
                        if cur is not None and cur_seg:
                            self._seg_lengths[cur].append(cur_seg)
                        cur     = RAW_ALIASES.get(m.group(1).strip(), m.group(1).strip())
                        cur_seg = 0
                    elif line and cur is not None:
                        n = len(line.split())
                        counts[cur] += n
                        cur_seg     += n
                if cur is not None and cur_seg:
                    self._seg_lengths[cur].append(cur_seg)

            self._raw_data[idx] = dict(counts)
            print(f"  [{idx:02d}] {fname}  ->  {len(counts)} speakers")

        print(f"\nTotal raw speaker strings: {len(self.global_word_totals)}")
        return self

    # ── Properties (mirror PdfExtractor) ──────────────────────────────────────

    @property
    def raw_data(self) -> dict[int, dict[str, int]]:
        return self._raw_data

    @property
    def seg_lengths(self) -> dict[str, list[int]]:
        return dict(self._seg_lengths)

    @property
    def global_word_totals(self) -> dict[str, int]:
        totals: dict[str, int] = defaultdict(int)
        for counts in self._raw_data.values():
            for name, wc in counts.items():
                totals[name] += wc
        return dict(totals)

    @property
    def global_file_sets(self) -> dict[str, set]:
        sets: dict[str, set] = defaultdict(set)
        for idx, counts in self._raw_data.items():
            for name in counts:
                sets[name].add(idx)
        return dict(sets)
