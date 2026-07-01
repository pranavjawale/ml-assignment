# Analysis Pipeline — Supreme Court ASR Transcripts

A single command-line tool (`analysis-pipeline.py`) that turns Supreme Court
oral-hearing **PDF transcripts** into clean, speaker-labelled, ASR-ready text
plus a set of analysis workbooks.

It extracts text from the official TERES-format PDFs, deduplicates the many
spelling/title variants of each speaker into stable `spk_NNNN` IDs, produces
written-form (**V1**) and spoken-form (**V3**) normalized transcripts, and emits
Excel reports on speaker statistics, content categories, and non-standard words
(NSW).

---

## 1. Setup (first-time clone)

From this directory (`text-extraction/analysis-pipeline/`):

```bash
python -m venv .venv
.venv\Scripts\activate            # Windows
source .venv/bin/activate         # macOS / Linux

pip install -r requirements.txt   # installs everything, incl. the spaCy model
```

- Requires **Python 3.10+** (the code uses `X | None` / `list[...]` type hints).
- `requirements.txt` pins the spaCy English model (`en_core_web_sm`) directly, so
  a single `pip install` is enough — no separate `spacy download` step.

Quick check:

```bash
python analysis-pipeline.py --help
```

---

## 2. Core concepts

| Concept | Meaning |
|---|---|
| **`--tag`** | A short label (e.g. `24-file-data`) prepended to every output file and directory, so reruns/datasets don't overwrite each other. Required by all modes. |
| **`analysis-outputs/`** | The single output directory (created automatically) where everything is written, all `<tag>`-prefixed. |
| **Transcript versions** | `transcript_original` = raw PDF speaker names. **V1** = `spk_NNNN` labels + written-form cleanup. **V3** = V1 + spoken-form expansion (`370` → `three seventy`, `Dr.` → `doctor`). |
| **Speaker IDs** | `spk_1XXX` = named individuals, `spk_0XXX` = role labels (COUNSEL, COURT MASTER…). Assigned by total-words rank. |

### Data flow

```
                 ┌─ --pdf-summary ───────────────► <tag>_pdf-analysis-summary.txt
                 │
   transcript    ├─ --extract-trans ─► <tag>_transcript_original/ ─┐
     PDFs  ──────┤                                                  │
                 │                          ┌── --content-stats ◄───┤
                 │                          ├── --dedup-names   ◄───┤
                 └─ --get-norm-trans ◄──────┘  (reads original/)    │
                        │                                            │
                        ├─ writes V1 + V3 transcripts ───────────────┘
                        ▼
                 --align-and-review   (original vs V3, NSW review)
                 --align-and-review2  (V1 vs V3, NSW review)
```

- `--get-norm-trans --pdf-dir …` does **everything** from the PDFs in one shot.
- `--content-stats`, `--dedup-names`, and `--get-norm-trans --tag` (no `--pdf-dir`)
  read the already-extracted `<tag>_transcript_original/`.

---

## 3. Dependencies

| Package | Used by |
|---|---|
| **PyMuPDF** (`fitz`) | PDF text/font extraction |
| **openpyxl** | all `.xlsx` outputs |
| **spaCy** + `en_core_web_sm` | sentence/POS analysis in `--content-stats` |
| **num2words** | V3 spoken-form number expansion |

Everything else is Python standard library.

---

## 4. Modes (flags)

Exactly one mode runs per invocation. The PDF directory used below is
`../transcript-pdfs` (relative to this folder).

### At a glance — what each flag does

| Flag | What it does |
|---|---|
| `--pdf-summary` | Structural metadata report on the PDFs (pages, words, speakers, fonts, noise lines). No transcripts. |
| `--extract-trans` | Extract raw transcripts from the PDFs (speaker names as-is, no ID resolution). Feeds the analysis modes. |
| `--content-stats` | spaCy content analysis of the extracted transcripts → sentence/character/word/NSW-frequency workbook. |
| `--dedup-names` | Deduplicate speaker-name variants into stable `spk_NNNN` IDs → mapping + speaker-analysis workbook. |
| `--get-norm-trans` | Full pipeline: dedup + spk_NNNN-mapped **V1** (written) and **V3** (spoken-form) transcripts + word-freq + content-stats (+ PDF summary when run with `--pdf-dir`). |
| `--align-and-review` | NSW review: align `transcript_original` vs **V3**, classify each difference by content-stats NSW category. Main review. |
| `--align-and-review2` | Same as above but the written side is **V1** instead of `transcript_original`. |

### Inputs & outputs per flag

All outputs land in `analysis-outputs/` and are `<tag>`-prefixed. "Input dirs" are
also under `analysis-outputs/` (produced by an earlier run).

| Flag | Required args | Reads | Writes |
|---|---|---|---|
| `--pdf-summary` | `--pdf-dir`, `--tag` | the PDFs | `<tag>_pdf-analysis-summary.txt` |
| `--extract-trans` | `--pdf-dir`, `--tag` | the PDFs | `<tag>_transcript_original/` |
| `--content-stats` | `--tag` | `<tag>_transcript_original/` | `<tag>_transcript_content_analysis.xlsx` |
| `--dedup-names` | `--tag` | `<tag>_transcript_original/` | `<tag>_spk-id-mapping.txt`, `<tag>_speaker-name-analysis.xlsx` |
| `--get-norm-trans` | `--tag` (`--pdf-dir` optional) | the PDFs (or `<tag>_transcript_original/` if no `--pdf-dir`) | `<tag>_transcript_original/`, `<tag>_spkid_mapped_transcript_v1/`, `<tag>_spkid_mapped_norm_transcript_v3/`, `<tag>_spk-id-mapping.txt`, `<tag>_speaker-name-analysis.xlsx`, `<tag>_norm_trans_word_freq.xlsx`, `<tag>_transcript_content_analysis.xlsx`, `<tag>_pdf-analysis-summary.txt` (PDF mode only) |
| `--align-and-review` | `--tag`, `--file-id` | `<tag>_transcript_original/` + `<tag>_spkid_mapped_norm_transcript_v3/` | `<tag>_nsw-review-<id>.xlsx` |
| `--align-and-review2` | `--tag`, `--file-id` | `<tag>_spkid_mapped_transcript_v1/` + `<tag>_spkid_mapped_norm_transcript_v3/` | `<tag>_nsw-review2-<id>.xlsx` |

Prerequisites: `--content-stats` / `--dedup-names` need `--extract-trans` (or
`--get-norm-trans`) run first; the two `--align-and-review*` flags need
`--get-norm-trans` run first.

### 4.1 `--pdf-summary`
Structural metadata report on the PDFs (no transcript output).

```bash
python analysis-pipeline.py --pdf-summary --pdf-dir ../transcript-pdfs --tag 24-file-data
```
| | |
|---|---|
| **Requires** | `--pdf-dir`, `--tag` |
| **Output** | `<tag>_pdf-analysis-summary.txt` — page/word/speaker counts, per-file table, font analysis, noise-line counts |

### 4.2 `--extract-trans`
Raw transcripts only — speaker names exactly as in the PDF, no ID resolution.
This is the feeder for the analysis modes.

```bash
python analysis-pipeline.py --extract-trans --pdf-dir ../transcript-pdfs --tag 24-file-data
```
| | |
|---|---|
| **Requires** | `--pdf-dir`, `--tag` |
| **Output** | `<tag>_transcript_original/transcript_NNN.txt` (one file per PDF) |

### 4.3 `--content-stats`
spaCy content analysis of the extracted transcripts.

```bash
python analysis-pipeline.py --content-stats --tag 24-file-data
```
| | |
|---|---|
| **Requires** | `--tag` (run `--extract-trans` first) |
| **Reads** | `<tag>_transcript_original/` |
| **Output** | `<tag>_transcript_content_analysis.xlsx` — 5 sheets: Overview (speaker count + sentence complete/partial stats), Category Counts, Character Frequency, Word Frequency, NSW Frequency (Bracketed / Acronym / LSEQ / Abbreviation / SPLT / Cardinal / Ordinal / Number-range / Time / Date / Percentage / Money / Telephone / Email / Hashtag) |

### 4.4 `--dedup-names`
Speaker-name deduplication only.

```bash
python analysis-pipeline.py --dedup-names --tag 24-file-data
```
| | |
|---|---|
| **Requires** | `--tag` (run `--extract-trans` first) |
| **Reads** | `<tag>_transcript_original/` |
| **Outputs** | `<tag>_spk-id-mapping.txt` (raw name → `spk_NNNN`) · `<tag>_speaker-name-analysis.xlsx` (SHEET1 stats, SHEET2 speaker×file matrix, SHEET3 dedup review) |

### 4.5 `--get-norm-trans`
The full pipeline. **`--pdf-dir` is optional**:

- **with `--pdf-dir`** — extract from PDFs, then produce *all* outputs.
- **without `--pdf-dir`** — reuse the existing `<tag>_transcript_original/` and
  produce everything except `pdf-analysis-summary.txt` (which needs the PDFs).

```bash
# full, from PDFs
python analysis-pipeline.py --get-norm-trans --pdf-dir ../transcript-pdfs --tag 24-file-data

# analysis-only, on already-extracted transcripts
python analysis-pipeline.py --get-norm-trans --tag 24-file-data
```
| | |
|---|---|
| **Requires** | `--tag` (`--pdf-dir` optional) |
| **Outputs** | `<tag>_transcript_original/` (PDF mode only) · `<tag>_spkid_mapped_transcript_v1/` (V1) · `<tag>_spkid_mapped_norm_transcript_v3/` (V3 spoken form) · `<tag>_spk-id-mapping.txt` · `<tag>_speaker-name-analysis.xlsx` · `<tag>_norm_trans_word_freq.xlsx` · `<tag>_transcript_content_analysis.xlsx` · `<tag>_pdf-analysis-summary.txt` (PDF mode only) |

### 4.6 `--align-and-review` (main NSW review)
For each `--file-id`, align the **written** transcript against the **spoken (V3)**
form token-by-token and classify every difference into the content-stats NSW
categories — a quick way to review/QA the V3 normalization.

The **written side is `transcript_original`** (original case + dots preserved),
cleaned only for ellipsis/punctuation so NSW forms survive: dotted initials
`A.K.` → **LSEQ**, abbreviations `Dr.` → **Abbreviation**, `[UNCLEAR]` →
**Bracketed**. (This is why original — not V1 — is the main choice: V1's cleanup
strips the dots that LSEQ/Abbreviation detection needs.)

```bash
python analysis-pipeline.py --align-and-review --tag 24-file-data --file-id 001,003
```
| | |
|---|---|
| **Requires** | `--tag`, `--file-id` (comma-separated ids, e.g. `001,003`; run `--get-norm-trans` first) |
| **Reads** | `<tag>_transcript_original/` and `<tag>_spkid_mapped_norm_transcript_v3/` |
| **Output** | `<tag>_nsw-review-<id>.xlsx` per file — a **SUMMARY** sheet (category → count, each name is a **clickable hyperlink** to its sheet) + one sheet per NSW category (Bracketed, Email, Hashtag, Money, Percentage, Telephone, Time, Date, Number range, Ordinal, LSEQ, SPLT, Acronym, Cardinal, Abbreviation, Other) with written→spoken pairs and highlighted context |

### 4.7 `--align-and-review2` (V1 variant)
Identical to `--align-and-review` — same NSW categories, same hyperlinked SUMMARY
— but the **written side is V1** (`<tag>_spkid_mapped_transcript_v1/`) instead of
`transcript_original`. Useful for comparison; note V1 flattens `A.K.` → `A K`, so
**LSEQ ≈ 0** and the `Other` bucket is larger (V1 keeps spaced punctuation).

```bash
python analysis-pipeline.py --align-and-review2 --tag 24-file-data --file-id 001,003
```
| | |
|---|---|
| **Requires** | `--tag`, `--file-id` (run `--get-norm-trans` first) |
| **Reads** | `<tag>_spkid_mapped_transcript_v1/` and `<tag>_spkid_mapped_norm_transcript_v3/` |
| **Output** | `<tag>_nsw-review2-<id>.xlsx` per file (same structure as 4.6) |

---

## 5. Configuration files (edit these, not the code)

| File | Purpose |
|---|---|
| `speaker-aliases.tsv` | Manual `raw_name → canonical` fixes applied **before** dedup (OCR errors, missing spaces, same-person variants). 3 columns: `raw_name`, `canonical`, `category`. |
| `abbrev-expansions.tsv` | Abbreviation/acronym → spoken-form table used by V3 (e.g. `Dr.` → `doctor`). |

---

## 6. Typical end-to-end workflow

```bash
# One shot: PDFs → everything
python analysis-pipeline.py --get-norm-trans --pdf-dir ../transcript-pdfs --tag 24-file-data

# Then review specific files' written→spoken normalization
python analysis-pipeline.py --align-and-review --tag 24-file-data --file-id 001,003
```

Or step by step:

```bash
python analysis-pipeline.py --extract-trans  --pdf-dir ../transcript-pdfs --tag 24-file-data
python analysis-pipeline.py --dedup-names    --tag 24-file-data    # inspect spk-id-mapping
python analysis-pipeline.py --content-stats  --tag 24-file-data
python analysis-pipeline.py --get-norm-trans --tag 24-file-data    # V1 / V3 + word freq
python analysis-pipeline.py --align-and-review --tag 24-file-data --file-id 001
```

> Excel note: close any open `.xlsx` before re-running, or Windows raises a
> `PermissionError` on save.

---

## 7. Output reference (`analysis-outputs/`)

| Output | Produced by |
|---|---|
| `<tag>_pdf-analysis-summary.txt` | `--pdf-summary`, `--get-norm-trans` (PDF mode) |
| `<tag>_transcript_original/` | `--extract-trans`, `--get-norm-trans` (PDF mode) |
| `<tag>_spk-id-mapping.txt` | `--dedup-names`, `--get-norm-trans` |
| `<tag>_speaker-name-analysis.xlsx` | `--dedup-names`, `--get-norm-trans` |
| `<tag>_transcript_content_analysis.xlsx` | `--content-stats`, `--get-norm-trans` |
| `<tag>_spkid_mapped_transcript_v1/` | `--get-norm-trans` |
| `<tag>_spkid_mapped_norm_transcript_v3/` | `--get-norm-trans` |
| `<tag>_norm_trans_word_freq.xlsx` | `--get-norm-trans` |
| `<tag>_nsw-review-<id>.xlsx` | `--align-and-review` (original vs V3) |
| `<tag>_nsw-review2-<id>.xlsx` | `--align-and-review2` (V1 vs V3) |

---

## 8. Notes & known limitations

- **Speaker ID mapping may change for different set of inputs** `spk_NNNN` is the speaker's
  rank by total words spoken. For identical input it is fully reproducible, but
  adding/removing transcripts or editing `speaker-aliases.tsv` can shift a few
  mid-ranked IDs. The two `--get-norm-trans` paths (PDF vs tag-only) also count
  words from slightly different sources, so a few IDs can differ between them.
  (A persistent ID registry would make them stable across datasets.)
- **Excel files must be closed** before re-running, or Windows raises a
  `PermissionError` on save.
- **`--content-stats` is the slow step** (~45 s; spaCy parses every utterance);
  the others run in seconds.
- **NSW review — original vs V1.** `--align-and-review` (original written side)
  detects LSEQ/Abbreviations and keeps the `Other` bucket small; `--align-and-review2`
  (V1 written side) is the same in every other respect but loses LSEQ and has a
  larger `Other`, because V1's `clean_v1` strips dotted initials and spaces punctuation.

---

## 9. Repo layout (this directory)

```
analysis-pipeline.py            # main entry point (all 7 modes; imports only from utils/)
README.md                       # this file
speaker-aliases.tsv             # manual speaker-name fixes
abbrev-expansions.tsv           # norm trans abbreviation expansions
requirements.txt                # dependencies
transcript-pdfs/                # input PDFs (transcript_NNN.pdf) — pass as --pdf-dir
utils/
  __init__.py                   # marks utils/ as a package
  extractor.py                  # PdfExtractor (PDF → speaker word counts)
  transcript_extractor.py       # TranscriptExtractor (TXT → speaker word counts)
  deduplicator.py               # SpeakerDeduplicator (clustering + spk_NNNN IDs)
  normalizer.py                 # V1/V3 transcript normalization (make_v1 / make_v3)
  nsw_align.py                  # V1↔V3 token alignment + NSW classify + review Excel writers
  excel_builder.py              # speaker-name-analysis.xlsx (3 sheets)
  reporters.py                  # pdf-analysis-summary.txt
  content_stats.py              # spaCy content analysis (5-sheet workbook)
  aliases.py                    # loads speaker-aliases.tsv as RAW_ALIASES
analysis-outputs/               # all outputs land here (created on first run)
```
