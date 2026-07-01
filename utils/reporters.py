"""
Text report writers.

Produces:
  pdf-analysis-summary.txt  — structural analysis of all PDFs

(The dedup candidate review is now SHEET3 of the speaker-name-analysis workbook;
see utils/excel_builder.py._build_sheet3.)
"""

from .extractor import PdfExtractor


# ── PDF Analysis Summary ──────────────────────────────────────────────────────

def write_pdf_analysis_summary(
    extractor: PdfExtractor,
    output_path: str,
) -> None:
    """
    Structural analysis of all PDFs:
    page counts, fonts (cover vs body), noise element statistics,
    per-file summary table, and extraction approach notes.
    """
    meta_all = extractor.pdf_meta
    if not meta_all:
        print("Warning: no PDF metadata available — run extractor.extract_all() first.")
        return

    pdf_indices = sorted(meta_all.keys())
    page_counts = [meta_all[i]["page_count"] for i in pdf_indices]
    word_counts = [meta_all[i]["word_count"]  for i in pdf_indices]
    turn_counts = [meta_all[i]["turn_count"]  for i in pdf_indices]
    spk_counts  = [meta_all[i]["speaker_count"] for i in pdf_indices]

    total_pages = sum(page_counts)
    total_words = sum(word_counts)
    total_turns = sum(turn_counts)

    # Aggregate noise counts
    noise_agg: dict[str, int] = {"teres": 0, "line_num": 0, "timestamp": 0, "email": 0}
    for i in pdf_indices:
        for k, v in meta_all[i]["noise_counts"].items():
            noise_agg[k] += v

    # Collect all body fonts across all PDFs
    all_body_fonts: set[str] = set()
    all_cover_fonts: set[str] = set()
    for i in pdf_indices:
        all_body_fonts.update(meta_all[i]["body_fonts"])
        all_cover_fonts.update(meta_all[i]["cover_fonts"])

    lines: list[str] = []
    w = lines.append

    w("PDF STRUCTURE ANALYSIS REPORT")
    w("Supreme Court Transcripts (TERES Format) -- 24 PDFs")
    w("=" * 78)
    w("")
    w("OVERVIEW")
    w(f"  Total PDFs analyzed        : {len(pdf_indices)}")
    w(f"  Total pages                : {total_pages:,}")
    w(f"  Page count range           : {min(page_counts)} - {max(page_counts)} pages")
    w(f"  Average pages per PDF      : {total_pages / len(pdf_indices):.1f}")
    w(f"  Total words extracted      : {total_words:,}")
    w(f"  Total speaker turns        : {total_turns:,}")
    w(f"  Unique raw speaker strings : {len(extractor.global_word_totals)}")
    w("")

    # Per-file table
    w("=" * 78)
    w("PER-FILE SUMMARY")
    w("=" * 78)
    w("")
    w(f"  {'File':<26}  {'Pages':>5}  {'Speakers':>8}  {'Words':>9}  {'Turns':>6}")
    w(f"  {'-'*26}  {'-'*5}  {'-'*8}  {'-'*9}  {'-'*6}")
    for i in pdf_indices:
        m = meta_all[i]
        w(f"  {m['filename']:<26}  {m['page_count']:>5}  {m['speaker_count']:>8}  "
          f"{m['word_count']:>9,}  {m['turn_count']:>6}")
    w("")

    # Font analysis
    w("=" * 78)
    w("FONT ANALYSIS")
    w("=" * 78)
    w("")
    w("  Cover page (page 1) -- rich font variety (case metadata, parties, dates):")
    for f in sorted(all_cover_fonts)[:20]:
        w(f"    {f}")
    if len(all_cover_fonts) > 20:
        w(f"    ... and {len(all_cover_fonts) - 20} more")
    w("")
    w("  Body pages (pages 2+) -- consistent 3-font pattern "
      "(speaker names / dialogue / line numbers):")
    for f in sorted(all_body_fonts)[:20]:
        w(f"    {f}")
    if len(all_body_fonts) > 20:
        w(f"    ... and {len(all_body_fonts) - 20} more")
    w("")

    # Noise elements
    w("=" * 78)
    w("NOISE ELEMENTS DETECTED AND FILTERED")
    w("=" * 78)
    w("")
    w(f"  TERES watermarks      ('Transcribed by TERES') : {noise_agg['teres']:>7,}")
    w(f"  Line / page numbers   (standalone integers)     : {noise_agg['line_num']:>7,}")
    w(f"  Timestamps            (HH:MM AM/PM IST)         : {noise_agg['timestamp']:>7,}")
    w(f"  Email addresses                                  : {noise_agg['email']:>7,}")
    w(f"  Total noise lines filtered                       : "
      f"{sum(noise_agg.values()):>7,}")
    w("")

    # Extraction approach
    w("=" * 78)
    w("EXTRACTION APPROACH")
    w("=" * 78)
    w("")
    w("  Library     : PyMuPDF (fitz)  --  page.get_text('text') for body text,")
    w("                page.get_text('rawdict') for font metadata.")
    w("  Cover skip  : Page index 0 (cover page) excluded from all extraction.")
    w("  Noise strip : TERES watermarks, standalone integers, HH:MM timestamps, emails.")
    w("")

    text = "\n".join(lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Written: {output_path}")
