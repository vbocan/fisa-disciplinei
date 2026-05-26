#!/usr/bin/env python3
"""Verify consistency of section 3 in a UPT syllabus (.docx).

Usage:
  python verify.py <fisa.docx>          # detailed report for one file
  python verify.py <folder>              # batch summary for all .docx in a folder
  python verify.py <folder> --verbose    # batch + detailed report per file

Exits 0 if all files are consistent, 1 if any mismatches are found.
"""
import math
import re
import sys
import zipfile
from pathlib import Path


def extract_text(docx_path):
    """Extract visible text, skipping footnote reference markers.

    Footnote references in OOXML are <w:footnoteReference w:id="N"/> elements,
    but the syllabus template inserts the footnote number as a styled run
    (rStyle="FootnoteReference") which appears inline as a digit. We drop
    runs that carry that style so "3.8 Total ore/săptămână ⁹3.6" becomes
    "3.8 Total ore/săptămână 3.6".
    """
    with zipfile.ZipFile(docx_path) as z:
        with z.open("word/document.xml") as f:
            xml = f.read().decode("utf-8")
    # Strip <w:footnoteReference .../> elements outright
    xml = re.sub(r"<w:footnoteReference[^/]*/>", "", xml)
    # Drop entire runs that are rendered as superscript — the UPT template
    # uses these for footnote-style numerals inline in field labels.
    xml = re.sub(
        r"<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:vertAlign\s+w:val=\"superscript\"\s*/>.*?</w:r>",
        "",
        xml,
        flags=re.DOTALL,
    )
    texts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)
    return "".join(texts)


# Match the next field marker. Section labels are like "3.1 Număr...", "3.2 ore curs",
# "3.1* Număr total...", "4.Precondiţii". Don't match bare "3.6" — that's a value.
_NEXT_LABEL = re.compile(
    r"(?:"
    r"3\.\d\*?\s*(?:Număr|ore\s|Total)"  # 3.X / 3.X* followed by Număr, ore, or Total
    r"|[4-9]\.\s*\w"                    # next major section (4. Precondiții, etc.)
    r"|,\s*format din"
    r"|ore curs"
    r"|ore seminar"
    r"|ore practică"
    r"|ore elaborare"
    r"|ore documentare"
    r"|ore studiu"
    r"|ore pregătire"
    r")"
)


def find_value_after(text, label):
    """Find the numeric value that immediately follows the given label.

    The doc text is concatenated without separators, so we cut at the next
    known label/section marker to avoid greedy regex eating "0" + "3.3" as "03.3".
    """
    idx = text.find(label)
    if idx < 0:
        return None
    tail = text[idx + len(label):]
    # Find where the next field/marker begins, slice up to there
    cut = _NEXT_LABEL.search(tail)
    window = tail[:cut.start()] if cut else tail[:80]
    m = re.search(r"([0-9]+(?:[.,][0-9]+)?)", window)
    if not m:
        return None
    return float(m.group(1).replace(",", "."))


def _cell_text(tc_xml):
    """Concatenate all <w:t> text inside a single <w:tc>...</w:tc> cell."""
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", tc_xml))


def sum_section_hours(docx_path, section_marker):
    """Sum the hour values in section 8.1 (Curs) or 8.2 (Activităţi aplicative).

    The UPT template packs BOTH 8.1 and 8.2 into a single <w:tbl>, separated by
    a 'Bibliografie' single-cell row. So we:
      1. Find any row whose text contains the marker AND 'Număr de ore'
         (column headers).
      2. Sum the SECOND cell of subsequent rows until hitting another header
         row or a 'Bibliografie' row (which marks the end of the sub-section).

    Direct XML parsing avoids two text-mode pitfalls: cells are concatenated
    without separators in the joined text, and digits in activity titles like
    'Laborator #1' look like hour values.

    Returns 0.0 if the sub-section exists but is empty, None if not found.
    """
    with zipfile.ZipFile(docx_path) as z:
        with z.open("word/document.xml") as f:
            xml = f.read().decode("utf-8")

    for tbl_m in re.finditer(r"<w:tbl>.*?</w:tbl>", xml, flags=re.DOTALL):
        rows = re.findall(r"<w:tr\b.*?</w:tr>", tbl_m.group(0), flags=re.DOTALL)
        row_texts = [
            "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", r)) for r in rows
        ]

        header_idx = None
        for i, t in enumerate(row_texts):
            if section_marker in t and "Număr de ore" in t:
                header_idx = i
                break
        if header_idx is None:
            continue

        total = 0.0
        found_any = False
        for i in range(header_idx + 1, len(rows)):
            row_text = row_texts[i]
            # Stop at the end of this sub-section
            if "Bibliografie" in row_text:
                break
            # Another section header inside same table (e.g., 8.2 after 8.1)
            if "Număr de ore" in row_text:
                break
            cells = re.findall(r"<w:tc\b.*?</w:tc>", rows[i], flags=re.DOTALL)
            if len(cells) < 2:
                continue
            hours_text = _cell_text(cells[1]).strip()
            if not hours_text:
                continue
            m = re.search(r"([0-9]+(?:[.,][0-9]+)?)", hours_text)
            if m:
                total += float(m.group(1).replace(",", "."))
                found_any = True
        return total if found_any else 0.0

    return None


def roundup(x, d=1):
    m = 10 ** d
    return math.ceil(x * m) / m


def num(v):
    if v is None:
        return "—"
    return f"{v:g}"


def verify_file(path, show_details=True):
    """Run all checks on one syllabus. Returns (mismatch_count, total_checks)."""
    text = extract_text(path)

    fields = {
        "3.1":  find_value_after(text, "3.1 Număr de ore asistate integral/săptămână"),
        "3.2":  find_value_after(text, "3.2 ore curs"),
        "3.3":  find_value_after(text, "3.3 ore seminar"),
        "3.1*": find_value_after(text, "3.1* Număr total de ore asistate integral/sem"),
        "3.2*": find_value_after(text, "3.2* ore curs"),
        "3.3*": find_value_after(text, "3.3* ore seminar"),
        "3.4":  find_value_after(text, "3.4 Număr de ore asistate parțial/săptămână")
                or find_value_after(text, "3.4 Număr de ore asistate parţial/săptămână"),
        "3.5":  find_value_after(text, "3.5 ore practică"),
        "3.6":  find_value_after(text, "3.6 ore elaborare proiect"),
        "3.4*": find_value_after(text, "3.4* Număr total de ore asistate parțial/")
                or find_value_after(text, "3.4* Număr total de ore asistate parţial/"),
        "3.5*": find_value_after(text, "3.5* ore practică"),
        "3.6*": find_value_after(text, "3.6* ore elaborare proiect"),
        "3.7":  find_value_after(text, "3.7 Număr de ore activități neasistate")
                or find_value_after(text, "3.7 Număr de ore activităţi neasistate"),
        "3.7*": find_value_after(text, "3.7* Număr total de ore activități neasistate")
                or find_value_after(text, "3.7* Număr total de ore activităţi neasistate"),
        "3.8":  find_value_after(text, "3.8 Total ore/săptămână"),
        "3.8*": find_value_after(text, "3.8* Total ore/semestru"),
        "3.9":  find_value_after(text, "3.9 Număr de credite"),
        "8.1_sum": sum_section_hours(path, "Curs"),
        "8.2_sum": sum_section_hours(path, "Activităţi aplicative")
                    or sum_section_hours(path, "Activități aplicative"),
    }

    # Treat missing partial-hours rows as zero (often blank for non-practica courses)
    for k in ("3.4", "3.5", "3.6", "3.4*", "3.5*", "3.6*"):
        if fields[k] is None:
            fields[k] = 0.0

    if show_details:
        print("Câmpuri detectate:")
        for k in ("3.2", "3.3", "3.5", "3.6", "3.9"):
            print(f"  {k:6s} = {num(fields[k])}")
        print()

    checks = [
        ("3.1 = 3.2 + 3.3",          fields["3.1"],  (fields["3.2"] or 0) + (fields["3.3"] or 0)),
        ("3.4 = 3.5 + 3.6",          fields["3.4"],  fields["3.5"] + fields["3.6"]),
        ("3.2* = 3.2 × 14",          fields["3.2*"], (fields["3.2"] or 0) * 14),
        ("3.3* = 3.3 × 14",          fields["3.3*"], (fields["3.3"] or 0) * 14),
        ("3.5* = 3.5 × 14",          fields["3.5*"], fields["3.5"] * 14),
        ("3.6* = 3.6 × 14",          fields["3.6*"], fields["3.6"] * 14),
        ("3.1* = 3.2* + 3.3*",       fields["3.1*"], (fields["3.2*"] or 0) + (fields["3.3*"] or 0)),
        ("3.4* = 3.5* + 3.6*",       fields["3.4*"], fields["3.5*"] + fields["3.6*"]),
        ("3.8* = 3.9 × 25",          fields["3.8*"], (fields["3.9"] or 0) * 25),
        ("3.8 = ROUNDUP(3.8*/14,1)", fields["3.8"],  roundup((fields["3.8*"] or 0) / 14, 1)),
        ("3.7* = 3.8* − 3.1* − 3.4*", fields["3.7*"],
            (fields["3.8*"] or 0) - (fields["3.1*"] or 0) - fields["3.4*"]),
        ("3.7 = 3.8 − 3.1 − 3.4",    fields["3.7"],
            (fields["3.8"] or 0) - (fields["3.1"] or 0) - fields["3.4"]),
        ("Σ 8.1 (curs) = 3.2*",     fields["8.1_sum"], fields["3.2*"] or 0),
        ("Σ 8.2 (aplicative) = 3.3*", fields["8.2_sum"], fields["3.3*"] or 0),
    ]

    if show_details:
        print(f"{'Verificare':<34} {'Doc':>6} {'Calc':>6}   Status")
        print("-" * 62)
    mismatches = 0
    for name, actual, expected in checks:
        if actual is None:
            status = "—"
        elif abs(actual - expected) < 0.05:
            status = "OK"
        else:
            status = "MISMATCH"
            mismatches += 1
        if show_details:
            print(f"{name:<34} {num(actual):>6} {num(expected):>6}   {status}")

    if show_details:
        print()
        if mismatches == 0:
            print("Fișa este consistentă.")
        else:
            print(f"{mismatches} discrepanțe găsite.")

    return mismatches, len(checks)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    verbose = "--verbose" in sys.argv or "-v" in sys.argv

    if not args:
        print(__doc__, file=sys.stderr)
        sys.exit(2)

    path = Path(args[0])
    if not path.exists():
        print(f"Path not found: {path}", file=sys.stderr)
        sys.exit(2)

    # Single-file mode
    if path.is_file():
        mismatches, _ = verify_file(path, show_details=True)
        sys.exit(0 if mismatches == 0 else 1)

    # Batch mode: scan directory for .docx files
    docx_files = sorted(p for p in path.rglob("*.docx") if not p.name.startswith("~$"))
    if not docx_files:
        print(f"No .docx files found in {path}", file=sys.stderr)
        sys.exit(2)

    print(f"Verificare batch: {len(docx_files)} fișe în {path}\n")
    results = []
    for f in docx_files:
        if verbose:
            print(f"\n{'='*70}\n{f.name}\n{'='*70}")
        try:
            mismatches, total = verify_file(f, show_details=verbose)
            results.append((f.name, mismatches, total, None))
        except Exception as e:
            results.append((f.name, None, None, str(e)))

    # Summary table
    print(f"\n{'='*70}")
    print(f"Sumar — {len(docx_files)} fișe verificate")
    print(f"{'='*70}")
    print(f"{'Fișier':<55} {'Verif.':>8}  Status")
    print("-" * 75)
    consistent = 0
    for name, mismatches, total, err in results:
        if err:
            status = f"EROARE: {err[:30]}"
            verif = "—"
        elif mismatches == 0:
            status = "CONSISTENT"
            verif = f"{total}/{total}"
            consistent += 1
        else:
            status = f"{mismatches} discrepanțe"
            verif = f"{total - mismatches}/{total}"
        # Truncate long names
        display_name = name if len(name) <= 53 else name[:50] + "..."
        print(f"{display_name:<55} {verif:>8}  {status}")

    print()
    print(f"Consistente: {consistent}/{len(results)}")
    sys.exit(0 if consistent == len(results) else 1)


if __name__ == "__main__":
    main()
