"""End-to-end ruled form regression. Requires pdfplumber and python-docx >= 1.2."""

import argparse
from collections import defaultdict
from io import BytesIO
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory
import unicodedata

import pdfplumber
from docx import Document


ROOT = Path(__file__).resolve().parents[1]


def text_key(value):
    # Paragraph/line wraps are not part of the cell contents.
    return "".join(unicodedata.normalize("NFC", value or "").split())


def source_cells(table):
    xs = sorted({x for x0, _, x1, _ in table.cells for x in (x0, x1)})
    ys = sorted({y for _, y0, _, y1 in table.cells for y in (y0, y1)})
    cells = {}
    for row, labels in zip(table.rows, table.extract()):
        for bounds, label in zip(row.cells, labels):
            if bounds is not None:
                x0, y0, x1, y1 = bounds
                cells[ys.index(y0), xs.index(x0), ys.index(y1), xs.index(x1)] = text_key(label)
    return cells


def word_cells(table):
    # python-docx resolves horizontal and vertical merges to the same physical
    # w:tc. Grouping references verifies that a putative blank is independently writable.
    positions = defaultdict(list)
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells, row.grid_cols_before):
            positions[cell._tc].append((r, c))
    cells = {}
    for slots in positions.values():
        rows, cols = zip(*slots)
        r0, c0, r1, c1 = min(rows), min(cols), max(rows) + 1, max(cols) + 1
        if len(slots) != (r1 - r0) * (c1 - c0):
            raise AssertionError(f"Nonrectangular Word merge at row {r0+1}, column {c0+1}")
        row = table.rows[r0]
        cells[r0, c0, r1, c1] = row.cells[c0 - row.grid_cols_before]
    return cells


def check_docx(source, docx):
    with pdfplumber.open(source) as pdf:
        settings = dict(snap_tolerance=0.45, join_tolerance=0.65, intersection_tolerance=0.65)
        expected = [source_cells(t) for p in pdf.pages for t in p.find_tables(settings)]
    document = Document(docx)
    if not expected or len(document.tables) != len(expected):
        raise AssertionError(f"Expected {len(expected)} tables, found {len(document.tables)}")

    filled = []
    for number, (reference, table) in enumerate(zip(expected, document.tables), 1):
        actual = word_cells(table)
        if actual.keys() != reference.keys():
            raise AssertionError(f"Table {number} cell topology changed; missing "
                                 f"{sorted(reference.keys()-actual.keys())}, extra {sorted(actual.keys()-reference.keys())}")
        markers = {}
        for pos, text in reference.items():
            if text != text_key(actual[pos].text):
                raise AssertionError(f"Table {number}, cell {pos}: {text!r} became {actual[pos].text!r}")
            if not text:
                marker = f"{number}-{len(markers)+1}"
                markers[pos] = marker
                actual[pos].paragraphs[0].add_run(marker)
        filled.append(markers)
    copy = BytesIO()
    document.save(copy)
    copy.seek(0)
    for number, (table, markers) in enumerate(zip(Document(copy).tables, filled), 1):
        cells = word_cells(table)
        for pos, marker in markers.items():
            if cells[pos].text.strip() != marker:
                raise AssertionError(f"Fill did not survive reopening table {number}, cell {pos}")
        print(f"Table {number}: {len(cells)} native cells, {len(markers)} independent blanks filled and reopened")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, nargs="?", default=ROOT / "samples/qualification.pdf")
    parser.add_argument("--exe", type=Path, default=ROOT / "bin/PDFtoWord.exe")
    args = parser.parse_args()
    with TemporaryDirectory() as directory:
        target = Path(directory) / "converted.docx"
        subprocess.run([str(args.exe.resolve()), "edit", str(args.source.resolve()), str(target)],
                       check=True, timeout=360)
        check_docx(args.source, target)
