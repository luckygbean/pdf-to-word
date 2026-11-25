"""Exercise mixed-page edit conversion. Pass --word to verify the printed pages in Word.

Requires reportlab, Pillow, pypdfium2, numpy and python-docx. The Word check runs
on Windows with Microsoft Word available to the current interactive user.
"""
import argparse
from io import BytesIO
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory

from docx import Document
import numpy as np
from PIL import Image, ImageDraw
import pypdfium2 as pdfium
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


def create_fixture(path):
    c = canvas.Canvas(str(path), pagesize=(420, 580), invariant=1)
    c.drawString(40, 525, "Inspection release checklist")
    for name, x, checked in (("released", 45, True), ("pending", 130, False)):
        c.acroForm.checkbox(name=name, x=x, y=450, size=28, checked=checked,
                            textColor=colors.black, fillColor=colors.white,
                            borderColor=colors.black, borderWidth=1)
    c.showPage()
    c.drawString(40, 525, "Procedure overview")
    c.drawString(40, 488, "Inspect the sample and enter the measured quantities.")
    c.drawString(40, 18, "End of procedure page")
    c.showPage()

    xs = [34, 154, 260, 388]
    ys = [465, 431, 397, 363, 329]
    for y in (ys[0], *ys[2:]):
        c.line(xs[0], y, xs[-1], y)
    # Office PDFs can paint a double header rule in separate pieces per column,
    # including a tiny filled rectangle at each vertical junction.
    for left, right in zip(xs, xs[1:]):
        for offset in (-1.08, .36):
            c.rect(left+.72, ys[1]+offset, right-left-1.44, .72, stroke=0, fill=1)
    for x in xs:
        c.rect(x-.72, ys[1]-1.14, 1.44, 2.28, stroke=0, fill=1)
    for i, x in enumerate(xs):
        c.line(x, ys[0] if i in (0, 3) else ys[1], x, ys[-1])
    c.drawCentredString(211, 445, "Inventory fields")
    for i, label in enumerate(("Item", "Planned", "Actual")):
        c.drawString(xs[i] + 5, 410, label)
    c.drawString(39, 376, "Valve")
    c.drawString(39, 342, "Pump")

    # A vertically merged cell starts far below its neighbors. A row-wide
    # top margin can move text out of every short scoring row beside it.
    columns = [34, 143, 272, 388]
    row_edges = [252, 211, 170, 129, 88]
    for x in columns:
        c.line(x, row_edges[0], x, row_edges[-1])
    for y in row_edges:
        c.line(columns[0] if y in (252, 88) else columns[1], y, columns[-1], y)
    c.drawString(40, 147, "Review group")
    for index, y in enumerate(row_edges[:-1], 1):
        c.drawString(149, y - 14, f"Checkpoint {index}")
        c.drawString(149, y - 30, f"Threshold {index}")
        c.drawString(278, y - 23, f"Grade {index}")
    c.showPage()

    c.setPageSize((580, 420))
    scan = Image.new("RGB", (500, 290), "white")
    pen = ImageDraw.Draw(scan)
    pen.rectangle((18, 26, 222, 232), fill="#207754")
    pen.ellipse((296, 73, 471, 247), fill="#bb582b")
    pen.line((10, 265, 480, 265), fill="#101010", width=6)
    c.drawImage(ImageReader(scan), 36, 59, width=500, height=290)
    c.showPage()
    c.save()


def check_editable(path):
    doc = Document(path)
    matches = [t for t in doc.tables if t.cell(0, 0).text.strip() == "Inventory fields"]
    assert len(matches) == 1, "Ruled page did not produce exactly one matching native table"
    table = matches[0]
    assert len(table.rows) == 4 and len(table.columns) == 3
    assert len({c._tc for row in table.rows for c in row.cells}) == 10
    for row in (2, 3):
        for col in (1, 2):
            cell = table.cell(row, col)
            assert not cell.text.strip()
            cell.text = f"entry-{row}-{col}"
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    reopened = next(t for t in Document(buf).tables if t.cell(0, 0).text.strip() == "Inventory fields")
    for row in (2, 3):
        for col in (1, 2):
            assert reopened.cell(row, col).text == f"entry-{row}-{col}"
    scores = next(t for t in doc.tables if t.cell(0, 0).text.strip() == "Review group")
    assert len(scores.rows) == 4 and len({row.cells[0]._tc for row in scores.rows}) == 1
    for row in range(4):
        assert f"Checkpoint {row+1}" in scores.cell(row, 1).text
        assert f"Threshold {row+1}" in scores.cell(row, 1).text
        assert f"Grade {row+1}" in scores.cell(row, 2).text
    # Text in frame paragraphs/text boxes is still editable in Word.
    text = "".join(doc.element.xpath("//w:t/text()"))
    assert "Procedure overview" in text and "End of procedure page" in text


def export_word(docx, printed, folder):
    script = folder / "render.ps1"
    script.write_text("""param([string]$Source, [string]$Output)
$ErrorActionPreference='Stop'
$word = New-Object -ComObject Word.Application
$doc = $null
try {
    $word.Visible=$false
    $word.DisplayAlerts=0
    $word.AutomationSecurity=3
    $doc=$word.Documents.Open($Source,$false,$true,$false)
    $doc.Repaginate()
    $doc.ExportAsFixedFormat($Output,17)
} finally {
    try { if ($doc) { $doc.Close([ref]0) } } finally { $word.Quit([ref]0) }
}
""", encoding="utf-8-sig")
    subprocess.run(["powershell.exe", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass",
                    "-File", str(script), str(docx), str(printed)], check=True, timeout=240)


def render(doc, page_no):
    page = doc[page_no]
    bmp = page.render(scale=2)
    image = bmp.to_pil().convert("RGB")
    bmp.close()
    page.close()
    return image


def check_print(original, printed):
    with pdfium.PdfDocument(original) as src, pdfium.PdfDocument(printed) as out:
        assert len(out) == len(src) == 4, f"Expected four distinct pages; Word exported {len(out)}"
        src.init_forms()
        out.init_forms()
        for i in range(4):
            assert max(abs(a-b) for a,b in zip(src.get_page_size(i), out.get_page_size(i))) < .6
        for i, markers in [(1, ["Procedure overview", "End of procedure page"]),
                           (2, ["Inventory fields", "Valve", "Pump"])]:
            page = out[i]
            text = page.get_textpage()
            normalized = "".join(text.get_text_bounded().split())
            for marker in markers:
                assert "".join(marker.split()) in normalized, f"Missing {marker!r} on page {i+1}"
            text.close()
            page.close()
        # Inspect the centers of the two checkbox appearances. The first must retain
        # substantially more dark ink than its unchecked twin after conversion.
        def checkbox_difference(image):
            dark = 255-np.asarray(image.convert("L"), dtype=np.float64)
            return dark[210:252, 97:139].mean() - dark[210:252, 267:309].mean()
        check_ref = checkbox_difference(render(src, 0))
        assert check_ref > 10, "Fixture does not render a distinct check mark"
        assert checkbox_difference(render(out, 0)) > .45*check_ref, "Checkbox state lost"
        # Measure actual ink at each baseline, inset from the visible cell rules.
        # Word can keep a clipped label in the DOCX or PDF text stream.
        ref, got = [np.asarray(render(doc, 2).convert("L")) < 180 for doc in (src, out)]
        for row_top in (252, 211, 170, 129):
            for left, right, baseline in ((147, 264, row_top-14),
                                          (147, 264, row_top-30), (276, 380, row_top-23)):
                area = (slice((580-baseline-11)*2, (580-baseline+3)*2), slice(left*2, right*2))
                expected = ref[area].sum()
                assert expected > 60, "Fixture scoring cell lacks visible text"
                assert got[area].sum() > .45*expected, f"Clipped score text at ({left}, {baseline})"
        scan = render(src, 3)
        actual = render(out, 3).resize(scan.size)
        a, b = np.asarray(scan, dtype=float), np.asarray(actual, dtype=float)
        assert np.abs(a-b).mean() < 3, "Scan artwork changed page, position or orientation"


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", type=Path, default=Path(__file__).resolve().parents[1]/"bin/PDFtoWord.exe")
    parser.add_argument("--word", action="store_true")
    args = parser.parse_args()
    with TemporaryDirectory() as root:
        folder = Path(root)
        source, target, printed = [folder/n for n in ("mixed.pdf","converted.docx","roundtrip.pdf")]
        create_fixture(source)
        subprocess.run([str(args.exe.resolve()),"edit",str(source),str(target)], check=True, timeout=360)
        check_editable(target)
        if args.word:
            export_word(target, printed, folder)
            check_print(source, printed)
        print("Mixed page regression passed.")
