"""Exercise edit mode on a generated scanned cover with two overlapping stamps.

Run on a Windows x64 machine with Chinese OCR available, Pillow, reportlab,
pypdf and python-docx. Inputs and outputs stay in the temporary test directory.
"""
import argparse
from io import BytesIO
import os
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory

from docx import Document
from PIL import Image, ImageChops, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject, DecodedStreamObject, DictionaryObject, NameObject,
    NumberObject, RectangleObject, TextStringObject,
)
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


def scanned_paper():
    width, height, scale = 420, 580, 3
    paper = Image.new("RGB", (width * scale, height * scale), "white")
    print_layer = ImageDraw.Draw(paper)
    font_path = Path(os.environ["WINDIR"]) / "Fonts/msyh.ttc"
    for left, top, size, label in [
        (105, 90, 26, "设备检验记录"),
        (80, 353, 14, "记录编号 24861"),
        (80, 403, 14, "记录部门 质量检查组"),
        (80, 452, 14, "日期 2025年5月"),
    ]:
        font = ImageFont.truetype(str(font_path), size * scale)
        print_layer.text((left * scale, top * scale), label, font=font, fill="black")
    # Multiply the red scan ink over already-printed text. Both layers occupy
    # the same pixels in the embedded raster, just as on a stamped paper form.
    red = Image.new("RGB", paper.size, "white")
    pen = ImageDraw.Draw(red)
    for inset in (0, 10):
        pen.ellipse((91*scale+inset, 321*scale+inset, 175*scale-inset, 405*scale-inset),
                    outline=(230, 80, 80), width=2*scale)
    pen.text((110*scale, 330*scale), "TEST", fill=(230, 80, 80),
             font=ImageFont.truetype(str(font_path), 14*scale))
    return ImageChops.multiply(paper, red)


def create_cover(path, scan=True, opaque=False):
    width, height = 420, 580
    base = BytesIO()
    c = canvas.Canvas(base, pagesize=(width, height), invariant=1)
    if scan:
        c.drawImage(ImageReader(scanned_paper()), 0, 0, width, height)
    else:
        c.drawString(80, 455, "Department inspection record")
        if opaque:
            c.drawString(160, 165, "MASKED")
        else:
            c.drawString(120, 165, "Batch reference QA-24861")
    c.showPage()
    c.save()
    base.seek(0)

    # A transparent, unsigned signature widget with a visible test appearance.
    # Use actual PDF curve operators, independent of the base image.
    pdf = PdfWriter()
    page = pdf.add_page(PdfReader(base).pages[0])
    appearance = DecodedStreamObject()
    if opaque:
        appearance.set_data(b"q 0 0 0 rg 0 0 80 80 re f Q")
    else:
        appearance.set_data(b"q 1 0 0 RG 2 w 79 40 m 79 61 61 79 40 79 c "
                            b"19 79 1 61 1 40 c 1 19 19 1 40 1 c 61 1 79 19 79 40 c S "
                            b"27 40 m 36 31 l 53 52 l S Q")
    appearance.update({NameObject("/Type"): NameObject("/XObject"),
                       NameObject("/Subtype"): NameObject("/Form"),
                       NameObject("/BBox"): RectangleObject([0, 0, 80, 80]),
                       NameObject("/Resources"): DictionaryObject()})
    widget = DictionaryObject({
        NameObject("/Type"): NameObject("/Annot"), NameObject("/Subtype"): NameObject("/Widget"),
        NameObject("/FT"): NameObject("/Sig"), NameObject("/F"): NumberObject(4),
        NameObject("/T"): TextStringObject("Test approval"), NameObject("/P"): page.indirect_reference,
        NameObject("/Rect"): RectangleObject([155, 130, 235, 210]),
        NameObject("/AP"): DictionaryObject({NameObject("/N"): pdf._add_object(appearance)}),
    })
    fields = ArrayObject([pdf._add_object(widget)])
    page[NameObject("/Annots")] = fields
    pdf.root_object[NameObject("/AcroForm")] = DictionaryObject({NameObject("/Fields"): fields})
    pdf.write(path)


def check_layers(path, scan=True):
    doc = Document(path)
    text = "".join("".join(doc.element.xpath("//w:t/text()")).split())
    terms = ("设备检验记录", "24861", "质量检查组", "2025") if scan else ("Departmentinspectionrecord", "QA-24861")
    for content in terms:
        assert content in text, f"Expected editable text {content!r}, got {text!r}"

    anchors = doc.element.xpath("//wp:anchor")
    front = [a for a in anchors if a.get("behindDoc") in ("0", "false")]
    expected = [(133, 363), (195, 410)] if scan else [(195, 410)]
    assert len(front) == len(expected), f"Expected {len(expected)} selectable stamps, got {len(front)}"
    centers = []
    for anchor in front:
        axis = [anchor.find(f"{{*}}position{letter}") for letter in ("H", "V")]
        assert all(pos.get("relativeFrom") == "page" for pos in axis)
        x, y = (int(pos.find("{*}posOffset").text) / 12700 for pos in axis)
        extent = anchor.find("{*}extent")
        w, h = (int(extent.get(k)) / 12700 for k in ("cx", "cy"))
        assert 72 < w < 100 and 72 < h < 100, (w, h)
        centers.append((x + w/2, y + h/2))
        image_id = anchor.xpath(".//a:blip/@r:embed")[0]
        with Image.open(BytesIO(doc.part.related_parts[image_id].blob)) as image:
            assert "A" in image.getbands() and image.getchannel("A").getextrema()[0] == 0
    for (x, y), (want_x, want_y) in zip(sorted(centers), expected):
        assert abs(x - want_x) < 4 and abs(y - want_y) < 4, centers


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", type=Path, default=Path(__file__).resolve().parents[1]/"bin/PDFtoWord.exe")
    args = parser.parse_args()
    with TemporaryDirectory() as out:
        for stem, scan in (("scan", True), ("digital", False), ("opaque", False)):
            source, target = [Path(out) / (stem + ext) for ext in (".pdf", ".docx")]
            create_cover(source, scan, opaque=stem == "opaque")
            subprocess.run([str(args.exe.resolve()), "edit", str(source), str(target)], check=True, timeout=180)
            if stem == "opaque":
                doc = Document(target)
                assert "MASKED" not in "".join(doc.element.xpath("//w:t/text()")), "Revealed masked content"
                assert not doc.element.xpath("//wp:docPr[starts-with(@name,'印章')]")
            else:
                check_layers(target, scan)
        print("Scanned and digital covers retain editable text and transparent stamps at the original positions.")
