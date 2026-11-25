"""Run with the Windows wheel installed (no checkout on the Python module path)."""
from importlib.metadata import distribution
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree
from zipfile import ZipFile

import pdfword


def main():
    installed = {str(f).replace("\\", "/") for f in distribution("pdfword-engine").files}
    assert "pdfword/PDFtoWord.Engine.dll" in installed
    assert not any(f.lower().endswith(".exe") for f in installed)

    source = Path(__file__).resolve().parents[1] / "samples/qualification.pdf"
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with TemporaryDirectory() as directory:
        for mode in ("edit", "form", "fixed", "ocr"):
            target = Path(directory) / f"result-{mode}.docx"
            report = pdfword.convert(source, target, mode=mode)
            assert isinstance(report, str)
            with ZipFile(target) as parts:
                body = ElementTree.fromstring(parts.read("word/document.xml"))
                if mode in ("edit", "form"):
                    assert len(list(body.iter(w + "tbl"))) == 3
                if mode in ("form", "fixed"):
                    assert sum(p.startswith("word/media/") for p in parts.namelist()) == 2
            print(mode, target.stat().st_size, "bytes")

        before = target.read_bytes()
        try:
            pdfword.convert(source, target)
        except pdfword.ConversionError:
            pass
        else:
            raise AssertionError("A conversion overwrote the existing output file")
        assert target.read_bytes() == before

    import clr
    from PdfToWord.Engine import PdfConverter
    assembly = Path(clr.GetClrType(PdfConverter).Assembly.Location).resolve()
    assert assembly == Path(pdfword.__file__).resolve().with_name("PDFtoWord.Engine.dll")
    print("Loaded the installed engine into this Python process")


if __name__ == "__main__":
    main()
