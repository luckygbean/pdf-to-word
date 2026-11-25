"""Convert local PDFs to editable or fixed-layout DOCX documents on Windows."""
from os import PathLike, fspath, path
from typing import Optional, Union


__all__ = ["convert", "ConversionError"]


class ConversionError(RuntimeError):
    """The PDF engine could not complete the requested conversion."""


def convert(
    input_pdf: Union[str, PathLike],
    output_docx: Union[str, PathLike],
    mode: str = "edit",
    dpi: int = 300,
    password: Optional[str] = None,
) -> str:
    """Write a new DOCX file and return the engine's conversion report.

    Modes: edit (editable text/tables), form (fillable cells), fixed (page
    images) and ocr (recognized text). Existing output files are protected.
    The edit importer may use installed Microsoft Word for general PDFs.
    """
    if mode not in ("edit", "form", "fixed", "ocr"):
        raise ValueError("mode must be 'edit', 'form', 'fixed' or 'ocr'")
    if not isinstance(dpi, int) or isinstance(dpi, bool):
        raise TypeError("dpi must be an integer")
    if password is not None and not isinstance(password, str):
        raise TypeError("password must be a string or None")
    source, target = fspath(input_pdf), fspath(output_docx)
    if not isinstance(source, str) or not isinstance(target, str):
        raise TypeError("input_pdf and output_docx must be text paths")
    from ._engine import ManagedException, PdfConverter

    try:
        return PdfConverter.Convert(path.abspath(source), path.abspath(target), mode, dpi, password)
    except ManagedException as exc:
        raise ConversionError(exc.Message) from exc
