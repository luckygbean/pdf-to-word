"""Load the managed DLL once, on the first conversion."""
from pathlib import Path
import sysconfig

from pythonnet import get_runtime_info, load

if sysconfig.get_platform() != "win-amd64":
    raise RuntimeError("pdfword-engine requires Windows x64 Python.")
runtime = get_runtime_info()
if runtime is not None and runtime.kind != ".NET Framework":
    raise RuntimeError("pdfword-engine requires Python.NET's netfx runtime. This process uses " + runtime.kind)
if runtime is None:
    load("netfx")

import clr

clr.AddReference(str(Path(__file__).with_name("PDFtoWord.Engine.dll").resolve()))
from PdfToWord.Engine import PdfConverter
from System import Exception as ManagedException
