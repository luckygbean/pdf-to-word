"""Compile the same managed engine used by the desktop app into a Windows wheel."""
import os
from pathlib import Path
import shutil
import subprocess
import sysconfig
from tempfile import TemporaryDirectory

from setuptools import Distribution, setup
from setuptools.command.bdist_wheel import bdist_wheel
from setuptools.command.build_py import build_py


ROOT = Path(__file__).resolve().parent


class BuildEngine(build_py):
    def run(self):
        if sysconfig.get_platform() != "win-amd64":
            raise RuntimeError("Building pdfword-engine requires Windows x64 Python.")
        super().run()
        package = Path(self.build_lib) / "pdfword"
        powershell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
        environment = os.environ.copy()
        # pip may inherit a PowerShell 7 module path. Windows PowerShell
        # must resolve its own Get-FileHash/Invoke-WebRequest modules.
        environment.pop("PSModulePath", None)
        environment.pop("PSMODULEPATH", None)
        with TemporaryDirectory() as output:
            command = [str(powershell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                       "-File", str(ROOT / "build.ps1"), "-EngineOnly", "-OutputDirectory", output]
            if os.environ.get("PDFWORD_WINDOWS_METADATA"):
                command.extend(["-WindowsMetadata", os.environ["PDFWORD_WINDOWS_METADATA"]])
            subprocess.run(command, cwd=ROOT, env=environment, check=True)
            shutil.copy2(Path(output) / "PDFtoWord.Engine.dll", package)
            shutil.copy2(ROOT / "obj/pdfium/PDFium-NOTICES.txt", package)


class WindowsWheel(bdist_wheel):
    def get_tag(self):
        # The DLL is x64 native/managed code and does not link to a Python ABI.
        return "py3", "none", "win_amd64"


class EngineDistribution(Distribution):
    def has_ext_modules(self):
        # Place the engine in platlib, beside the Python wrapper.
        return True


setup(cmdclass={"build_py": BuildEngine, "bdist_wheel": WindowsWheel}, distclass=EngineDistribution)
