"""Capture engine: let an installed spreadsheet application print the workbook to PDF.

This gives pixel-faithful output (including every chart type, SmartArt, form
controls...) when Microsoft Excel or LibreOffice is available on the machine.
Neither is bundled or required; the native engine needs neither.

* Excel on Windows - driven through COM from PowerShell (no extra Python packages).
* Excel on macOS   - driven through AppleScript (osascript).
* LibreOffice      - ``soffice --headless --convert-to pdf`` on any OS.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path
from typing import List, Optional


class CaptureError(RuntimeError):
    pass


_LO_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/opt/libreoffice/program/soffice",
    "/snap/bin/libreoffice",
]
_MAC_EXCEL = "/Applications/Microsoft Excel.app"


def find_libreoffice() -> Optional[str]:
    env = os.environ.get("EXCEL2PDF_SOFFICE")
    if env and Path(env).exists():
        return env
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for cand in _LO_CANDIDATES:
        if Path(cand).exists():
            return cand
    if sys.platform == "win32":
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if base:
                for d in Path(base).glob("LibreOffice*/program/soffice.exe"):
                    return str(d)
    return None


def has_excel() -> bool:
    if sys.platform == "win32":
        try:
            import winreg

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
                return True
        except OSError:
            return False
    if sys.platform == "darwin":
        return Path(_MAC_EXCEL).exists()
    return False


def available_engines() -> List[str]:
    out = []
    if has_excel():
        out.append("excel")
    if find_libreoffice():
        out.append("libreoffice")
    return out


def convert(src: Path, dst: Path, engine: str = "capture", timeout: float = 180.0) -> None:
    """Convert ``src`` to ``dst`` with Excel or LibreOffice. engine: capture | excel | libreoffice."""
    if engine == "capture":
        engines = available_engines()
        if not engines:
            raise CaptureError(
                "The capture engine needs Microsoft Excel (Windows/macOS) or LibreOffice installed. "
                "Install one of them, or use engine='native'."
            )
        engine = engines[0]
    if engine == "excel":
        if sys.platform == "win32":
            _excel_windows(src, dst, timeout)
        elif sys.platform == "darwin":
            _excel_mac(src, dst, timeout)
        else:
            raise CaptureError("Microsoft Excel capture is only available on Windows and macOS.")
    elif engine == "libreoffice":
        _libreoffice(src, dst, timeout)
    else:
        raise ValueError(f"unknown capture engine {engine!r}")
    if not dst.exists() or dst.stat().st_size == 0:
        raise CaptureError(f"{engine} did not produce a PDF")


def _run(cmd, timeout: float, env=None) -> subprocess.CompletedProcess:
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=env, check=False)
    except subprocess.TimeoutExpired as exc:
        raise CaptureError(f"timed out after {timeout:.0f}s: {cmd[0]}") from exc
    except OSError as exc:
        raise CaptureError(f"cannot start {cmd[0]}: {exc}") from exc


def _libreoffice(src: Path, dst: Path, timeout: float) -> None:
    soffice = find_libreoffice()
    if not soffice:
        raise CaptureError("LibreOffice (soffice) was not found. Set EXCEL2PDF_SOFFICE to its path.")
    with tempfile.TemporaryDirectory(prefix="excel2pdf-lo-") as tmp:
        profile = Path(tmp) / "profile"
        outdir = Path(tmp) / "out"
        outdir.mkdir()
        cmd = [
            soffice, "--headless", "--norestore", "--nolockcheck", "--nodefault", "--nologo",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to", "pdf", "--outdir", str(outdir), str(src),
        ]
        res = _run(cmd, timeout)
        produced = outdir / (src.stem + ".pdf")
        if not produced.exists():
            raise CaptureError(f"LibreOffice failed: {res.stderr.strip() or res.stdout.strip()}")
        shutil.move(str(produced), str(dst))


_PS_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$src = $env:EXCEL2PDF_SRC
$dst = $env:EXCEL2PDF_DST
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
try {
    $wb = $excel.Workbooks.Open($src, 0, $true)
    try {
        $wb.ExportAsFixedFormat(0, $dst, 0, $true, $false)
    } finally {
        $wb.Close($false)
    }
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
"""


def _excel_windows(src: Path, dst: Path, timeout: float) -> None:
    env = dict(os.environ)
    env["EXCEL2PDF_SRC"] = str(src.resolve())
    env["EXCEL2PDF_DST"] = str(dst.resolve())
    shell = shutil.which("powershell") or shutil.which("pwsh")
    if not shell:
        raise CaptureError("PowerShell is required to drive Excel.")
    res = _run([shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", _PS_SCRIPT],
               timeout, env)
    if res.returncode != 0:
        raise CaptureError(f"Excel failed: {res.stderr.strip() or res.stdout.strip()}")


_APPLESCRIPT = r"""
on run argv
    set srcPath to item 1 of argv
    set dstPath to item 2 of argv
    tell application "Microsoft Excel"
        set display alerts to false
        set wb to open workbook workbook file name srcPath read only true
        save workbook as wb filename dstPath file format PDF file format
        close wb saving no
    end tell
end run
"""


def _excel_mac(src: Path, dst: Path, timeout: float) -> None:
    # Excel for Mac is sandboxed: work inside its container to avoid file-access prompts.
    container = Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "tmp"
    work = container / f"excel2pdf-{uuid.uuid4().hex[:8]}"
    work.mkdir(parents=True, exist_ok=True)
    try:
        local_src = work / ("input" + src.suffix)
        local_dst = work / "output.pdf"
        shutil.copyfile(src, local_src)
        res = _run(["osascript", "-e", _APPLESCRIPT, str(local_src), str(local_dst)], timeout)
        if res.returncode != 0 or not local_dst.exists():
            raise CaptureError(f"Excel failed: {res.stderr.strip() or res.stdout.strip() or 'no output'}")
        shutil.move(str(local_dst), str(dst))
    finally:
        shutil.rmtree(work, ignore_errors=True)
