"""Export PnP PPTX (and the prep instruction) to PDF via LibreOffice."""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"

PPTX_STEMS = (
    "svet-moy-zerkalce-rules",
    "svet-moy-zerkalce-cards",
    "svet-moy-zerkalce-card-backs",
    "svet-moy-zerkalce-votes",
    "svet-moy-zerkalce-letters",
)

SOFFICE_CANDIDATES = (
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
)


def _soffice() -> Path:
    for path in SOFFICE_CANDIDATES:
        if path.is_file():
            return path
    found = shutil.which("soffice")
    if found:
        return Path(found)
    raise FileNotFoundError("LibreOffice soffice.exe not found")


def export(sources: list[Path]) -> list[Path]:
    missing = [p for p in sources if not p.is_file()]
    if missing:
        raise FileNotFoundError("missing: " + ", ".join(str(p) for p in missing))
    OUT.mkdir(parents=True, exist_ok=True)
    profile = OUT / ".lo-profile"
    profile.mkdir(parents=True, exist_ok=True)
    uri = profile.resolve().as_uri()
    cmd = [
        str(_soffice()),
        f"-env:UserInstallation={uri}",
        "--headless",
        "--norestore",
        "--nologo",
        "--nolockcheck",
        "--convert-to",
        "pdf",
        "--outdir",
        str(OUT),
        *[str(p) for p in sources],
    ]
    subprocess.run(cmd, check=True)
    pdfs = [OUT / f"{p.stem}.pdf" for p in sources]
    for pdf in pdfs:
        if not pdf.is_file():
            raise FileNotFoundError(f"PDF was not written: {pdf}")
    return pdfs


def build() -> list[Path]:
    sources = [OUT / f"{stem}.pptx" for stem in PPTX_STEMS]
    return export(sources)


if __name__ == "__main__":
    try:
        paths = build()
    except Exception as exc:
        print(exc, file=sys.stderr)
        sys.exit(1)
    for path in paths:
        print(f"Wrote {path}")
