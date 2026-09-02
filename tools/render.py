#!/usr/bin/env python3
"""pptx 를 슬라이드별 PNG 로 렌더링한다 (LibreOffice Impress + poppler).

  python3 tools/render.py output/deck.pptx output/preview  [--dpi 100]
"""
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PROFILE = ROOT / ".render_profile"


def render(pptx, outdir, dpi=100):
    pptx, outdir = Path(pptx).resolve(), Path(outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    tmp = outdir / "_pdf"
    if tmp.exists():
        shutil.rmtree(tmp)
    tmp.mkdir()
    subprocess.run(
        ["soffice", "--headless", "--norestore",
         f"-env:UserInstallation=file://{PROFILE}",
         "--convert-to", "pdf", "--outdir", str(tmp), str(pptx)],
        check=True, capture_output=True,
    )
    pdf = next(tmp.glob("*.pdf"))
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(outdir / "slide")],
                   check=True, capture_output=True)
    shutil.rmtree(tmp)
    pngs = sorted(outdir.glob("slide-*.png"))
    for p in pngs:
        print(p)
    return pngs


if __name__ == "__main__":
    dpi = 100
    if "--dpi" in sys.argv:
        dpi = int(sys.argv[sys.argv.index("--dpi") + 1])
    render(sys.argv[1], sys.argv[2], dpi)
