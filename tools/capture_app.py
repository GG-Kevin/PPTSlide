#!/usr/bin/env python3
"""배포용 HTML 프로그램을 실제로 띄워 화면별 스크린샷을 만든다.

  python3 tools/capture_app.py <app.html> [--out assets/screens]

이 앱은 Tailwind 를 CDN 에서 받는데 오프라인/차단 환경에서는 로드되지 않아
스타일 없이 렌더된다. 그래서 Tailwind 를 로컬에서 빌드해 HTML 에 인라인한
사본을 만들어 캡처한다. 원본 파일은 건드리지 않는다.
"""
import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
CDN_TAG = '<script src="https://cdn.tailwindcss.com"></script>'


def localize_tailwind(src: Path, work: Path) -> Path:
    """Tailwind CDN 스크립트를 로컬 빌드 CSS 로 바꾼 사본을 만든다."""
    html = src.read_text(encoding="utf-8")
    out = work / "app_local.html"
    if CDN_TAG not in html:
        out.write_text(html, encoding="utf-8")
        return out
    work.mkdir(parents=True, exist_ok=True)
    shutil.copy(src, work / "app.html")
    (work / "tw-in.css").write_text(
        "@tailwind base;\n@tailwind components;\n@tailwind utilities;\n", encoding="utf-8")
    if not (work / "node_modules" / ".bin" / "tailwindcss").exists():
        subprocess.run(["npm", "i", "-D", "tailwindcss@3", "--silent"],
                       cwd=work, check=True, capture_output=True)
    subprocess.run([str(work / "node_modules" / ".bin" / "tailwindcss"),
                    "-i", "tw-in.css", "-o", "tw-out.css",
                    "--content", "./app.html", "--minify"],
                   cwd=work, check=True, capture_output=True)
    css = (work / "tw-out.css").read_text(encoding="utf-8")
    out.write_text(html.replace(CDN_TAG, f"<style>{css}</style>"), encoding="utf-8")
    return out



def click_by_text(pg, label):
    """버튼 텍스트로 클릭한다. 아이콘 문자가 섞여 있어도 부분일치로 찾는다."""
    ok = pg.evaluate("""(label) => {
      const norm = s => (s||'').replace(/\\s+/g,' ').trim();
      const els = [...document.querySelectorAll('button,[role=tab],a')];
      let hit = els.find(e => norm(e.innerText) === label)
             || els.find(e => norm(e.innerText).endsWith(label))
             || els.find(e => norm(e.innerText).includes(label));
      if (!hit) return false;
      hit.click(); return true;
    }""", label)
    if not ok:
        raise RuntimeError(f"'{label}' 버튼을 찾지 못함")


# (파일명, 설명, 클릭 순서) — 텍스트로 버튼을 찾는다
SHOTS = [
    ("lnl_dashboard",   "L&L 대시보드",        ["L&L", "Dashboard"]),
    ("lnl_review",      "L&L 프로젝트 사전검토", ["L&L", "Review for Project"]),
    ("lnl_browser",     "L&L 브라우저",         ["L&L", "Browser"]),
    ("lnl_detail",      "L&L 사례 상세",        ["L&L", "Browser", "CIVIL"]),
    ("lnl_lists",       "L&L 리스트",           ["L&L", "Lists"]),
    ("tm_log",          "TM Log 브라우저",      ["TM", "Log Browser"]),
    ("tm_master",       "TM Master 브라우저",   ["TM", "Master Browser"]),
    ("tm_trend",        "TM Trend 시트",        ["TM", "Trend Browser"]),
    ("tm_mywork",       "TM My Work",           ["TM", "My Work"]),
    ("risk_dashboard",  "Risk 커맨드 센터",     ["Risk", "Dashboard"]),
    ("risk_register",   "Risk Register",        ["Risk", "Risk Register"]),
    ("risk_actions",    "Risk Action Plan",     ["Risk", "Action Plan"]),
    ("master_lists",    "Master Lists",         ["L&L", "Master Lists"]),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("app")
    ap.add_argument("--out", default="assets/screens")
    ap.add_argument("--work", default=".capture")
    ap.add_argument("--width", type=int, default=1600)
    ap.add_argument("--height", type=int, default=1000)
    a = ap.parse_args()

    from playwright.sync_api import sync_playwright

    root = Path(__file__).resolve().parent.parent
    outdir = root / a.out
    outdir.mkdir(parents=True, exist_ok=True)
    local = localize_tailwind(Path(a.app).resolve(), root / a.work)
    print(f"로컬 tailwind 적용본: {local}")

    report = []
    with sync_playwright() as pw:
        b = pw.chromium.launch(executable_path=CHROME, args=["--no-sandbox"])
        pg = b.new_page(viewport={"width": a.width, "height": a.height},
                        device_scale_factor=2)
        for name, desc, clicks in SHOTS:
            try:
                pg.goto(local.as_uri())
                pg.wait_for_timeout(2600)
                for label in clicks:
                    click_by_text(pg, label)
                    pg.wait_for_timeout(1100)
                pg.wait_for_timeout(700)
                path = outdir / f"{name}.png"
                pg.screenshot(path=str(path))
                report.append({"name": name, "desc": desc, "file": str(path.relative_to(root)),
                               "ok": True})
                print(f"  ✔ {name:16s} {desc}")
            except Exception as e:
                report.append({"name": name, "desc": desc, "ok": False, "error": str(e)[:120]})
                print(f"  ✘ {name:16s} {desc} — {str(e)[:90]}")
        b.close()
    (outdir / "_capture_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n성공 {sum(1 for r in report if r['ok'])}/{len(report)}")
    return 0 if all(r["ok"] for r in report) else 1


if __name__ == "__main__":
    sys.exit(main())
