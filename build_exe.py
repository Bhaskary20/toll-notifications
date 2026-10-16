"""Builds the shareable folder  dist\\CO Report Tool\\  (run:  py build_exe.py).

The folder holds the program, the report template and settings, and the browser the tool drives, so it works on a
laptop with nothing installed. Zip the whole folder to send it. Secrets (.env, auth_state.json) and old downloads are
never copied."""
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "CO Report Tool"
DIST = ROOT / "dist"
OUT = DIST / NAME
WORK = ROOT / "build"
DATA = ["static", "templates/CO_KPI_template.docx", "report_config.json", "report_history.json",
        "pbi_notification_url.txt", "pbi_remittance_url.txt", "pbi_remittance_data_url.txt"]

README = """CO Division Report Tool
=======================

1. Unzip this whole folder anywhere (for example on the Desktop). Do not move the .exe out of it.
2. Double-click "CO Report Tool.exe". A black window opens and the dashboard opens in your browser.
   Keep the black window open while you work; close it to stop the tool.
3. First time only: type your NHAI Data Lake user name and password (HQ Technical Division). They are saved on
   this computer only.
4. Click "Connect" at the bottom of the sidebar (or "Connect to the portal" on the overview) and type the captcha
   shown. Then click "Run weekly report", or open "Weekly KPI report" in the sidebar and click "Generate .docx".
   Press Ctrl+K to find any task, and ? to see the keyboard shortcuts.

Reports and the source files they are built from are downloaded by your browser to your own Downloads folder.
The tool keeps no copy of them.

Where things are (all next to the .exe)
- templates\\         CO_KPI_template.docx, the report layout. Edit the manual sections here (Toll Plaza Overview,
                      BoT to PF, empanelment, FASTag, Figure 1, Annexure Tables 3 and 4); they are kept as written.
- report_history.json  the daily numbers behind the trend columns. Keep it; the tool adds a line every run.
- report_config.json   baseline date and HQ remarks.
- .env                 your saved portal login (created on first use). Never send this file to anyone.

Tips
- If the browser asks whether the site may download multiple files, choose Allow (the weekly report downloads
  the report and a zip of its source data).
- Windows may warn "unknown publisher" the first time: choose More info > Run anyway.
- The report is always "as on today"; the portal only shows current data.
"""


def run(*cmd):
    print(">", " ".join(map(str, cmd)))
    subprocess.run(cmd, check=True)


def main():
    shutil.rmtree(OUT, ignore_errors=True)
    run(sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir", "--name", NAME,
        "--distpath", str(DIST), "--workpath", str(WORK), "--specpath", str(WORK),
        "--collect-all", "playwright", "--exclude-module", "fitz", "--exclude-module", "pymupdf",
        "--hidden-import", "paths", "--hidden-import", "report_builder", "--hidden-import", "pbi_toll",
        "--hidden-import", "pbi_remit", "--hidden-import", "pbi_decode", "--paths", str(ROOT),
        str(ROOT / "launcher.py"))
    for item in DATA:
        src, dst = ROOT / item, OUT / item
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(src, dst) if src.is_dir() else shutil.copy2(src, dst)
    pw = Path(os.environ["LOCALAPPDATA"]) / "ms-playwright"
    shell = sorted(pw.glob("chromium_headless_shell-*"))
    if not shell:
        sys.exit("No Playwright headless browser installed; run: py -m playwright install chromium")
    (OUT / "browsers").mkdir(exist_ok=True)
    shutil.copytree(shell[-1], OUT / "browsers" / shell[-1].name)
    (OUT / "READ ME FIRST.txt").write_text(README, encoding="utf-8")
    print("Built:", OUT)


if __name__ == "__main__":
    main()
