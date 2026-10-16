"""Test bench for the dashboard page. Never contacts the NHAI portal.

  py tools/mock_server.py                 serve the page with fake data on http://127.0.0.1:5077
  py tools/mock_server.py --shots         also save screenshots to tools/shots/ and exit
  py tools/mock_server.py --state idle    pick the session state: idle | captcha_required | logged_in
  py tools/mock_server.py --data empty    pick the KPI data: full | empty | stale

The numbers are a real summary built from the downloads of 14 Oct 2026.
"""
import argparse
import json
import threading
import time
from pathlib import Path

from flask import Flask, jsonify, send_from_directory

ROOT = Path(__file__).resolve().parent.parent
SHOTS = Path(__file__).resolve().parent / "shots"

ap = argparse.ArgumentParser()
ap.add_argument("--state", default="logged_in")
ap.add_argument("--data", default="full")
ap.add_argument("--port", type=int, default=5077)
ap.add_argument("--shots", action="store_true")
args = ap.parse_args()

KPI = {
    "file": {"name": "CO Division KPIs Status Update Report 14102026.docx", "size": 170208},
    "fmc": dict(total=32, piu=20, agency=8, ro=3, hq=1, closed=0, pendency=12, piu_share=62, agency_pendency=8, ro_pendency=12),
    "bg": dict(total=45, piu=41, agency=3, ro=0, hq=0, closed=1, pendency=15, piu_share=91, agency_pendency=13, ro_pendency=0),
    "notice": {"total": 14, "rect": 0, "open": 14, "types": {"Toilets at Fee Plaza": 11, "Admin Building/100 mtr either side of Toll Plaza": 3},
               "agencies": 13, "pius": 9, "pendency": 40, "over30": 4, "max_days": 279, "open_share": 100},
    "ongoing": {"total": 742, "regular": 597, "eq": 143, "manpower": 2, "top10_total": 296, "agencies": 121, "rows": 848, "blank_dates": 76, "duplicates": 30,
                "top10": [{"agency": a, "total": t} for a, t in [("Coral Associates", 46), ("Skylark Infra Engineering Pvt. Ltd.", 44),
                          ("Ashmi Road Carriers Pvt. Ltd.", 38), ("Riddhi Siddhi Associates", 35), ("Rakshak Securitas Pvt. Ltd.", 32),
                          ("Innovision Ltd.", 27), ("Y Meenakshi", 24), ("Hule Construction (P) ltd", 19), ("Md. Yousuf Ali", 16),
                          ("Surendra Kumar Shukla", 15)]]},
    "issues": [], "date": "2026-10-14", "saved_at": "14 Oct 17:59",
}
NOTIF = {"files": [{"name": "Toll Notification.xlsx", "size": 41000}], "summary": [["Notification Under Process", 37], ["File Within NHAI", 12],
         ["File Forwarded to MoRTH", 9], ["Gazette Published", 4]], "gt85_total": 20, "projects": 61, "date": "2026-10-14", "saved_at": "14 Oct 17:40"}
WIND = {"files": [{"name": "TollPlazaWindfall.xlsx", "size": 71361}], "period": "08-10-2026 to 14-10-2026", "ro_total": 112, "plazas": 112, "daily": 640,
        "entire": 12408, "missing": [], "last": "14-10-2026", "date": "2026-10-14", "saved_at": "14 Oct 17:45"}
HIST = json.loads((ROOT / "report_history.json").read_text(encoding="utf-8"))
# newest snapshot carries the new fields the weekly report now writes
HIST["2026-10-14"].update(fmc_pending=32, bg_pending=44, notice_open=14, fmc_wait=12, bg_wait=15)

app = Flask(__name__)
downloads = ROOT / "downloads"
files = [{"name": p.name, "size": p.stat().st_size, "time": time.strftime("%d %b %H:%M", time.localtime(p.stat().st_mtime))}
         for p in sorted(downloads.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True) if p.is_file()] if downloads.exists() else []
state = {"s": args.state}


@app.get("/")
def index():
    return send_from_directory(ROOT / "static", "index.html")


@app.get("/api/status")
def status():
    on = state["s"] == "logged_in"
    return jsonify(state=state["s"], message="", user="GMC", has_credentials=True, captcha_version=1, files=files,
                   nav=[{"title": f"Section {i}", "url": f"https://datalakeg.nhai.gov.in/nhai/x{i}"} for i in range(12)] if on else [], scan=None,
                   logs=[{"t": f"17:59:0{i}", "m": m} for i, m in enumerate(["Logged in: Dashboard", "Found 42 portal sections.",
                         "Exporting EXCEL from FMC Cases...", "Saved CO Division KPIs Status Update Report 14102026.docx"])] if on else [])


@app.get("/api/dashboard")
def dashboard():
    if args.data == "empty":
        return jsonify(ok=True, kpi=None, notification=None, windfall=None, history=[], today="2026-10-15")
    kpi = dict(KPI, date="2026-10-01") if args.data == "stale" else KPI
    return jsonify(ok=True, kpi=kpi, notification=NOTIF, windfall=WIND,
                   history=[{"date": d, **v} for d, v in sorted(HIST.items())], today="2026-10-15" if args.data != "stale" else "2026-10-15")


@app.post("/api/mod/filters")
def mod_filters():
    return jsonify(ok=True, agencies=[{"value": "1", "text": "Coral Associates"}], info="Showing 1 to 10 of 1,234 entries")


@app.post("/api/co/notification")
def fake_notification():
    time.sleep(3)
    return jsonify(ok=True, **NOTIF)


@app.post("/api/report/build")
def fake_build():
    time.sleep(4)
    return jsonify(ok=True, **KPI)


def shoot():
    from playwright.sync_api import sync_playwright
    SHOTS.mkdir(exist_ok=True)
    base = f"http://127.0.0.1:{args.port}/"
    with sync_playwright() as p:
        b = p.chromium.launch()
        errs = []
        for name, size in (("1440", (1440, 900)), ("1366", (1366, 768)), ("390", (390, 844))):
            pg = b.new_page(viewport={"width": size[0], "height": size[1]})
            pg.on("pageerror", lambda e: errs.append(str(e)))
            pg.on("console", lambda m: m.type == "error" and errs.append(m.text))
            for v in ("overview", "kpi", "reports", "co", "library", "explore"):
                pg.goto(base + "#" + v)
                pg.reload()
                pg.wait_for_timeout(2400)
                pg.screenshot(path=str(SHOTS / f"{name}_{v}.png"), full_page=(v == "overview"))
            if name == "1440":
                pg.goto(base + "#overview"); pg.reload(); pg.wait_for_timeout(1800)
                pg.click('.kpi[data-k="notice"]'); pg.wait_for_timeout(900); pg.screenshot(path=str(SHOTS / "1440_drawer_notice.png"))
                pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
                pg.click('.kpi[data-k="ongoing"]'); pg.wait_for_timeout(900); pg.screenshot(path=str(SHOTS / "1440_drawer_ongoing.png"))
                pg.keyboard.press("Escape")
                pg.hover(".chart"); pg.mouse.move(700, 560); pg.wait_for_timeout(500); pg.screenshot(path=str(SHOTS / "1440_chart_hover.png"))
                pg.keyboard.press("Control+k"); pg.wait_for_timeout(500); pg.screenshot(path=str(SHOTS / "1440_palette.png")); pg.keyboard.press("Escape")
            if name == "390":
                print("overflow@390:", pg.evaluate("document.documentElement.scrollWidth"))
            pg.close()
        print("errors:", errs or "none")
        b.close()


if __name__ == "__main__":
    if args.shots:
        threading.Thread(target=lambda: app.run(port=args.port, threaded=True), daemon=True).start()
        time.sleep(1.5)
        shoot()
    else:
        print(f"Mock dashboard: http://127.0.0.1:{args.port}  (state={args.state}, data={args.data})")
        app.run(port=args.port, threaded=True)
