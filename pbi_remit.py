"""Toll Plaza Windfall (public Power BI reports) -> filtered tables + the entire daily list -> Excel.

Steps mirror what a person does:
  1. Windfall tab -> Select Period = Last 1 Weeks -> Category = 40% and above
  2. Keep: RO Wise Data, Tolls with Windfall Gain (plaza list). The first "Please select any one plaza"
     table (contract dates) is skipped on purpose.
  3. The database icon at the last table opens a second public report holding the entire list.
     Power BI serves it in windows of 500 rows, so every window is fetched.
"""
import copy
import json
import re
import sys
import uuid
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from playwright.sync_api import sync_playwright
from pbi_decode import decode, to_date

from paths import APP as BASE
import os
# the report links: PBI_REMITTANCE_URL / PBI_REMITTANCE_DATA_URL when set (hosted), else the text files
MAIN_URL = (os.environ.get("PBI_REMITTANCE_URL") or (BASE / "pbi_remittance_url.txt").read_text()).strip()
LIST_URL = (os.environ.get("PBI_REMITTANCE_DATA_URL") or (BASE / "pbi_remittance_data_url.txt").read_text()).strip()
PAGE = 500


def _find(captured, need):
    best = None
    for body in captured:
        try:
            cols, rows, complete = decode(body)
        except Exception:
            continue
        j = " | ".join(cols)
        if all(n in j for n in need) and (best is None or len(rows) > len(best[1])):
            best = (cols, rows, complete)
    return best


def _short(c):
    c = c.split(".")[-1]
    return c[:-1] if c.endswith(")") else c


def _row(x, cols):
    return [x.get(c) for c in cols]


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def fetch_main(log):
    cap = []
    period = ""
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        ctx = b.new_context(viewport={"width": 1900, "height": 2000})
        pg = ctx.new_page()
        pg.on("response", lambda r: cap.append(r.body())
              if ("querydata" in r.url and r.request.method == "POST") else None)
        log("Opening Toll Plaza Windfall report...")
        pg.goto(MAIN_URL, wait_until="domcontentloaded", timeout=90000)
        pg.wait_for_timeout(14000)
        pg.get_by_text("Windfall Gain", exact=False).nth(0).click()
        pg.wait_for_timeout(6000)
        log("Select Period: Last 1 Weeks")
        # the third control in the Select Period slicer is the unit dropdown
        pg.locator(".relative-slicer .ui-selectmenu-text:visible").filter(
            has_text=re.compile(r"^Select$")).first.click()
        pg.wait_for_timeout(1500)
        pg.get_by_role("option", name="Weeks", exact=True).first.click()
        pg.wait_for_timeout(4000)
        log("Category: 40% and above")
        cap.clear()
        pg.get_by_role("checkbox", name="40% and above").first.click()
        pg.wait_for_timeout(7000)
        m = re.search(r"\d+/\d+/\d{4}\s*-\s*\d+/\d+/\d{4}", pg.inner_text("body"))
        period = m.group(0) if m else ""
        b.close()

    ro = _find(cap, ["RO_ADDED", "Toll Plazas with Windfall"])
    pl = _find(cap, ["plazacode", "plazaname", "Category_Flag_WF"])
    dy = _find(cap, ["plazaname", "Percent_Gain", "Cumulative_Actual_Collection"])
    for name, got in (("RO wise", ro), ("plaza list", pl), ("daily (selected view)", dy)):
        if not got:
            raise RuntimeError(f"Could not find the '{name}' table in the report data")

    t = {"period": period}
    rows = [x for x in ro[1] if not x["_total"] and x.get("Toll_Master.RO_ADDED")]
    key_ro, key_n = "Toll_Master.RO_ADDED", "Canara_Summary_Plazacodes_Shortfall.Toll Plazas with Windfall"
    t["ro"] = {"cols": ["RO", "Toll Plazas with Windfall"],
               "rows": [[x[key_ro], x.get(key_n)] for x in rows]}
    t["ro"]["total"] = ["Total", sum(r[1] or 0 for r in t["ro"]["rows"])]

    pc = ["Toll_Master.plazacode", "Toll_Master.plazaname", "Toll_Master.current_agency_name",
          "Toll_Master.Category_Flag_WF", "Toll_Master.bid_type_tex_filter",
          "Toll_Master.Value_in_Crores_Agreed_Remittance"]
    t["plazas"] = {"cols": ["Plaza Code", "Plaza Name", "Current Agency Name", "Category", "Bid Type",
                            "Agreed Remittance (Avg. per Day)"],
                   "rows": [_row(x, pc) for x in pl[1] if not x["_total"]]}

    S = "toll_revenue_api_summary."
    dc = ["Toll_Master.plazaname", S + "date", "Sum(" + S + "Agreed_Remittance)",
          "Sum(" + S + "Cumulative_Agreed_Remittance)", "Sum(" + S + "revenue)",
          "Sum(" + S + "Cumulative_Actual_Collection)", "Sum(" + S + "15DayMovingAverage)",
          "Sum(" + S + "Percent_Gain)"]
    out = []
    for x in dy[1]:
        if x["_total"]:
            continue
        v = _row(x, dc)
        v[1] = to_date(v[1])
        v[7] = round(_num(v[7]) * 100, 2) if isinstance(_num(v[7]), float) else v[7]
        out.append(v)
    t["daily"] = {"cols": ["Plaza Name", "Date", "Agreed Remittance", "Cumulative Agreed Remittance",
                           "ETC + Cash by IHMCL Data", "Cumulative Actual Collection",
                           "Actual Collection (15 Days Moving Avg)", "Percent Gain (%)"],
                  "rows": out}
    return t


def fetch_entire_list(log):
    """The report behind the database icon: every plaza x date row."""
    first = {}
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        ctx = b.new_context(viewport={"width": 1900, "height": 1500})
        pg = ctx.new_page()

        def on_resp(r):
            if "querydata" in r.url and r.request.method == "POST":
                try:
                    body = r.body()
                    cols, rows, _ = decode(body)
                    if "Cumulative_Agreed_Remittance" in " ".join(cols) and len(rows) > 50:
                        first.update(url=r.url, headers=r.request.headers, req=r.request.post_data, body=body)
                except Exception:
                    pass

        pg.on("response", on_resp)
        log("Opening the entire-list report...")
        pg.goto(LIST_URL, wait_until="domcontentloaded", timeout=90000)
        pg.wait_for_timeout(15000)
        if not first:
            raise RuntimeError("The entire-list table did not load")
        req0 = json.loads(first["req"])
        hdr = {k: v for k, v in first["headers"].items()
               if k in ("x-powerbi-resourcekey", "content-type", "accept", "referer")}
        res = json.loads(first["body"])
        cols = None
        allrows = []
        n = 0
        while True:
            n += 1
            cols, rows, _ = decode(res)
            allrows += [x for x in rows if not x["_total"]]
            ds = res["results"][0]["result"]["data"]["dsr"]["DS"][0]
            log(f"Entire list: window {n}, {len(allrows)} rows so far")
            if "RT" not in ds:
                break
            rq = copy.deepcopy(req0)
            rq["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]["Binding"][
                "DataReduction"]["Primary"]["Window"] = {"Count": PAGE, "RestartTokens": ds["RT"]}
            h = dict(hdr)
            h["requestid"], h["activityid"] = str(uuid.uuid4()), str(uuid.uuid4())
            r = ctx.request.post(first["url"], data=json.dumps(rq), headers=h)
            if r.status != 200:
                raise RuntimeError(f"Power BI refused window {n + 1} (HTTP {r.status})")
            res = json.loads(r.body())
            if n > 400:
                raise RuntimeError("Too many windows, stopping")
        b.close()

    keep = [c for c in cols if not _short(c).startswith("plazacode") or c.endswith("data.plazacode") or c.endswith(".plazacode")]
    keep = [c for c in cols if not c.startswith("Count(")]
    names = [_short(c) for c in keep]
    di = names.index("date")
    pi = names.index("Percent_Gain") if "Percent_Gain" in names else None
    seen, out = set(), []
    ki = names.index("key") if "key" in names else None
    for x in allrows:
        v = _row(x, keep)
        if ki is not None:
            if v[ki] in seen:
                continue
            seen.add(v[ki])
        v[di] = to_date(v[di])
        if pi is not None and isinstance(_num(v[pi]), float):
            v[pi] = round(_num(v[pi]) * 100, 2)
        out.append(v)
    names = ["Percent Gain (%)" if n == "Percent_Gain" else n for n in names]
    return {"cols": names, "rows": out}


def filter_entire(t, entire):
    """Keep only the plazas selected by the Windfall filters (week + 40% and above)."""
    names = {str(r[0]): r[1] for r in t["plazas"]["rows"]}
    ci = entire["cols"].index("plazacode")
    rows = []
    for r in entire["rows"]:
        if str(r[ci]) in names:
            rows.append(r[:ci + 1] + [names[str(r[ci])]] + r[ci + 1:])
    cols = entire["cols"][:ci + 1] + ["Plaza Name"] + entire["cols"][ci + 1:]
    found = {str(r[ci]) for r in rows}
    missing = [f"{names[c]} ({c})" for c in names if c not in found]
    di = entire["cols"].index("date")
    last = max((datetime.strptime(r[di + 1], "%d-%m-%Y") for r in rows), default=None)
    return {"cols": cols, "rows": rows}, missing, last


def to_excel(t, entire, path, missing=(), last=None, total_rows=0):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Filters"
    ws.append(["Toll Plaza Windfall / Shortfall - Windfall Gain tab"])
    ws.append(["Select Period", "Last 1 Weeks" + (f"  ({t['period']})" if t["period"] else "")])
    ws.append(["Category", "40% and above"])
    ws.append(["Generated", datetime.now().strftime("%d-%m-%Y %H:%M")])
    ws.append(["Entire list", f"Taken from the report behind the database icon ({total_rows:,} rows for all plazas) "
                              f"and cut down to the plazas that pass the two filters ({len(t['plazas']['rows'])} plazas)."])
    ws.append(["Data limit", "That report holds data only up to "
                             + (last.strftime("%d-%m-%Y") if last else "n/a")
                             + ", so it has no rows for the selected week. Sheet 'Daily - selected week' "
                               "holds the week's rows from the main report."])
    if missing:
        ws.append(["Not in that report", ", ".join(missing)])
    ws["A1"].font = Font(bold=True, size=13)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).font = Font(bold=True)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 100

    def sheet(name, tb, total=None):
        s = wb.create_sheet(name[:31])
        s.append(tb["cols"])
        for r in tb["rows"]:
            s.append(["" if v is None else v for v in r])
        if total:
            s.append(total)
            for c in s[s.max_row]:
                c.font = Font(bold=True)
        for c in s[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0B3A6E")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in s.columns:
            w = max(len(str(c.value or "")) for c in list(col)[:300])
            s.column_dimensions[col[0].column_letter].width = min(max(12, w + 2), 50)
        s.freeze_panes = "A2"

    sheet("RO Wise Data", t["ro"], t["ro"]["total"])
    sheet("Tolls with Windfall Gain", t["plazas"])
    sheet("Daily - selected week", t["daily"])
    sheet("Entire list (filtered)", entire)
    wb.save(path)


def run(outdir, log=print):
    outdir = Path(outdir)
    outdir.mkdir(exist_ok=True)
    t = fetch_main(log)
    full = fetch_entire_list(log)
    entire, missing, last = filter_entire(t, full)
    f = outdir / f"{datetime.now():%Y%m%d_%H%M%S}_TollPlazaWindfall.xlsx"
    to_excel(t, entire, f, missing, last, len(full["rows"]))
    entire["missing"], entire["last"], entire["fetched"] = missing, last, len(full["rows"])
    return t, entire, [f]


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    t, entire, files = run(BASE / "downloads")
    print("period:", t["period"])
    print("RO rows:", t["ro"]["rows"], t["ro"]["total"])
    print("plazas:", len(t["plazas"]["rows"]), "daily:", len(t["daily"]["rows"]))
    print("entire:", len(entire["rows"]), entire["cols"])
    print(entire["rows"][0])
    print(files)
