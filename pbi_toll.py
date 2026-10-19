"""Toll Notification Process Monitoring (public Power BI report) -> tables -> Excel / Word."""
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from playwright.sync_api import sync_playwright
from pbi_decode import decode, to_date

from paths import APP as BASE
import os
# the report link: PBI_NOTIFICATION_URL when set (hosted), else the text file next to the program
NOTIF_URL = (os.environ.get("PBI_NOTIFICATION_URL") or (BASE / "pbi_notification_url.txt").read_text()).strip()
P = "GetProjectInformationForPowerBI."


def _find(captured, need):
    """Largest decoded response whose column names contain every `need` fragment."""
    best = None
    for body in captured:
        try:
            cols, rows, complete = decode(body)
        except Exception:
            continue
        joined = " | ".join(cols)
        if all(n in joined for n in need):
            if best is None or len(rows) > len(best[1]):
                best = (cols, rows, complete)
    return best


def fetch_notification(log=print):
    captured = []
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        ctx = b.new_context(viewport={"width": 1900, "height": 2600})
        pg = ctx.new_page()

        def on_resp(r):
            if "querydata" in r.url and r.request.method == "POST":
                try:
                    captured.append(r.body())
                except Exception:
                    pass

        pg.on("response", on_resp)
        log("Opening Toll Notification report...")
        pg.goto(NOTIF_URL, wait_until="domcontentloaded", timeout=90000)
        pg.wait_for_timeout(18000)
        b.close()
    log(f"Captured {len(captured)} data responses")

    t = {}
    r = _find(captured, ["NotificationStartedCount > 75", "data.ro"])
    u = _find(captured, ["IsEOfficeUpdated", "IsSubmittedByPIU", "IsSOPermanentUpdated", "data.ro"])
    l = _find(captured, ["project_name", "physical_progress"])
    for name, got in (("progress>85", r), ("under-process RO wise", u), ("project list", l)):
        if not got:
            raise RuntimeError(f"Could not find the '{name}' table in the report data")
    if not l[2] or len(l[1]) >= 500:
        log("Warning: project list may be truncated by Power BI paging")

    rows = r[1]
    ro = [x for x in rows if not x["_total"] and x.get(P + "data.ro")]
    tot = next((x for x in rows if x["_total"]), None)

    def g85(x):
        return [x.get(P + "data.ro", "Total"),
                x.get("Min(" + P + "data.upc)"), x.get(P + "NotStartedCount"),
                x.get(P + "NotificationStartedCount > 75"), x.get(P + "GazettePublishCount >75")]

    t["gt85"] = {"cols": ["RO", "No. of Projects", "Project Notification Not Started",
                          "Project Notification Started", "Gazette Published"],
                 "rows": [g85(x) for x in ro], "total": g85(tot) if tot else None}

    rows = u[1]
    ro = [x for x in rows if not x["_total"] and x.get(P + "data.ro")]
    tot = next((x for x in rows if x["_total"]), None)

    def gu(x):
        return [x.get(P + "data.ro", "Total"),
                x.get("Min(" + P + "data.upc)"),
                x.get("Sum(" + P + "IsSubmittedByPIU new)"),
                x.get("Sum(" + P + "IsEOfficeUpdated new)"),
                x.get("Sum(" + P + "IsSOPermanentUpdated new)")]

    t["under"] = {"cols": ["RO", "No. of Projects Notification Started", "File Within NHAI",
                           "File Forwarded to MoRTH", "Gazette Published"],
                  "rows": [gu(x) for x in ro], "total": gu(tot) if tot else None}
    tv = gu(tot) if tot else [None] * 5
    t["summary"] = {"cols": ["Metric", "Count"],
                    "rows": [["Notification Under Process", tv[1]], ["File Within NHAI", tv[2]],
                             ["File Forwarded to MoRTH", tv[3]], ["Gazette Published", tv[4]]]}

    keys = [P + "data.project_name", P + "data.upc", P + "data.ro", P + "data.piu",
            P + "data.start_appointed_date", P + "data.schedule_completion_date", P + "data.pcod_date",
            "Sum(" + P + "data.physical_progress)", P + "data.current_project_stage",
            P + "dESCRIPTION cOLUMN", "Sum(" + P + "e file PIU)", "Sum(" + P + "e file co)"]
    heads = ["Project Name", "UPC", "RO", "PIU", "Appointed Date", "Schedule Completion Date",
             "PCOD Date", "Progress %", "Stage", "Notification Description",
             "e-File No. by PIU", "e-File No. by CO"]
    out = []
    for x in l[1]:
        v = [x.get(k) for k in keys]
        for i in (4, 5, 6):
            v[i] = to_date(v[i])
        out.append(v)
    t["projects"] = {"cols": heads, "rows": out}
    return t


# ---------------- exports ----------------
def _num(v):
    return v if isinstance(v, (int, float)) else (0 if v is None else v)


def to_excel(t, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    wb.remove(wb.active)
    sheets = [("Progress above 85%", t["gt85"]), ("Under Process - Summary", t["summary"]),
              ("Under Process - RO wise", t["under"]), ("Project List", t["projects"])]
    for name, tb in sheets:
        ws = wb.create_sheet(name[:31])
        ws.append(tb["cols"])
        for r in tb["rows"]:
            ws.append(["" if v is None else v for v in r])
        if tb.get("total"):
            ws.append(["" if v is None else v for v in tb["total"]])
            for c in ws[ws.max_row]:
                c.font = Font(bold=True)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0B3A6E")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in ws.columns:
            w = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(12, w + 2), 60)
        ws.freeze_panes = "A2"
    wb.save(path)


def narrative(t):
    g, s = t["gt85"], t["summary"]["rows"]
    gt = dict(zip(g["cols"], g["total"]))
    return [
        f"Projects with progress greater than 85%: there are {_num(gt['No. of Projects'])} such projects. "
        f"Toll notification has not started for {_num(gt['Project Notification Not Started'])}, has started for "
        f"{_num(gt['Project Notification Started'])}, and the gazette is published for {_num(gt['Gazette Published'])}.",
        f"Notification under process: {_num(s[0][1])} projects. {_num(s[1][1])} files are within NHAI, "
        f"{_num(s[2][1])} are forwarded to MoRTH and {_num(s[3][1])} have the gazette published.",
    ]


def ro_sentences(tb):
    out = []
    for r in tb["rows"]:
        parts = [f"{c.lower()} {_num(v)}" for c, v in zip(tb["cols"][1:], r[1:]) if v not in (None, 0, "")]
        out.append(f"{r[0]}: " + (", ".join(parts) if parts else "no projects") + ".")
    return out


def to_word(t, path):
    from docx import Document
    d = Document()
    d.add_heading("Toll Notification Process Monitoring", 0)
    d.add_paragraph(f"Source: NHAI CO Division Power BI report. Generated {datetime.now():%d-%m-%Y %H:%M}.")
    d.add_heading("Summary", 1)
    for ln in narrative(t):
        d.add_paragraph(ln)

    def table(tb):
        tbl = d.add_table(rows=1, cols=len(tb["cols"]))
        tbl.style = "Light Grid Accent 1"
        for i, c in enumerate(tb["cols"]):
            tbl.rows[0].cells[i].text = c
        for r in tb["rows"] + ([tb["total"]] if tb.get("total") else []):
            cells = tbl.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = "" if v is None else str(v)

    d.add_heading("1. Projects with progress greater than 85%", 1)
    table(t["gt85"])
    d.add_heading("In words, RO wise", 2)
    for ln in ro_sentences(t["gt85"]):
        d.add_paragraph(ln, style="List Bullet")

    d.add_heading("2. Notification under process", 1)
    d.add_heading("Summary", 2)
    table(t["summary"])
    d.add_heading("RO-wise break-up", 2)
    table(t["under"])
    d.add_heading("In words, RO wise", 2)
    for ln in ro_sentences(t["under"]):
        d.add_paragraph(ln, style="List Bullet")

    # project list: landscape section, compact table, header repeats on every page
    from docx.enum.section import WD_SECTION, WD_ORIENT
    from docx.shared import Pt, Cm
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    sec = d.add_section(WD_SECTION.NEW_PAGE)
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(1.2))
    d.add_heading(f"3. Project list ({len(t['projects']['rows'])} projects)", 1)
    pt = t["projects"]
    heads = ["#"] + pt["cols"]
    widths = [0.8, 6.0, 2.6, 2.2, 2.0, 1.8, 1.8, 1.8, 1.3, 2.2, 2.8, 1.5, 1.5]  # cm, ~27.7 total
    tbl = d.add_table(rows=1, cols=len(heads))
    tbl.style = "Table Grid"
    tbl.autofit = False
    for i, h in enumerate(heads):
        c = tbl.rows[0].cells[i]
        c.text = ""
        run = c.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(8)
        sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), "D9E2F3")
        c._tc.get_or_add_tcPr().append(sh)
    trpr = tbl.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trpr.append(th)
    for n, r in enumerate(pt["rows"], 1):
        cells = tbl.add_row().cells
        vals = [n] + ["" if v is None else v for v in r]
        for i, v in enumerate(vals):
            cells[i].text = ""
            para = cells[i].paragraphs[0]
            run = para.add_run(f"{v:.2f}" if i == 8 and isinstance(v, (int, float)) else str(v))
            run.font.size = Pt(7.5)
    for row in tbl.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    d.save(path)


def run(outdir, log=print):
    outdir = Path(outdir)
    outdir.mkdir(exist_ok=True)
    t = fetch_notification(log)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    x = outdir / f"{stamp}_TollNotification.xlsx"
    w = outdir / f"{stamp}_TollNotification.docx"
    to_excel(t, x)
    to_word(t, w)
    return t, [x, w]


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    t, files = run(BASE / "downloads")
    for k in ("gt85", "summary", "under"):
        print(k, t[k]["cols"])
        for r in t[k]["rows"]:
            print("  ", r)
        print("  total", t[k].get("total"))
    print("projects", len(t["projects"]["rows"]))
    print(files)
