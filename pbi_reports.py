import sys, json, os
from pathlib import Path
from playwright.sync_api import sync_playwright
sys.stdout.reconfigure(encoding="utf-8")
SCR = Path(os.environ["TEMP"]) / "pbi_capture"; SCR.mkdir(exist_ok=True)
url = open("pbi_url.txt").read().strip()
TILES = {"toll_notification": (360, 300), "toll_remittance": (360, 660)}
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    for name, (x, y) in TILES.items():
        ctx = b.new_context(viewport={"width":1500,"height":950})
        home = ctx.new_page(); home.goto(url, wait_until="domcontentloaded", timeout=90000); home.wait_for_timeout(9000)
        with ctx.expect_page(timeout=30000) as np_:
            home.mouse.click(x, y)
        rep = np_.value
        rurl = rep.url
        calls = []
        def on_resp(r, calls=calls):
            if "querydata" in r.url and r.request.method == "POST":
                try: calls.append((r.request.post_data, r.body()))
                except Exception: pass
        rep.on("response", on_resp)
        rep.reload(wait_until="domcontentloaded"); rep.wait_for_timeout(16000)
        open(SCR / f"{name}_url.txt", "w").write(rurl)
        rep.screenshot(path=f"pbi_{name}.png")
        print("=====", name, "|", rep.title(), "| querydata calls:", len(calls))
        pages = rep.evaluate("() => (document.body.innerText.match(/\d+\s*of\s*\d+/)||['?'])[0]")
        print("page counter:", pages)
        for i, (req, body) in enumerate(calls):
            (SCR / f"{name}_{i}.req.json").write_bytes((req or "").encode()); (SCR / f"{name}_{i}.res.json").write_bytes(body)
            try:
                q = json.loads(req)["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]["Query"]
                cols = [s.get("Name") for s in q.get("Select", [])]
            except Exception: cols = ["?"]
            print(f"  [{i}] {len(body):>6}B  cols={cols[:8]}")
        ctx.close()
    b.close()
