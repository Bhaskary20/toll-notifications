import os, sys, json
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv()
URL = "https://datalakeg.nhai.gov.in/nhai/"

with sync_playwright() as p:
    b = p.chromium.launch(headless=False)
    ctx = b.new_context(accept_downloads=True)
    pg = ctx.new_page()
    pg.goto(URL, wait_until="networkidle", timeout=60000)
    pg.select_option("#ddlUserType", "2")  # HQ Technical Division
    pg.click("#userName"); pg.fill("#userName", os.environ["NHAI_USER"])
    pg.click("#password"); pg.fill("#password", os.environ["NHAI_PASS"])
    pg.focus("#txtInput")
    print(">>> Type the captcha in the browser window and click SIGN IN (10 min limit).", flush=True)
    pg.wait_for_function("() => !document.querySelector('#loginBtn')", timeout=600000)
    pg.wait_for_load_state("networkidle")
    print("Logged in. URL:", pg.url, "| title:", pg.title(), flush=True)
    ctx.storage_state(path="auth_state.json")
    pg.screenshot(path="after_login.png", full_page=True)
    links = pg.eval_on_selector_all("a", "els => els.map(e => [e.innerText.trim().slice(0,60), e.href])")
    with open("after_login_links.json", "w", encoding="utf-8") as f:
        json.dump(links, f, indent=1, ensure_ascii=False)
    print(len(links), "links saved to after_login_links.json")
    b.close()
