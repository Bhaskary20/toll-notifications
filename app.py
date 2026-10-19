"""NHAI Data Lake dashboard backend.

Run:  py app.py   ->  http://127.0.0.1:5055
Credentials are read from .env (NHAI_USER / NHAI_PASS) and never leave the server.
The captcha is shown on the dashboard and typed by the person using it.
The dashboard itself asks for its own sign-in (DASHBOARD_USER / DASHBOARD_PASS_HASH in .env override the default).
Files a task produces are never kept on this computer: they go straight to the browser that asked for them.
"""
import hmac
import io
import os
import re
import queue
import secrets
import shutil
import sys
import tempfile
import threading
import time
import zipfile
from contextlib import contextmanager
from datetime import timedelta
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote

from dotenv import load_dotenv, set_key
from flask import Flask, jsonify, redirect, request, send_file, send_from_directory, session, Response
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from werkzeug.security import check_password_hash

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import APP as BASE, FROZEN

ENV_FILE = BASE / ".env"
load_dotenv(ENV_FILE)

HOST = "datalakeg.nhai.gov.in"
ROOT = f"https://{HOST}/nhai/"
ROLE_VALUE = "2"  # HQ Technical Division
DOC_PATTERN = re.compile(r"download|\.(pdf|docx?|xlsx?|csv|zip|rar|pptx?)(\?|$)|getfile|viewdocument", re.I)

app = Flask(__name__, static_folder=str(BASE / "static"), static_url_path="/static")

# ---- dashboard sign-in --------------------------------------------------------
# Only a hash of the password is kept; the default account is Admin / Admin@123.
DASH_USER = os.environ.get("DASHBOARD_USER", "Admin")
DASH_HASH = os.environ.get("DASHBOARD_PASS_HASH") or (
    "scrypt:32768:8:1$55BzghRUT7cXEfHj$f708b91ea3b87e5fcd0160354686fa08732d20836b846294f211274ab0f58c1701483dd8926a9b4"
    "c3a142f0d7acde17d5e1c18fb076bc60cedc95ca6649d6345")
if not os.environ.get("DASHBOARD_SECRET"):
    # a fixed key per computer, so a sign-in survives server restarts
    ENV_FILE.touch(exist_ok=True)
    os.environ["DASHBOARD_SECRET"] = secrets.token_hex(32)
    set_key(str(ENV_FILE), "DASHBOARD_SECRET", os.environ["DASHBOARD_SECRET"], quote_mode="always")
app.secret_key = os.environ["DASHBOARD_SECRET"]
app.config.update(SESSION_COOKIE_NAME="tod_session", SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  PERMANENT_SESSION_LIFETIME=timedelta(days=30))
SIGNED_OUT_OK = {"login_page", "login"}
_fails, _fails_lock = {}, threading.Lock()
MAX_FAILS, LOCK_SECONDS = 5, 60


def allowed(url: str) -> bool:
    p = urlparse(url)
    return p.scheme == "https" and p.hostname == HOST


class Worker(threading.Thread):
    """Owns the Playwright browser; every browser action is serialised through its queue."""

    def __init__(self):
        super().__init__(daemon=True)
        self.q = queue.Queue()
        self.state = "idle"  # idle | starting | captcha_required | logging_in | logged_in | error
        self.message = ""
        self.captcha = None
        self.captcha_version = 0
        self.logs = []
        self.nav = []
        self.scan = None
        self.last_dialog = ""
        self.pw = self.browser = self.ctx = self.page = None

    # ---- plumbing -------------------------------------------------------
    def log(self, msg):
        self.logs.append({"t": time.strftime("%H:%M:%S"), "m": msg})
        self.logs = self.logs[-100:]

    def call(self, fn, *args, timeout=150):
        box = {"ev": threading.Event()}
        self.q.put((fn, args, box))
        if not box["ev"].wait(timeout):
            raise TimeoutError("The browser session is busy, try again.")
        if "err" in box:
            raise box["err"]
        return box.get("res")

    def run(self):
        while True:
            fn, args, box = self.q.get()
            try:
                box["res"] = fn(*args)
            except Exception as e:  # surfaced to the API caller
                box["err"] = e
                self.log(f"Error: {str(e).splitlines()[0][:160]}")
            box["ev"].set()

    # ---- helpers (worker thread only) -----------------------------------
    def _fill_form(self):
        pg = self.page
        user, pw = os.environ.get("NHAI_USER"), os.environ.get("NHAI_PASS")
        if not user or not pw:
            raise RuntimeError("Enter the portal user name and password first.")
        pg.select_option("#ddlUserType", ROLE_VALUE)
        pg.click("#userName")  # field is readonly until focused
        pg.fill("#userName", user)
        pg.click("#password")
        pg.fill("#password", pw)

    def _grab_captcha(self):
        self.page.wait_for_selector("#captchaCanvas", timeout=15000)
        time.sleep(0.4)  # let the canvas paint
        self.captcha = self.page.locator("#captchaCanvas").screenshot()
        self.captcha_version += 1

    def _collect_nav(self):
        links = self.page.eval_on_selector_all(
            "a", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,70), e.href])")
        seen, nav = set(), []
        for text, href in links:
            if not text or not allowed(href) or "?upc=" in href or "DownloadbyName" in href:
                continue
            if href.endswith("#") or href in seen:
                continue
            seen.add(href)
            nav.append({"title": text, "url": href})
        self.nav = nav[:150]

    def _disconnect(self):
        for closer in (lambda: self.ctx and self.ctx.close(), lambda: self.browser and self.browser.close(),
                       lambda: self.pw and self.pw.stop()):
            try:
                closer()
            except Exception:
                pass
        self.pw = self.browser = self.ctx = self.page = None
        self.captcha = None
        self.nav, self.scan = [], None

    # ---- commands (run on worker thread) --------------------------------
    def connect(self):
        self._disconnect()
        self.state, self.message = "starting", ""
        self.log("Opening NHAI Data Lake...")
        try:
            self.pw = sync_playwright().start()
            self.browser = self.pw.chromium.launch(headless=True)
            self.ctx = self.browser.new_context(accept_downloads=True, viewport={"width": 1400, "height": 900})
            self.page = self.ctx.new_page()
            self.page.on("dialog", self._on_dialog)
            self.page.goto(ROOT, wait_until="networkidle", timeout=60000)
            self._fill_form()
            self.log("Role, username and password filled. Captcha ready.")
            self._grab_captcha()
            self.state = "captcha_required"
        except Exception as e:
            self.state, self.message = "error", str(e).splitlines()[0][:200]
            self._disconnect()
            raise

    def _on_dialog(self, d):
        self.last_dialog = d.message
        self.log(f"Site says: {d.message[:120]}")
        try:
            d.accept()
        except Exception:
            pass  # page already navigated/closed; nothing left to dismiss

    def refresh_captcha(self):
        self.page.evaluate("createCaptcha()")
        self._grab_captcha()

    def submit_captcha(self, text):
        try:
            self._submit_captcha(text)
        except Exception as e:
            # never leave the page spinning on "Signing in": go back to a state the user can act on
            self.log(f"Sign-in step failed: {str(e).splitlines()[0][:120]}")
            try:
                if self.page and self.page.query_selector("#loginBtn") is None:
                    self.state = "logged_in"  # the portal did let us in; the failure came afterwards
                    self.ctx.storage_state(path=str(self.STATE_FILE))
                    self._collect_nav()
                    return
            except Exception:
                pass
            self.state, self.message = "error", "Sign-in did not finish. Click Connect to get a new captcha."
            self._disconnect()

    def _submit_captcha(self, text):
        pg = self.page
        self.state, self.message, self.last_dialog = "logging_in", "", ""
        self.log("Submitting captcha...")
        # The site's SIGN IN handler encrypts #password in place, so a retry would
        # double-encrypt it. Always re-enter the credentials before submitting.
        self._fill_form()
        pg.click("#txtInput")
        pg.fill("#txtInput", text)
        pg.click("#loginBtn")
        toast = ""
        try:
            # success = login button disappears; failure = an error toast appears
            pg.wait_for_function(
                "() => !document.querySelector('#loginBtn') || document.querySelector('.toast-message')", timeout=25000)
        except PWTimeout:
            pass
        if pg.query_selector("#loginBtn"):
            el = pg.query_selector(".toast-message")
            toast = (el.inner_text().strip() if el else "")
        else:
            try:
                pg.wait_for_load_state("networkidle", timeout=30000)
            except PWTimeout:
                pass
        if pg.query_selector("#loginBtn") is None:
            self.state = "logged_in"
            self.log(f"Logged in: {pg.title()[:60]}")
            self.ctx.storage_state(path=str(self.STATE_FILE))
            self._collect_nav()
            self.log(f"Found {len(self.nav)} portal sections.")
            return
        # still on the login page: wrong captcha or credentials
        self.message = toast or self.last_dialog or "Login was not accepted. Check the captcha and try again."
        self.log(f"Portal said: {self.message[:100]}")
        pg.evaluate("createCaptcha()")
        self._grab_captcha()
        self.state = "captcha_required"

    def open_page(self, url):
        if not allowed(url):
            raise ValueError("Only pages on the NHAI Data Lake can be opened.")
        pg = self.page
        self.log(f"Opening {urlparse(url).path[:70]}")
        pg.goto(url, wait_until="domcontentloaded", timeout=60000)
        try:
            pg.wait_for_load_state("networkidle", timeout=20000)
        except PWTimeout:
            pass
        if pg.query_selector("#loginBtn"):
            self.state, self.message = "error", "Session expired. Connect again."
            raise RuntimeError("Session expired. Connect again.")
        raw = pg.eval_on_selector_all(
            "a[href]", "els => els.map(e => [e.innerText.trim().replace(/\\s+/g,' ').slice(0,120), e.href])")
        seen, links = set(), []
        for text, href in raw:
            if allowed(href) and DOC_PATTERN.search(href) and href not in seen:
                seen.add(href)
                links.append({"title": text or unquote(href.rsplit("/", 1)[-1])[:80], "url": href})
        tables = pg.eval_on_selector_all("table", "els => els.length")
        self.scan = {"url": pg.url, "title": pg.title()[:120], "links": links[:400], "tables": tables}
        self.log(f"Scan done: {len(links)} downloadable links, {tables} tables.")
        return self.scan

    def download(self, url, dest_dir):
        if not allowed(url):
            raise ValueError("Only NHAI Data Lake links can be downloaded.")
        r = self.ctx.request.get(url, timeout=90000)
        if not r.ok:
            raise RuntimeError(f"Server returned {r.status}")
        ctype = r.headers.get("content-type", "")
        body = r.body()
        if "text/html" in ctype and b"loginBtn" in body[:200000]:
            raise RuntimeError("Session expired. Connect again.")
        name = None
        cd = r.headers.get("content-disposition", "")
        m = re.search(r"filename\*?=(?:UTF-8'')?\"?([^\";]+)", cd, re.I)
        if m:
            name = unquote(m.group(1))
        if not name:
            q = parse_qs(urlparse(url).query)
            name = (q.get("FileName") or q.get("filename") or [urlparse(url).path.rsplit("/", 1)[-1]])[0]
        name = re.sub(r"[^\w.\- ()]", "_", name).strip(" .") or "download"
        if "." not in name:
            name += ".pdf" if "pdf" in ctype else ""
        dest = dest_dir / name
        dest.write_bytes(body)
        self.log(f"Fetched {dest.name} ({len(body) // 1024} KB)")
        return dest

    def inspect(self, url=None):
        pg = self.page
        if url:
            self.open_page(url)
        return pg.evaluate("""() => {
          const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
          const d = e => ({tag:e.tagName, id:e.id, name:e.name||'', type:e.type||'', cls:(e.className||'').toString().slice(0,60),
                           text:(e.innerText||e.value||'').trim().replace(/\\s+/g,' ').slice(0,60), onclick:(e.getAttribute('onclick')||'').slice(0,80)});
          return {
            url: location.href, title: document.title,
            buttons: [...document.querySelectorAll('button,input[type=button],input[type=submit],a.btn,a[onclick]')].filter(vis).map(d),
            inputs: [...document.querySelectorAll('input:not([type=button]):not([type=submit]),textarea')].filter(vis).map(d),
            selects: [...document.querySelectorAll('select')].map(s => ({id:s.id, name:s.name, visible:vis(s),
                       options:[...s.options].slice(0,40).map(o => o.value+'='+o.text.trim())})),
            tables: [...document.querySelectorAll('table')].map(t => ({id:t.id, rows:t.rows.length,
                       head:[...(t.tHead?t.tHead.rows[0].cells:[])].map(c => c.innerText.trim()).slice(0,20)})),
            tabs: [...document.querySelectorAll('[role=tab],.nav-tabs a,.nav-link')].filter(vis).map(d)
          };
        }""")

    # ---- report modules (BG Release, FMC Cases, FMC Cases Report) --------
    MODULES = {
        "bg": {"label": "BG Release & Securities", "tag": "BG", "path": "BGReleaseSecurities/BGReleaseSecurities",
               "grid": "#tbl_bg_release_securities_grid", "get": "button[onclick='onclickBtsubmit()']",
               "report_path": "BGReleaseSecurities/BGReleaseSecuritiesReport",
               "report_grid": "#tbl_ro_wise_bg_release_securities_grid_report"},
        "fmc": {"label": "FMC Cases", "tag": "FMC", "path": "FMCCases/FMCCases",
                "grid": "#tbl_fmc_master_grid", "get": "button[onclick='onclickBtsubmit()']",
                "report_path": "FMCCases/FMCCasesReport",
                "report_grid": "#tbl_ro_wise_fmc_grid_report"},
        "notice": {"label": "Notice on Poor Keepup of Toilets, Admin Building", "tag": "NOTICE",
                   "path": "IssueNoticeToUserFeeAgency/IssueNoticeToUserFeeAgency",
                   "grid": "#tbl_toll_bidding_master", "get": None,
                   "report_path": "IssueNoticeToUserFeeAgency/IssueNoticeToUserFeeAgencyreport",
                   "report_grid": "#tbl_toll_bidding_master"},
    }

    def _mod(self, key):
        if key not in self.MODULES:
            raise ValueError("Unknown module.")
        return self.MODULES[key]

    REPORT_GET = "button[onclick='onclickBtnmultiplesubmit()']"

    def _mod_goto(self, key, report=False):
        m, pg = self._mod(key), self.page
        url = ROOT + (m["report_path"] if report else m["path"])
        # exact path match: ".../BGReleaseSecurities" is a prefix of ".../BGReleaseSecuritiesReport"
        if urlparse(pg.url).path.rstrip("/").lower() != urlparse(url).path.rstrip("/").lower():
            pg.goto(url, wait_until="domcontentloaded", timeout=60000)
            try:
                pg.wait_for_load_state("networkidle", timeout=25000)
            except PWTimeout:
                pass
        if pg.query_selector("#loginBtn"):
            self.state, self.message = "error", "Session expired. Connect again."
            raise RuntimeError("Session expired. Connect again.")
        pg.wait_for_selector(m["report_grid"] if report else m["grid"], state="attached", timeout=30000)
        return m

    def mod_filters(self, key):
        m = self._mod_goto(key)
        grab = lambda sel: self.page.eval_on_selector_all(
            sel + " option", "os => os.map(o => ({value:o.value, text:o.text.trim()})).filter(o => o.value)")
        out = {"agencies": grab("#agencylogin_id")}
        out["info"] = self.page.evaluate(
            "s => { const i=document.querySelector(s+'_info'); return i?i.innerText:'' }", m["grid"])
        return out

    def _save_download(self, dl, tag, dest_dir):
        suggested = re.sub(r"[^\w.\- ()]", "_", dl.suggested_filename or f"{tag}.xlsx")
        dest = dest_dir / f"{time.strftime('%Y%m%d_%H%M%S')}_{tag}_{suggested}"
        dl.save_as(str(dest))
        self.log(f"Fetched {dest.name} ({dest.stat().st_size // 1024} KB)")
        return dest

    def _apply_agency(self, agency, get_btn):
        pg = self.page
        self.log("Applying agency filter...")
        pg.select_option("#agencylogin_id", agency)
        pg.click(get_btn)
        try:
            pg.wait_for_load_state("networkidle", timeout=25000)
        except PWTimeout:
            pass
        time.sleep(1.5)

    def _export_grid(self, grid, fmt, tag, label, dest_dir):
        self.log(f"Exporting {fmt.upper()} from {label}...")
        with self.page.expect_download(timeout=60000) as info:
            # some grids are hidden drill-downs, so trigger the button directly
            self.page.locator(f"{grid}_wrapper button.buttons-{fmt}").first.dispatch_event("click")
        return self._save_download(info.value, tag, dest_dir)

    def mod_export(self, key, kind, agency, dest_dir):
        """kind: excel | pdf (the table) or report (the portal's report page: RO-wise summary as Excel + PDF).
        Returns the paths of the files written to dest_dir."""
        if kind == "report":
            m = self._mod_goto(key, report=True)
            if agency:
                self._apply_agency(agency, self.REPORT_GET)
            tag = m["tag"] + "Report"
            return [self._export_grid(m["report_grid"], fmt, tag, m["label"] + " report", dest_dir)
                    for fmt in ("excel", "pdf")]
        if kind in ("excel", "pdf"):
            m = self._mod_goto(key)
            if agency:
                self._apply_agency(agency, m["get"])
            return [self._export_grid(m["grid"], kind, m["tag"], m["label"], dest_dir)]
        raise ValueError("Unknown export type.")

    ONGOING_PATH = "OnboardingUserFeeAgencyV2/OngoingContracts"
    ONGOING_GRID = "#tbl_toll_bidding_master_ongoing"

    def ongoing_export(self, dest_dir):
        """Onboarding-User Fee Agency > FILTER > Ongoing Contracts: Excel dump of every ongoing contract.
        Returns (path of the file written to dest_dir, number of entries the portal shows)."""
        pg = self.page
        self.log("Opening Onboarding-User Fee Agency > Ongoing Contracts...")
        pg.goto(ROOT + self.ONGOING_PATH, wait_until="domcontentloaded", timeout=90000)
        if pg.query_selector("#loginBtn"):
            self.state, self.message = "error", "Session expired. Connect again."
            raise RuntimeError("Session expired. Connect again.")
        info = self.ONGOING_GRID + "_info"
        # the grid is filled client-side from one large request; export only once it reports its entries
        pg.wait_for_function(
            "s => { const e = document.querySelector(s); return e && /of [\\d,]+ entries/.test(e.innerText) "
            "&& !/Showing 0 to 0/.test(e.innerText) }", arg=info, timeout=120000)
        text = pg.inner_text(info)
        entries = int(re.search(r"of ([\d,]+) entries", text).group(1).replace(",", ""))
        if re.search(r"filtered from", text):
            raise RuntimeError("The Ongoing Contracts table is filtered; clear the filter on the portal first.")
        saved = self._export_grid(self.ONGOING_GRID, "excel", "ONGOING", "Ongoing Contracts", dest_dir)
        return saved, entries

    # ---- session persistence -------------------------------------------
    STATE_FILE = BASE / "auth_state.json"

    def resume(self):
        """Re-use a saved session (survives server restarts) without a new captcha."""
        if not self.STATE_FILE.exists() or self.state != "idle":
            return False
        try:
            self.pw = sync_playwright().start()
            self.browser = self.pw.chromium.launch(headless=True)
            self.ctx = self.browser.new_context(accept_downloads=True, viewport={"width": 1400, "height": 900},
                                                storage_state=str(self.STATE_FILE))
            self.page = self.ctx.new_page()
            self.page.on("dialog", self._on_dialog)
            self.page.goto(ROOT + "CentralDashboard/CentralAlertDashboard", wait_until="domcontentloaded", timeout=60000)
            try:
                self.page.wait_for_load_state("networkidle", timeout=20000)
            except PWTimeout:
                pass
            if self.page.query_selector("#loginBtn") or "login" in self.page.url.lower():
                self._disconnect()
                return False
            self.state = "logged_in"
            self._collect_nav()
            self.log("Resumed saved session (no captcha needed).")
            return True
        except Exception:
            self._disconnect()
            return False

    def shot(self):
        return self.page.screenshot(type="jpeg", quality=70) if self.page else None

    def logout(self):
        self._disconnect()
        self.state, self.message = "idle", ""
        self.log("Disconnected.")


W = Worker()
W.start()
threading.Thread(target=lambda: W.call(W.resume, timeout=120), daemon=True).start()


# ---- files for the browser ------------------------------------------------------
# A task works in a temporary folder that is deleted when it ends. Its files are handed to the browser session that
# ran it, which downloads them to its own computer. A copy is held in memory only (never on disk) for KEEP_MINUTES, so
# a blocked or cancelled download can be repeated from the Downloads page; then it is dropped.
KEEP_MINUTES = 15
_outbox, _outbox_lock = {}, threading.Lock()


@contextmanager
def work_dir():
    d = Path(tempfile.mkdtemp(prefix="tod_job_"))
    try:
        yield d
    finally:
        shutil.rmtree(d, ignore_errors=True)


def _prune():
    cut = time.time() - KEEP_MINUTES * 60
    for k in [k for k, f in _outbox.items() if f["t"] < cut]:
        del _outbox[k]


def _sid():
    return session.get("sid", "")


def _public(fid, f):
    return {"id": fid, "name": f["name"], "size": len(f["data"]),
            "time": time.strftime("%d %b %H:%M", time.localtime(f["t"])),
            "until": time.strftime("%H:%M", time.localtime(f["t"] + KEEP_MINUTES * 60))}


def hand_over(items):
    """items: paths, or (name, bytes) pairs. Holds them for the current browser session; returns what the page needs
    to download them."""
    sid, out = _sid(), []
    with _outbox_lock:
        _prune()
        for it in items:
            name, data = (it.name, it.read_bytes()) if isinstance(it, Path) else it
            fid = secrets.token_urlsafe(16)
            _outbox[fid] = {"name": name, "data": data, "t": time.time(), "sid": sid}
            out.append(_public(fid, _outbox[fid]))
    return out


def files_list():
    """This browser session's files that can still be downloaded again, newest first."""
    sid = _sid()
    with _outbox_lock:
        _prune()
        mine = [(k, f) for k, f in _outbox.items() if f["sid"] == sid]
    return [_public(k, f) for k, f in sorted(mine, key=lambda kf: kf[1]["t"], reverse=True)]


def zip_of(paths, name):
    """One zip of the given files (in memory), each named as the portal named it."""
    buf, used = io.BytesIO(), set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            arc = re.sub(r"^\d{8}_\d{6}_[A-Za-z]+_", "", p.name)
            arc = p.name if arc in used else arc
            used.add(arc)
            z.write(p, arc)
    return name, buf.getvalue()


def ok(**kw):
    return jsonify(ok=True, **kw)


def fail(msg, code=400):
    return jsonify(ok=False, error=str(msg)), code


def signed_in():
    return session.get("user") == DASH_USER


@app.before_request
def require_sign_in():
    if signed_in():
        session.setdefault("sid", secrets.token_urlsafe(16))  # whose downloads are whose
        return None
    if request.endpoint in SIGNED_OUT_OK:
        return None
    if request.path.startswith(("/api/", "/files/")):
        return fail("Your dashboard session has ended. Sign in again.", 401)
    return redirect("/login")


@app.get("/login")
def login_page():
    if signed_in():
        return redirect("/")
    r = send_from_directory(app.static_folder, "login.html")
    r.headers["Cache-Control"] = "no-store"
    return r


@app.post("/api/login")
def login():
    body = request.get_json(silent=True) or {}
    user, pw = str(body.get("user") or "").strip(), str(body.get("password") or "")
    who = request.remote_addr or "?"
    with _fails_lock:
        n, until = _fails.get(who, (0, 0.0))
        if until > time.time():
            return jsonify(ok=False, error="Too many attempts. Try again shortly.", retry=int(until - time.time()) + 1), 429
    if not user or not pw:
        return fail("Enter your user name and password.")
    good = hmac.compare_digest(user.lower().encode(), DASH_USER.lower().encode()) & check_password_hash(DASH_HASH, pw)
    if not good:
        time.sleep(0.4)  # slow down guessing
        with _fails_lock:
            n += 1
            _fails[who] = (0, time.time() + LOCK_SECONDS) if n >= MAX_FAILS else (n, 0.0)
            if n >= MAX_FAILS:
                return jsonify(ok=False, error="Too many attempts. Try again shortly.", retry=LOCK_SECONDS), 429
        return fail(f"That user name and password don't match. {MAX_FAILS - n} attempt{'s' if MAX_FAILS - n != 1 else ''} "
                    "left before a short pause.", 401)
    with _fails_lock:
        _fails.pop(who, None)
    session.clear()
    session["user"] = DASH_USER
    session["sid"] = secrets.token_urlsafe(16)
    session.permanent = bool(body.get("remember"))
    return ok(user=DASH_USER)


@app.post("/api/logout")
def logout():
    session.clear()
    return ok()


@app.get("/")
def index():
    r = send_from_directory(app.static_folder, "index.html")
    r.headers["Cache-Control"] = "no-store"  # always serve the current page after an update
    return r


@app.get("/api/status")
def status():
    return jsonify(state=W.state, message=W.message, logs=W.logs[-40:], nav=W.nav, scan=W.scan,
                   captcha_version=W.captcha_version, files=files_list(),
                   has_credentials=bool(os.environ.get("NHAI_USER") and os.environ.get("NHAI_PASS")),
                   user=os.environ.get("NHAI_USER", ""), account=session.get("user", ""))


def _read_json(name):
    import json
    f = BASE / name
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
    except ValueError:
        return None


def _save_latest(name, key, payload):
    """Remember the last result of a job so the dashboard cards survive a restart. Never blocks a job."""
    import json
    try:
        data = _read_json(name) or {}
        data[key] = {**payload, "date": time.strftime("%Y-%m-%d"), "saved_at": time.strftime("%d %b %H:%M")}
        (BASE / name).write_text(json.dumps(data, indent=1, default=str), encoding="utf-8")
    except Exception as e:
        W.log(f"Could not remember the last {key} result: {str(e)[:80]}")


@app.get("/api/history")
def history():
    """Daily snapshots written by the KPI report (report_history.json), for the dashboard trend charts."""
    data = _read_json("report_history.json") or {}
    return ok(history=[{"date": d, **v} for d, v in sorted(data.items())])


@app.get("/api/dashboard")
def dashboard():
    """Everything the KPI cards need in one request: last report summary, last CO results, daily snapshots."""
    hist = _read_json("report_history.json") or {}
    co = _read_json("co_latest.json") or {}
    return ok(kpi=(_read_json("kpi_latest.json") or {}).get("report"), notification=co.get("notification"),
              windfall=co.get("windfall"), exports=co.get("exports"),
              history=[{"date": d, **v} for d, v in sorted(hist.items())], today=time.strftime("%Y-%m-%d"))


@app.post("/api/credentials")
def save_credentials():
    """First-run setup: keep the portal user name and password in .env on this computer only."""
    origin = request.headers.get("Origin")
    if request.host not in ("127.0.0.1:5055", "localhost:5055") or (origin and urlparse(origin).netloc != request.host):
        return fail("Open the tool from http://127.0.0.1:5055 to save credentials.", 403)
    body = request.get_json(silent=True) or {}
    user, pw = str(body.get("user") or "").strip(), str(body.get("password") or "")
    if not user or not pw:
        return fail("Enter both the portal user name and the password.")
    if "\n" in user + pw or "\r" in user + pw:
        return fail("User name and password cannot contain line breaks.")
    ENV_FILE.touch(exist_ok=True)
    set_key(str(ENV_FILE), "NHAI_USER", user, quote_mode="always")
    set_key(str(ENV_FILE), "NHAI_PASS", pw, quote_mode="always")
    os.environ["NHAI_USER"], os.environ["NHAI_PASS"] = user, pw
    return ok()


@app.post("/api/connect")
def connect():
    try:
        W.call(W.connect, timeout=180)
        return ok()
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


@app.get("/api/captcha.png")
def captcha_png():
    if not W.captcha:
        return Response(status=404)
    return Response(W.captcha, mimetype="image/png", headers={"Cache-Control": "no-store"})


@app.post("/api/captcha/refresh")
def captcha_refresh():
    try:
        W.call(W.refresh_captcha)
        return ok()
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


@app.post("/api/captcha")
def captcha_submit():
    text = (request.get_json(silent=True) or {}).get("text", "").strip()
    if not text:
        return fail("Enter the captcha text.")
    try:
        W.call(W.submit_captcha, text, timeout=120)
        return ok(state=W.state, message=W.message)
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


@app.post("/api/open")
def open_page():
    url = (request.get_json(silent=True) or {}).get("url", "")
    if W.state != "logged_in":
        return fail("Not connected.", 409)
    try:
        return ok(scan=W.call(W.open_page, url, timeout=120))
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


@app.post("/api/download")
def download():
    url = (request.get_json(silent=True) or {}).get("url", "")
    if W.state != "logged_in":
        return fail("Not connected.", 409)
    try:
        with work_dir() as d:
            return ok(file=hand_over([W.call(W.download, url, d, timeout=150)])[0])
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


@app.post("/api/mod/filters")
def mod_filters():
    key = (request.get_json(silent=True) or {}).get("module", "")
    if W.state != "logged_in":
        return fail("Not connected.", 409)
    try:
        return ok(**W.call(W.mod_filters, key, timeout=120))
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


_co_lock = threading.Lock()


@app.post("/api/co/notification")
def co_notification():
    """Toll Notification Power BI report: public link, no portal login needed."""
    if not _co_lock.acquire(blocking=False):
        return fail("A CO Division download is already running.", 409)
    try:
        import sys
        sys.path.insert(0, str(BASE))
        import pbi_toll
        with work_dir() as d:
            t, files = pbi_toll.run(d, log=W.log)
            files = hand_over(files)
        out = dict(summary=t["summary"]["rows"], gt85_total=t["gt85"]["total"], projects=len(t["projects"]["rows"]))
        _save_latest("co_latest.json", "notification", out)
        return ok(files=files, **out)
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)
    finally:
        _co_lock.release()


@app.post("/api/co/remittance")
def co_remittance():
    """Toll Plaza Windfall Power BI reports: public links, no portal login needed."""
    if not _co_lock.acquire(blocking=False):
        return fail("A CO Division download is already running.", 409)
    try:
        import sys
        sys.path.insert(0, str(BASE))
        import pbi_remit
        with work_dir() as d:
            t, entire, files = pbi_remit.run(d, log=W.log)
            files = hand_over(files)
        out = dict(period=t["period"], ro_total=t["ro"]["total"][1], plazas=len(t["plazas"]["rows"]),
                   daily=len(t["daily"]["rows"]), entire=len(entire["rows"]),
                   missing=entire["missing"], last=entire["last"].strftime("%d-%m-%Y") if entire["last"] else "")
        _save_latest("co_latest.json", "windfall", out)
        return ok(files=files, **out)
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)
    finally:
        _co_lock.release()


@app.post("/api/report/build")
def report_build():
    """Weekly CO Division KPI report: fresh FMC + BG exports from the portal -> filled Word template."""
    if W.state != "logged_in":
        return fail("Not connected. Sign in first; the report needs fresh FMC and BG data.", 409)
    if not _co_lock.acquire(blocking=False):
        return fail("Another download is already running.", 409)
    try:
        import sys
        sys.path.insert(0, str(BASE))
        import report_builder
        with work_dir() as d:
            got, dumps = {}, []
            for key, kind, tag in (("fmc", "excel", "fmc_reg"), ("fmc", "report", "fmc_cnt"),
                                   ("bg", "excel", "bg_reg"), ("bg", "report", "bg_cnt"),
                                   ("notice", "excel", "notice_reg"), ("notice", "report", "notice_cnt")):
                files = W.call(W.mod_export, key, kind, None, d, timeout=240)
                dumps += files
                xl = [f for f in files if f.suffix.lower() == ".xlsx"]
                if not xl:
                    return fail(f"The portal did not return an Excel file for {key} {kind}.", 500)
                got[tag] = xl[0]
            got["ongoing"], entries = W.call(W.ongoing_export, d, timeout=300)
            dumps.append(got["ongoing"])
            W.log("Building the KPI report...")
            s = report_builder.build(got["fmc_reg"], got["fmc_cnt"], got["bg_reg"], got["bg_cnt"], d / "report",
                                     log=W.log, notice_reg=got["notice_reg"], notice_cnt=got["notice_cnt"],
                                     ongoing=got["ongoing"], ongoing_entries=entries)
            W.log(f"Saved {s['file'].name}")
            # the report, and every portal export it was built from (the data behind each dashboard card)
            files = hand_over([s["file"], zip_of(dumps, f"Source data {s['date']:%d%m%Y}.zip")])
        pick = ("total", "piu", "agency", "ro", "hq", "closed", "pendency", "piu_share", "agency_pendency", "ro_pendency")
        out = dict(fmc={k: s["fmc"][k] for k in pick}, bg={k: s["bg"][k] for k in pick},
                   notice=s["notice"],
                   ongoing={k: s["ongoing"][k] for k in ("total", "regular", "eq", "manpower", "top10_total")}
                   | {"agencies": len(s["ongoing"]["agencies"]), "top10": s["ongoing"]["top10"]}
                   | s["ongoing"]["cleaning"],
                   kpi=[[d, v] for d, v in s["kpi"]], issues=s["issues"])
        _save_latest("kpi_latest.json", "report", out)
        return ok(files=files, **out)
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)
    finally:
        _co_lock.release()


@app.post("/api/mod/export")
def mod_export():
    body = request.get_json(silent=True) or {}
    if W.state != "logged_in":
        return fail("Not connected.", 409)
    if body.get("kind") not in ("excel", "pdf", "report"):
        return fail("Unknown export type.")
    agency = str(body.get("agency") or "")
    if agency and not agency.isdigit():
        return fail("Invalid filter.")
    try:
        with work_dir() as d:
            files = hand_over(W.call(W.mod_export, body.get("module", ""), body["kind"], agency or None, d,
                                     timeout=240))
        _save_latest("co_latest.json", "exports", {"module": body.get("module", ""), "kind": body["kind"]})
        return ok(files=files)
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


@app.post("/api/inspect")
def inspect_page():
    url = (request.get_json(silent=True) or {}).get("url")
    if W.state != "logged_in":
        return fail("Not connected.", 409)
    try:
        return ok(page=W.call(W.inspect, url, timeout=120))
    except Exception as e:
        return fail(str(e).splitlines()[0], 500)


@app.get("/api/screenshot.jpg")
def screenshot():
    try:
        img = W.call(W.shot, timeout=30)
    except Exception:
        img = None
    return Response(img, mimetype="image/jpeg", headers={"Cache-Control": "no-store"}) if img else Response(status=404)


@app.post("/api/disconnect")
def disconnect():
    W.call(W.logout)
    return ok()


@app.get("/files/<fid>")
def get_file(fid):
    with _outbox_lock:
        _prune()
        f = _outbox.get(fid)
    if not f or f["sid"] != _sid():
        return Response(f"This file is no longer held (files are kept for {KEEP_MINUTES} minutes after a task). "
                        "Run the task again.", status=404, mimetype="text/plain")
    r = send_file(io.BytesIO(f["data"]), as_attachment=True, download_name=f["name"])
    r.headers["Cache-Control"] = "no-store"
    return r


if __name__ == "__main__":
    print("Dashboard: http://127.0.0.1:5055")
    app.run(host="127.0.0.1", port=5055, threaded=True)
