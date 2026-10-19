# Running the tool on a Windows server inside the NHAI network

The tool must run on a machine the NHAI Data Lake answers; a server on the NHAI network does. People then open it at
a subdomain such as `https://toll.<your-domain>` and use it like the local version.

```
People's browsers --HTTPS--> Caddy (or IIS) on the server --> serve.py on 127.0.0.1:5055 --> headless browser --> Data Lake
```

`serve.py` listens on this machine only. Only the web server in front of it (port 443) is reachable from the network.

## 1. Ask NHAI IT for three things
1. A DNS record for the subdomain pointing to the server's address.
2. A TLS certificate and key for that name (from the organisation's CA).
3. Port 443 open to the networks the users are on. Do not open 5055.

## 2. Install on the server
1. Install Python 3.12 or newer, for all users (tick "Add python.exe to PATH").
2. Put the project in `C:\TollDesk` (`git clone https://github.com/Bhaskary20/toll-notifications.git C:\TollDesk`).
3. In `C:\TollDesk` run:
   ```
   py -m pip install -r requirements.txt
   set PLAYWRIGHT_BROWSERS_PATH=C:\TollDesk\browsers
   py -m playwright install chromium-headless-shell
   ```
   The browser goes into `C:\TollDesk\browsers`, where `serve.py` looks for it, so any account can run the tool.
4. Make the dashboard password hash: `py tools/make_password_hash.py`.
5. Create `C:\TollDesk\.env` (use single quotes around the hash, it contains `$` signs):
   ```
   HOSTED=1
   DASHBOARD_USER=Admin
   DASHBOARD_PASS_HASH='scrypt:...the line from step 4...'
   PUBLIC_HOST=toll.<your-domain>
   NHAI_USER=<Data Lake user name>
   NHAI_PASS=<Data Lake password>
   ```
   Add `SKIP_WORD_PAGES=1` if Word is not installed on the server or the tool runs as a Windows service (see 4).
   Keep this file readable only by the administrators and the account that runs the tool. Never commit it.

## 3. Try it
```
cd C:\TollDesk
py serve.py
```
On the server open `http://127.0.0.1:5055`, sign in, click Connect and check that the captcha appears. Then open
`http://127.0.0.1:5055/api/diag`: every step should say `"ok": true`.

## 4. Keep it running
Pick one:

**A. Windows service with NSSM** (starts at boot, restarts on a crash). Download NSSM, then:
```
nssm install TollDesk "C:\Path\To\Python\python.exe" "C:\TollDesk\serve.py"
nssm set TollDesk AppDirectory C:\TollDesk
nssm set TollDesk Start SERVICE_AUTO_START
nssm start TollDesk
```
A service has no desktop, so Microsoft Word cannot lay out the report for the Contents page numbers. Keep
`SKIP_WORD_PAGES=1` in `.env`; the report then opens with blank page numbers and Word fills them in on "Yes".

**B. Task Scheduler** (needed only if you want Word to fill the Contents page numbers on the server). Create a task that runs
`py C:\TollDesk\serve.py` "At log on" for a dedicated account that stays logged in, with "Run only when user is logged
on". Word must be installed and activated for that account. Leave `SKIP_WORD_PAGES` out.

## 5. Put HTTPS in front (Caddy, simplest)
Download Caddy for Windows and make `C:\Caddy\Caddyfile`:
```
toll.<your-domain> {
    tls C:\certs\toll.crt C:\certs\toll.key
    reverse_proxy 127.0.0.1:5055 {
        transport http {
            response_header_timeout 15m
        }
    }
}
```
Run it as a service too (`caddy run --config C:\Caddy\Caddyfile`, or with NSSM). The long timeout matters: the weekly
report runs for several minutes in one request.

If IIS is used instead: URL Rewrite + Application Request Routing, forward `X-Forwarded-For`, `X-Forwarded-Proto` and
`X-Forwarded-Host`, and raise the ARR proxy time-out to 900 seconds. Caddy sends those headers by itself.

## 6. Check from a user's computer
- `https://toll.<your-domain>/healthz` shows `ok`.
- The login page opens; sign in; Connect shows the captcha.

## Notes
- One process serves everyone and they share the one signed-in portal session. Two reports at once queue up.
- Generated files go straight to the user's browser and are kept in memory for 15 minutes, never on the server's disk.
- Back up `report_history.json`, `.env` and `auth_state.json` (the saved portal sign-in).
- Update: `git pull`, then `nssm restart TollDesk`.
- `PUBLIC_HOST` lets the dashboard save the portal login from the subdomain. Without it the dashboard refuses, and the
  user name and password must come from `.env`.
