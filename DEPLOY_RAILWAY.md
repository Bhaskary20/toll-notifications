# Running the tool on Railway

The tool is a web server that drives a headless Chromium browser. Railway runs it from the `Dockerfile` in this
folder. Nothing here changes how it runs on a computer (`py app.py` or the .exe still work as before).

## Before you start
- Check that the NHAI portal opens from a cloud server. Railway has no India region, and some government sites refuse
  foreign addresses. Step 6 is the real test: if the portal does not load or the captcha never appears there, hosting
  will not work and you can delete the service.
- Use a private GitHub repo or be sure no secrets are committed (`.env`, `auth_state.json` are already ignored).

## Steps
1. Make a dashboard password hash on your computer: `py tools/make_password_hash.py` and copy the
   `DASHBOARD_PASS_HASH=...` line it prints.
2. railway.com > New Project > Deploy from GitHub repo > `Bhaskary20/toll-notifications`. Railway finds the
   `Dockerfile` and `railway.json` by itself.
3. Service > Variables, add:

   | Variable | Value |
   |---|---|
   | `DASHBOARD_PASS_HASH` | the hash from step 1 (required: the server refuses to start without it) |
   | `DASHBOARD_USER` | the name people sign in with (default `Admin`) |
   | `NHAI_USER`, `NHAI_PASS` | the Data Lake login (optional; otherwise type them once in the dashboard) |
   | `PBI_NOTIFICATION_URL` | contents of `pbi_notification_url.txt` (only for the Toll Notification card) |
   | `PBI_REMITTANCE_URL`, `PBI_REMITTANCE_DATA_URL` | contents of the two remittance files (only for the Windfall card) |

4. Service > Settings > Volumes > New Volume, mount path `/data`. Without it the daily history and the saved portal
   sign-in are lost on every restart and redeploy.
5. Service > Settings > Networking > Generate Domain. Open the link and sign in.
6. Click Connect, type the captcha, then run the weekly report.

## Things to know
- Files you generate download straight to your browser, as before. They are held in the server's memory for 15 minutes
  and never written to disk.
- Contents page numbers: on your computer Word fills them in. A server has no Word, so the report opens with the
  page numbers blank; press Yes when Word asks to update fields (or press F9 on the Contents page).
- The server runs one worker on purpose: all users share the one signed-in portal session. Two people running the
  report at once will queue behind each other.
- `report_history.json` on the volume starts as a copy of the one in the repo and is then updated by every report.
  Later edits to the file in git do not overwrite the volume's copy.
- Cost: one always-on service plus Chromium needs roughly 1 GB of memory.
