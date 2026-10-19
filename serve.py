"""Runs the dashboard as an always-on server on a Windows machine (see DEPLOY_WINDOWS_SERVER.md).

Unlike launcher.py it opens no browser window. It listens on this machine only (127.0.0.1) so that the web server in
front of it (Caddy / IIS) is the one thing reachable from the network, and it serves requests with waitress, which
runs on Windows. One process on purpose: every request shares the one signed-in portal browser."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import APP  # noqa: E402

HOST = os.environ.get("LISTEN_HOST", "127.0.0.1")
PORT = int(os.environ.get("LISTEN_PORT", "5055"))


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    browsers = APP / "browsers"   # the browser installed next to the program, shared by whichever account runs it
    if browsers.exists():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)
    from waitress import serve
    import app
    print(f"Toll Operations Desk on http://{HOST}:{PORT}  (hosted mode: {app.HOSTED})", flush=True)
    serve(app.app, host=HOST, port=PORT, threads=16, channel_timeout=900, ident="tod")


if __name__ == "__main__":
    main()
