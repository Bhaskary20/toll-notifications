"""Entry point of the packaged tool (CO Division Report Tool.exe).

Starts the dashboard on this computer only, opens it in the default browser and keeps running until the window is
closed. Run from source it behaves the same way."""
import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import APP, FROZEN  # noqa: E402

PORT = 5055
URL = f"http://127.0.0.1:{PORT}"


def already_running():
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", PORT)) == 0


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if already_running():
        print("The tool is already running; opening it in the browser.")
        webbrowser.open(URL)
        return
    browsers = APP / "browsers"
    if browsers.exists():  # the browser that ships with the tool; no internet install needed
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)
    elif FROZEN:
        print("The 'browsers' folder next to this program is missing. Copy the whole folder, not just the .exe.")
        input("Press Enter to close.")
        return
    import app  # noqa: E402  (after the browser path is set)

    print("CO Division Report Tool")
    print(f"Dashboard: {URL}")
    print("Keep this window open while you work. Close it to stop the tool.")
    threading.Thread(target=lambda: (time.sleep(1.5), webbrowser.open(URL)), daemon=True).start()
    app.app.run(host="127.0.0.1", port=PORT, threaded=True)


if __name__ == "__main__":
    main()
