"""Where the tool keeps its files.

Run from source: the project folder. Packaged as .exe: the folder that holds the .exe, so the template, history,
settings, saved sign-in and downloads sit next to it where the user can see and edit them.

Hosted (DATA_DIR set, e.g. a Railway volume): the program stays in APP, but everything the tool writes while it runs
(daily history, saved portal sign-in, last results, .env) lives in DATA so it survives a redeploy."""
import os
import shutil
import sys
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)
APP = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
DATA = Path(os.environ["DATA_DIR"]).resolve() if os.environ.get("DATA_DIR") else APP

if DATA != APP:
    DATA.mkdir(parents=True, exist_ok=True)
    # first start on an empty disk: begin from the history that was deployed with the program
    if not (DATA / "report_history.json").exists() and (APP / "report_history.json").exists():
        shutil.copy2(APP / "report_history.json", DATA / "report_history.json")
