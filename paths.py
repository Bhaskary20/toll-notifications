"""Where the tool keeps its files.

Run from source: the project folder. Packaged as .exe: the folder that holds the .exe, so the template, history,
settings, saved sign-in and downloads sit next to it where the user can see and edit them."""
import sys
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)
APP = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
