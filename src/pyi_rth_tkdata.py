"""PyInstaller runtime hook for environments where Tcl detection is skipped."""
import os
import sys
from pathlib import Path

base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
os.environ["TCL_LIBRARY"] = str(base / "_tcl_data")
os.environ["TK_LIBRARY"] = str(base / "_tk_data")
