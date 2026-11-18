#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 src/psd_to_png_gui.py 打成单文件无控制台 EXE。

按 PyInstaller 在受限环境的经验：永远不删旧产物，只挪到 %TEMP%，
避免触发批量删除保护；固定 --workpath，不加 --clean。
"""
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src" / "psd_to_png_gui.py"
APP_NAME = "PSD_PSB_to_PNG_JPG"
OUT_NAME = "PSD-PSB一键转PNG-JPG.exe"
PY = r"C:\Users\Administrator\.workbuddy-ai\binaries\python\envs\psdtool\Scripts\python.exe"

WORK = ROOT / "build" / "work"
OUT = ROOT / "build" / "out"
SPEC = ROOT / "build" / "spec"


def recycle(path: Path) -> None:
    """把旧产物挪走而不是删除。"""
    if not path.exists():
        return
    dst = Path(tempfile.gettempdir()) / "old_build" / f"{path.name}_{int(time.time())}"
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(path), str(dst))
    print(f"[recycle] {path} -> {dst}")


def main() -> int:
    for p in (WORK, OUT, SPEC):
        p.parent.mkdir(parents=True, exist_ok=True)
    recycle(WORK)
    recycle(OUT)
    recycle(SPEC)
    WORK.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    SPEC.mkdir(parents=True, exist_ok=True)

    cmd = [
        PY, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile", "--windowed",
        "--name", APP_NAME,
        "--workpath", str(WORK),
        "--distpath", str(OUT),
        "--specpath", str(SPEC),
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "windnd",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
        "--exclude-module", "PySide6",
        "--exclude-module", "PyQt5",
        "--exclude-module", "PIL",
        "--exclude-module", "numpy",
        "--exclude-module", "matplotlib",
        "--exclude-module", "scipy",
        "--exclude-module", "pandas",
        "--exclude-module", "pytest",
        str(SRC),
    ]
    print("[build] " + " ".join(cmd), flush=True)
    rc = subprocess.call(cmd, cwd=str(ROOT))
    if rc != 0:
        print(f"[build] PyInstaller failed: {rc}")
        return rc

    built = OUT / f"{APP_NAME}.exe"
    if not built.exists():
        print("[build] exe not found")
        return 1

    dist = ROOT / "dist"
    dist.mkdir(parents=True, exist_ok=True)
    target = dist / OUT_NAME
    # 覆盖写，不做删除
    with built.open("rb") as f, target.open("wb") as g:
        shutil.copyfileobj(f, g)
    print(f"[build] OK -> {target} ({target.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
