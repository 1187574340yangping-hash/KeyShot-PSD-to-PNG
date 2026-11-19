#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 src/psd_to_png_gui.py 打成 Windows 可执行程序。

两种模式：

  python build_exe.py               默认：onefile（仓库分发用，单文件）
  python build_exe.py --onedir      onedir（本机推荐）：固定目录运行，
                                    不往 %TEMP% 解压，避免火绒等安全软件
                                    锁住 DLL 导致退出时弹
                                    "Failed to remove temporary directory"

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
ICON = ROOT / "assets" / "app_icon.ico"
ASSETS = ROOT / "assets"
APP_NAME = "PSD一键转PNG"
ASCII_NAME = "PSDTool"  # onedir 专用：中文名会让 PKG 路径解析失败
OUT_NAME = "PSD-PSB一键转PNG-JPG.exe"
PY = r"C:\Users\Administrator\.workbuddy-ai\binaries\python\envs\psdtool\Scripts\python.exe"

WORK = ROOT / "build" / "work"
OUT = ROOT / "build" / "out"
SPEC = ROOT / "build" / "spec"

# onedir 安装目录：固定在桌面文件夹，避开 AppData\Local 下 EXE 被安全软件拦截执行的问题
INSTALL_DIR = Path(r"C:\Users\Administrator\Desktop\PSD一键转PNG工具")


def recycle(path: Path) -> None:
    """把旧产物挪走而不是删除。"""
    if not path.exists():
        return
    dst = Path(tempfile.gettempdir()) / "old_build" / f"{path.name}_{int(time.time())}"
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(path), str(dst))
    print(f"[recycle] {path} -> {dst}")


def pyinstaller_args(mode: str, work: Path, out: Path, spec: Path) -> list:
    cmd = [
        PY, "-m", "PyInstaller",
        "--noconfirm",
        "--windowed",
        "--icon", str(ICON),
        "--add-data", f"{ASSETS};assets",
        "--name", APP_NAME,
        "--workpath", str(work),
        "--distpath", str(out),
        "--specpath", str(spec),
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--collect-all", "tkinterdnd2",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
        "--exclude-module", "PySide6",
        "--exclude-module", "PyQt5",
        "--collect-all", "PIL",
        "--collect-all", "customtkinter",
        "--exclude-module", "numpy",
        "--exclude-module", "matplotlib",
        "--exclude-module", "scipy",
        "--exclude-module", "pandas",
        "--exclude-module", "pytest",
    ]
    # onefile：固定运行时临时目录，避免每次在系统 %TEMP% 生成随机 _MEI 目录。
    # 随机目录会被安全软件锁定导致退出时弹 "Failed to remove temporary directory"；
    # 固定到程序自己的目录后，删除失败也不会污染系统临时目录，且可加入安全软件白名单。
    if mode == "onefile":
        cmd += ["--runtime-tmpdir", r"C:\Users\Administrator\PSDToolRuntime"]

    cmd.append("--onedir" if mode == "onedir" else "--onefile")
    # onedir 用 ASCII 程序名：中文名在部分环境下会让 bootloader 无法定位 PKG
    if mode == "onedir":
        cmd[cmd.index("--name") + 1] = ASCII_NAME
    cmd.append(str(SRC))
    return cmd


def build(mode: str) -> int:
    for p in (WORK, OUT, SPEC):
        p.parent.mkdir(parents=True, exist_ok=True)
    recycle(WORK)
    recycle(OUT)
    recycle(SPEC)
    WORK.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    SPEC.mkdir(parents=True, exist_ok=True)

    cmd = pyinstaller_args(mode, WORK, OUT, SPEC)
    print("[build] " + " ".join(cmd), flush=True)
    rc = subprocess.call(cmd, cwd=str(ROOT))
    if rc != 0:
        print(f"[build] PyInstaller failed: {rc}")
        return rc

    if mode == "onedir":
        bundle = OUT / ASCII_NAME
        exe = bundle / f"{ASCII_NAME}.exe"
        if not exe.exists():
            print("[build] onedir exe not found")
            return 1
        # 固定目录：旧目录挪走，新目录整体搬过去
        recycle(INSTALL_DIR)
        INSTALL_DIR.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(bundle), str(INSTALL_DIR))
        print(f"[build] onedir installed -> {INSTALL_DIR} ({exe.name})")
        print(f"[build] files: {sum(1 for _ in INSTALL_DIR.rglob('*'))}")
        return 0

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
    mode = "onedir" if "--onedir" in sys.argv else "onefile"
    print(f"[build] mode={mode}")
    sys.exit(build(mode))
