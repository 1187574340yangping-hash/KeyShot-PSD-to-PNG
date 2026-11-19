#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PSD / PSB 一键转 PNG · JPG  ·  增强版 v2.3

面向「PS 大画板 / 大画布」场景：画布很大、内容只占中间一小块、
四周大量透明或白色留白；文件常常是 1~2 GB 的 .psb。

相对原版（只认 .psd、不裁切、只出 PNG）新增：
  1. 支持 .psd 与 .psb（Photoshop 大型文档）
  2. 导出前自动裁切到内容边界（透明边 / 纯色边 / 两者 / 不裁）
  3. 导出格式可勾选：只要 PNG、只要 JPG、或两者都要
  4. 中文 / 全角括号路径、文件大小与逐张进度、单张失败不中断整批
  5. 默认导出一张完整合成 JPG，并在同名文件夹写出图层名称清单

仍然保持原版承诺：输出集中在源文件同名文件夹、不覆盖同名、不修改原稿、
后台线程转换、COM 显式释放、关掉 Photoshop 全部弹窗。

用法：
    双击运行（GUI）
    psd_to_png_gui.py --selftest                       # 不启动 Photoshop 的自检
    psd_to_png_gui.py --cli <文件或目录> [...] --formats png jpg
"""

import argparse
import json
import os
import sys
import tempfile
import threading
import time
import traceback
from types import SimpleNamespace
from pathlib import Path

APP_TITLE = "PSD / PSB 一键转 PNG · JPG"
VERSION = "2.5"
SUPPORTED_EXTS = {".psd", ".psb"}
FORMATS = ("png", "jpg")

STAGE_INFO = {
    "connecting": (3, "连接 Photoshop"),
    "opening": (8, "打开大文件"),
    "preparing": (30, "转换色彩/位深"),
    "layers": (38, "读取图层名称"),
    "trimming": (45, "裁切画布"),
    "saving_png": (60, "保存 PNG"),
    "saving_jpg": (80, "拼合并保存 JPG"),
    "closing": (95, "关闭内存文档"),
    "done": (100, "完成"),
}

TRIM_MODES = {
    "transparent": "裁掉透明边（默认，适合四周透明留白）",
    "topleft": "裁掉纯色边（按左上角像素颜色，适合白底留白）",
    "both": "先裁透明边，再裁纯色边（最彻底）",
    "none": "不裁切，保留原始画布",
}


# ---------------------------------------------------------------------------
# 路径与 ExtendScript
# ---------------------------------------------------------------------------
def js_escape(s) -> str:
    """把 Windows 路径安全地嵌进 ExtendScript 的单引号字符串字面量。"""
    return (str(s or "")
            .replace("\\", "\\\\")
            .replace("'", "\\'")
            .replace("\r", "\\r")
            .replace("\n", "\\n"))


def build_script(psd, png, jpg, trim_mode, formats, progress_file=None,
                 layer_list=None) -> str:
    """
    生成交给 Photoshop.Application.DoJavaScript 执行的 ExtendScript。

    formats 可含 'png' / 'jpg'，任意组合：
        只 png  → 保存带透明通道的 PNG
        只 jpg  → 拼合填白后保存 JPG
        两者     → 先存 PNG（保留透明），再拼合填白存 JPG
    """
    fmts = set(formats or ())
    progress_js = ""
    if progress_file:
        progress_js = (
            f"var pf=new File('{js_escape(progress_file)}');"
            "function pg(s){try{pf.open('a');pf.writeln(s);pf.close();}catch(e){}}")

    def stage(name):
        return f"pg('{name}');" if progress_file else ""

    trim_js = {
        "none": "",
        "transparent": "try{d.trim(TrimType.TRANSPARENT,true,true,true,true);}catch(e){}",
        "topleft": "try{d.trim(TrimType.TOPLEFT,true,true,true,true);}catch(e){}",
        "both": ("try{d.trim(TrimType.TRANSPARENT,true,true,true,true);}catch(e){}"
                 "try{d.trim(TrimType.TOPLEFT,true,true,true,true);}catch(e){}"),
    }.get(trim_mode, "")

    png_js = ""
    if "png" in fmts:
        png_js = (f"{stage('saving_png')}"
                  "var po=new PNGSaveOptions();po.interlaced=false;"
                  f"d.saveAs(new File('{js_escape(png)}'),po,true,Extension.LOWERCASE);")

    jpg_js = ""
    if "jpg" in fmts:
        # JPG 不支持 alpha：先 flatten 用背景色（默认白）填充透明，
        # 再降为 8 位/通道（JPG 不支持 16 位），最后保存。
        jpg_js = (f"{stage('saving_jpg')}"
                  "d.flatten();"
                  "if(d.bitsPerChannel==BitsPerChannelType.SIXTEEN)"
                  "d.bitsPerChannel=BitsPerChannelType.EIGHT;"
                  "var jo=new JPEGSaveOptions();jo.quality=12;"
                  "jo.embedColorProfile=true;"
                  "jo.formatOptions=FormatOptions.STANDARDBASELINE;"
                  f"d.saveAs(new File('{js_escape(jpg)}'),jo,true,Extension.LOWERCASE);")

    layer_js = ""
    if layer_list:
        layer_js = (
            f"{stage('layers')}"
            f"var lf=new File('{js_escape(layer_list)}');lf.encoding='UTF8';"
            "if(lf.open('w')){"
            "function dl(c,n){for(var i=0;i<c.layers.length;i++){"
            "var l=c.layers[i],p='';for(var j=0;j<n;j++)p+='  ';"
            "lf.writeln(p+l.name);if(l.typename=='LayerSet')dl(l,n+1);}}"
            "lf.writeln('PSD/PSB: '+d.name);"
            "lf.writeln('图层列表（含分组）:');dl(d,0);lf.close();}")

    return (
        "app.displayDialogs=DialogModes.NO;"
        f"{progress_js}"
        "var d=null;"
        "try{"
        f"{stage('opening')}"
        f"d=app.open(new File('{js_escape(psd)}'));"
        # PNG 不支持 32 位浮点通道，也不支持 CMYK —— 先做内存里的转换。
        # 文档最终以 DONOTSAVECHANGES 关闭，磁盘原稿不受任何影响。
        f"{stage('preparing')}"
        "if(d.bitsPerChannel==BitsPerChannelType.THIRTYTWO)"
        "d.bitsPerChannel=BitsPerChannelType.SIXTEEN;"
        "if(d.mode!=DocumentMode.RGB)d.changeMode(ChangeMode.RGB);"
        f"{layer_js}"
        f"{stage('trimming')}{trim_js}"
        f"{png_js}"
        f"{jpg_js}"
        f"}}finally{{{stage('closing')}if(d)d.close(SaveOptions.DONOTSAVECHANGES);"
        f"{stage('done')}}}"
    )


def collect(paths) -> list:
    """文件与文件夹混合输入 → 递归收集 .psd / .psb → 大小写不敏感去重。"""
    found = []
    seen = set()

    def add(p: Path):
        key = os.path.normcase(os.path.abspath(str(p)))
        if key in seen:
            return
        seen.add(key)
        found.append(p)

    for raw in paths:
        try:
            p = Path(raw)
            if p.is_dir():
                for child in sorted(p.rglob("*")):
                    try:
                        if child.is_file() and child.suffix.lower() in SUPPORTED_EXTS:
                            add(child)
                    except (OSError, PermissionError):
                        continue
            elif p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                add(p)
        except (OSError, ValueError):
            continue
    return found


def unique_outputs(src: Path, formats, use_folder=False) -> dict:
    """
    为同一份内容挑一个未被占用的 stem，PNG 与 JPG 共用它，保证成对。
    阶梯：name.png/.jpg → name_PNG.png/.jpg → name_PNG_2.png/.jpg → ...
    """
    fmts = set(formats or ())
    base = src.parent if str(src.parent) else Path(".")
    # 普通导出默认直接放在源文件旁；需要集中管理时可手动勾选同名文件夹。
    directory = base / src.stem if use_folder else base
    directory.mkdir(parents=True, exist_ok=True)
    stem = src.stem

    for i in range(10000):
        suffix = "" if i == 0 else ("_PNG" if i == 1 else f"_PNG_{i}")
        out = {}
        if "png" in fmts:
            out["png"] = directory / (stem + suffix + ".png")
        if "jpg" in fmts:
            out["jpg"] = directory / (stem + suffix + ".jpg")
        if out and all(not p.exists() for p in out.values()):
            layer_name = "图层列表.txt" if use_folder else f"{src.stem}_图层列表.txt"
            out["layers"] = directory / layer_name
            return out
    raise RuntimeError("同名输出过多（超过 10000 个），请先清理目录")


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} GB"


# ---------------------------------------------------------------------------
# 亚马逊高级 A+ 连屏切图
# ---------------------------------------------------------------------------
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
A_PLUS_PRESETS = {
    1464: (600, "电脑第一屏"),
    1200: (900, "手机第一屏"),
}


def _open_rgb(path):
    from PIL import Image
    image = Image.open(path)
    image.load()
    if image.mode != "RGB":
        converted = image.convert("RGB")
        image.close()
        image = converted
    return image


def _save_jpeg(image, path, quality=95):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path, "JPEG", quality=int(quality), subsampling=0, optimize=True)


def split_a_plus_fixed(src: Path, output_root=None, quality=95):
    """按内置电脑/手机规格，把一张高级 A+ 连屏图从上到下无缝切开。"""
    src = Path(src)
    image = _open_rgb(src)
    try:
        if image.width not in A_PLUS_PRESETS:
            raise ValueError(
                f"无法识别 {image.width}px 宽度；电脑端应为 1464px，手机端应为 1200px")
        height, prefix = A_PLUS_PRESETS[image.width]
        if image.height % height:
            raise ValueError(
                f"{src.name} 高度 {image.height}px 不能被单屏高度 {height}px 整除")
        root = Path(output_root) if output_root else src.parent
        folder = root / src.stem
        outputs = []
        for index, top in enumerate(range(0, image.height, height), 1):
            out = folder / f"{prefix}{index}.jpg"
            tile = image.crop((0, top, image.width, top + height))
            try:
                _save_jpeg(tile, out, quality)
            finally:
                tile.close()
            outputs.append(out)
        return outputs
    finally:
        image.close()


def _reference_groups(reference_root: Path, width: int, total_height: int):
    """找出宽度一致、累计高度等于连屏图的旧版图片组。"""
    from PIL import Image
    groups = {}
    for path in Path(reference_root).rglob("*"):
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTS:
            continue
        try:
            with Image.open(path) as im:
                if im.width == width:
                    groups.setdefault(path.parent, []).append((path, im.height))
        except Exception:
            continue
    valid = []
    for parent, items in groups.items():
        items.sort(key=lambda item: item[0].name)
        if sum(height for _, height in items) == total_height:
            valid.append((parent, items))
    return valid


def _normalized_a_plus_name(name):
    """忽略设计文件常见的“拷贝/copy”和空格后缀后比较模块名。"""
    value = Path(name).stem.lower().replace(" ", "").replace("_", "")
    for token in ("拷贝", "copy"):
        value = value.replace(token, "")
    return value


def _reference_module_groups(reference_root: Path, src: Path, width: int,
                             total_height: int):
    """识别完整 A+ 项目中的同名模块，或首页对应的第一屏分图。"""
    from PIL import Image
    by_parent = {}
    for path in Path(reference_root).rglob("*"):
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTS:
            continue
        try:
            if path.resolve() == Path(src).resolve():
                continue
        except OSError:
            pass
        try:
            with Image.open(path) as im:
                if im.width == width:
                    by_parent.setdefault(path.parent, []).append((path, im.height))
        except Exception:
            continue

    src_name = _normalized_a_plus_name(src.name)
    matched = []
    for parent, items in by_parent.items():
        items.sort(key=lambda item: item[0].name)
        direct = [item for item in items
                  if _normalized_a_plus_name(item[0].name) == src_name and
                  item[1] == total_height]
        if len(direct) == 1:
            matched.append((parent, direct))
            continue
        if "首页" in src_name:
            first_screen = [item for item in items if "第一屏" in item[0].stem]
            if first_screen and sum(height for _, height in first_screen) == total_height:
                matched.append((parent, first_screen))
    return matched


def split_a_plus_reference(src: Path, reference_root: Path, quality=95):
    """按照旧版文件夹的目录、文件名和尺寸，用新版连屏图原位更新。"""
    src, reference_root = Path(src), Path(reference_root)
    image = _open_rgb(src)
    try:
        groups = _reference_module_groups(
            reference_root, src, image.width, image.height)
        if not groups:
            groups = _reference_groups(reference_root, image.width, image.height)
        # 文件名中出现颜色/版本文件夹名时优先精确匹配，例如“粉色”“米白色”。
        named = [group for group in groups if group[0].name in src.stem]
        if len(named) == 1:
            parent, items = named[0]
        elif len(groups) == 1:
            parent, items = groups[0]
        elif not groups:
            raise ValueError(
                f"参考文件夹中没有宽 {image.width}px、累计高 {image.height}px 的旧版图片组")
        else:
            choices = "、".join(group[0].name for group in groups)
            raise ValueError(f"找到多个可能的参考组（{choices}），请让连屏文件名包含对应文件夹名")

        outputs, top = [], 0
        for out, height in items:
            tile = image.crop((0, top, image.width, top + height))
            try:
                _save_jpeg(tile, out, quality)
            finally:
                tile.close()
            outputs.append(out)
            top += height
        return outputs
    finally:
        image.close()


# ---------------------------------------------------------------------------
# Photoshop COM（只在真正转换时导入）
# ---------------------------------------------------------------------------
def photoshop():
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("缺少 pywin32，无法调用 Photoshop COM") from exc
    try:
        return win32com.client.Dispatch("Photoshop.Application")
    except Exception as exc:
        raise RuntimeError("没有找到 Photoshop，请先确认已安装桌面完整版") from exc


def convert_one(psd: Path, outputs: dict, trim_mode: str, formats,
                progress_cb=None) -> None:
    """在调用方线程里执行；COM 需要每个线程自己 CoInitialize。"""
    import pythoncom

    pythoncom.CoInitialize()
    marker = (Path(tempfile.gettempdir()) /
              f"psdtool_progress_{os.getpid()}_{threading.get_ident()}.txt")
    stop_monitor = threading.Event()
    monitor = None
    emitted_count = [0]
    try:
        try:
            marker.unlink(missing_ok=True)
        except OSError:
            pass
        if progress_cb:
            progress_cb("connecting")

            def watch_progress():
                while not stop_monitor.wait(0.2):
                    try:
                        stages = marker.read_text(encoding="utf-8").splitlines()
                    except (OSError, UnicodeError):
                        continue
                    for current in stages[emitted_count[0]:]:
                        if current:
                            progress_cb(current)
                        emitted_count[0] += 1

            monitor = threading.Thread(target=watch_progress, daemon=True)
            monitor.start()
        app = photoshop()
        script = build_script(str(psd),
                              outputs.get("png"), outputs.get("jpg"),
                              trim_mode, formats, marker if progress_cb else None,
                              outputs.get("layers"))
        try:
            # Photoshop 刚启动、正在载入插件或有短暂模态状态时会返回
            # RPC_E_SERVERCALL_RETRYLATER (0x8001010A)。大文件场景下等待后
            # 重试比立即判定失败可靠。
            for attempt in range(60):
                try:
                    app.DoJavaScript(script)
                    break
                except Exception as exc:
                    code = getattr(exc, "hresult", None)
                    if code is None and getattr(exc, "args", None):
                        code = exc.args[0]
                    if code not in (-2147417846, -2147418111) or attempt == 59:
                        raise
                    time.sleep(2)
        finally:
            try:
                del app
            except Exception:
                pass
    finally:
        stop_monitor.set()
        if monitor:
            monitor.join(timeout=1)
        if progress_cb:
            try:
                stages = marker.read_text(encoding="utf-8").splitlines()
                for final_stage in stages[emitted_count[0]:]:
                    if final_stage:
                        progress_cb(final_stage)
                    emitted_count[0] += 1
            except (OSError, UnicodeError):
                pass
        try:
            marker.unlink(missing_ok=True)
        except OSError:
            pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# 自检（不启动 Photoshop，不读取用户素材内容）
# ---------------------------------------------------------------------------
def selftest() -> int:
    import tempfile

    checks = []

    def check(name, ok, detail=""):
        checks.append({"name": name, "ok": bool(ok), "detail": detail})

    with tempfile.TemporaryDirectory(prefix="psdtool_test_") as td:
        root = Path(td)
        (root / "a.psd").write_bytes(b"")
        (root / "b.PSB").write_bytes(b"")
        (root / "c.txt").write_bytes(b"")
        sub = root / "子目录（全角）"
        sub.mkdir()
        (sub / "NTT72-D.psb").write_bytes(b"")

        got = collect([str(root)])
        names = sorted(p.name for p in got)
        check("同时收集 .psd / .psb / 大写扩展名",
              names == ["NTT72-D.psb", "a.psd", "b.PSB"], str(names))
        check("排除非 Photoshop 文件", "c.txt" not in names)
        check("递归进入中文/全角括号目录", "NTT72-D.psb" in names)

        dup = collect([str(root / "a.psd"), str(root / "a.PSD")])
        check("大小写不敏感去重", len(dup) == 1, str(len(dup)))

    tricky = r"F:\杨平-产品图片\01-产品图片（产品系列分类）\NTT系列\NTT72\O'Brien 的图.psb"

    js_all = build_script(tricky, r"F:\out\O'Brien.png", r"F:\out\O'Brien.jpg",
                          "transparent", ("png", "jpg"))
    check("反斜杠转义", "\\\\杨平" in js_all)
    check("单引号转义", "O\\'Brien" in js_all)
    check("中文与全角括号路径", "杨平-产品图片" in js_all and "（产品系列分类）" in js_all)
    check("裁切指令", "TrimType.TRANSPARENT" in js_all)
    check("PNG + JPG 同时保存",
          "PNGSaveOptions" in js_all and "JPEGSaveOptions" in js_all)
    check("JPG 前拼合填白 + 降 8 位",
          "d.flatten()" in js_all and "BitsPerChannelType.EIGHT" in js_all)
    check("不保存原稿", "DONOTSAVECHANGES" in js_all)
    check("关闭所有弹窗", "displayDialogs=DialogModes.NO" in js_all)

    js_png = build_script(tricky, r"F:\out\x.png", None, "both", ("png",))
    check("只选 PNG 时不生成 JPG 代码", "JPEGSaveOptions" not in js_png)
    check("只选 PNG 时仍生成 PNG 代码", "PNGSaveOptions" in js_png)
    check("both 模式含两次 trim", js_png.count("d.trim(") == 2)

    js_jpg = build_script(tricky, None, r"F:\out\x.jpg", "topleft", ("jpg",))
    check("只选 JPG 时不生成 PNG 代码", "PNGSaveOptions" not in js_jpg)
    check("只选 JPG 时仍生成 JPG 代码", "JPEGSaveOptions" in js_jpg)
    check("topleft 模式含 TOPLEFT", "TrimType.TOPLEFT" in js_jpg)
    check("未选的路径不写入脚本", "None" not in js_jpg)

    js_none = build_script(tricky, None, None, "none", ("png",))
    check("不裁切时不生成 trim", "d.trim(" not in js_none)

    with tempfile.TemporaryDirectory(prefix="psdtool_pair_") as td:
        d = Path(td)
        src = d / "NTT72-D.psb"
        src.write_bytes(b"")
        o1 = unique_outputs(src, ("png", "jpg"))
        for p in o1.values():
            p.write_bytes(b"")
        o2 = unique_outputs(src, ("png", "jpg"))
        check("输出成对且不覆盖",
              o1["png"].name == "NTT72-D.png" and
              o2["png"].name == "NTT72-D_PNG.png" and
              o2["jpg"].name == "NTT72-D_PNG.jpg",
              f"{sorted(o1.keys())} -> {o2['png'].name} / {o2['jpg'].name}")
        check("渲染模式输出到源文件旁",
              o1["jpg"].parent == src.parent)
        check("渲染模式图层清单不与其他 PSD 冲突",
              o1["layers"].name == "NTT72-D_图层列表.txt")
        folder_outputs = unique_outputs(src, ("jpg",), use_folder=True)
        check("主副图画板模式进入同名文件夹",
              folder_outputs["jpg"].parent.name == "NTT72-D" and
              folder_outputs["layers"].name == "图层列表.txt")
        o3 = unique_outputs(src, ("jpg",))
        check("只选 JPG 时仅输出 JPG 与图层清单",
              set(o3) == {"jpg", "layers"})

    js_layers = build_script(tricky, None, r"F:\out\x.jpg", "none", ("jpg",),
                             None, r"F:\out\图层列表.txt")
    check("递归读取图层与分组名称",
          "function dl" in js_layers and "LayerSet" in js_layers and
          "图层列表.txt" in js_layers)

    ok = all(c["ok"] for c in checks)
    report = {
        "selftest": "PASS" if ok else "FAIL",
        "version": VERSION,
        "supported_exts": sorted(SUPPORTED_EXTS),
        "formats": list(FORMATS),
        "trim_modes": list(TRIM_MODES),
        "checks": checks,
    }
    text = json.dumps(report, ensure_ascii=False, indent=2)
    print(text)
    # 冻结成 --windowed 后 stdout 可能不可见，同时落一份到临时目录便于验证
    try:
        import tempfile
        (Path(tempfile.gettempdir()) / "psdtool_selftest.json").write_text(
            text, encoding="utf-8")
    except Exception:
        pass
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class App:
    def __init__(self, startup_paths=None):
        import tkinter as tk
        from tkinter import filedialog, messagebox

        # windnd 通过 ctypes 安装 Win32 回调，在部分 64 位 Python/Tk 组合上
        # 拖入文件时会直接触发 ucrtbase.dll 0xc0000409，进程来不及抛出
        # Python 异常。tkinterdnd2 使用 TkDND，避免原生回调崩溃。
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
        except ImportError:
            DND_FILES = None
            TkinterDnD = None

        self.tk = tk
        self.filedialog = filedialog
        self.messagebox = messagebox
        self.DND_FILES = DND_FILES
        self.files = []
        self.busy = False
        self.stop_requested = threading.Event()

        self.root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()
        self.root.report_callback_exception = self.report_callback_exception
        self.root.title(APP_TITLE)
        self.root.geometry("640x580")
        self.root.minsize(580, 520)
        self.root.configure(bg="#f5f6f8")

        self._build_ui()

        # 窗口拖放使用 TkDND；不可用时仍可用按钮或拖到 EXE 图标。
        self.dnd_ok = False
        if DND_FILES:
            try:
                self.drop.drop_target_register(DND_FILES)
                self.drop.dnd_bind("<<Drop>>", self.on_drop)
                self.dnd_ok = True
            except Exception:
                pass

        if startup_paths:
            self.add_paths(startup_paths)

        self.log("就绪。" + ("" if self.dnd_ok else
                            " 窗口拖放不可用，请用按钮选择，或把文件拖到 EXE 图标上启动。"))
        self.log("提示：1 GB 以上的 PSB 打开可能需要数分钟，转换期间请勿操作 Photoshop。")

    def _build_ui(self):
        tk = self.tk
        from tkinter import ttk

        title = tk.Label(self.root, text="PSD / PSB 一键转 PNG · JPG",
                         font=("Microsoft YaHei UI", 17, "bold"),
                         bg="#f5f6f8", fg="#1f2733")
        title.pack(anchor="w", padx=20, pady=(16, 2))

        hint = tk.Label(self.root,
                        text="支持大画板 .psd / .psb · 自动裁到内容边界 · 导出格式可勾选",
                        font=("Microsoft YaHei UI", 9), bg="#f5f6f8", fg="#6b7280")
        hint.pack(anchor="w", padx=20, pady=(0, 10))

        self.drop = tk.Label(self.root,
                            text="↓\n把 PSD / PSB 文件或文件夹拖到这里\n"
                                 "默认导出到源文件旁边，不修改原稿",
                            font=("Microsoft YaHei UI", 12),
                            bg="#ffffff", fg="#2b6cb0", relief="solid", bd=1,
                            justify="center")
        self.drop.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # 第 1 行：工作模式
        mode_row = tk.Frame(self.root, bg="#f5f6f8")
        mode_row.pack(fill="x", padx=20, pady=(0, 6))
        tk.Label(mode_row, text="工作模式：", font=("Microsoft YaHei UI", 9),
                 bg="#f5f6f8").pack(side="left")
        self.output_mode = tk.StringVar(value="渲染图片（JPG 放源文件旁）")
        mode_combo = ttk.Combobox(mode_row, textvariable=self.output_mode,
                                  state="readonly", width=34,
                                  font=("Microsoft YaHei UI", 9))
        mode_combo["values"] = (
            "渲染图片（JPG 放源文件旁）",
            "主副图画板（建立同名文件夹）",
        )
        mode_combo.current(0)
        mode_combo.pack(side="left", padx=4)
        tk.Button(mode_row, text="高级 A+ 连屏模式", command=self.open_a_plus,
                  font=("Microsoft YaHei UI", 8), fg="#7c3aed").pack(side="right")

        # 第 2 行：裁切方式
        row1 = tk.Frame(self.root, bg="#f5f6f8")
        row1.pack(fill="x", padx=20, pady=(0, 6))
        tk.Label(row1, text="裁切方式：", font=("Microsoft YaHei UI", 9),
                 bg="#f5f6f8").pack(side="left")
        self.trim = tk.StringVar(value=TRIM_MODES["transparent"])
        combo = ttk.Combobox(row1, textvariable=self.trim, state="readonly",
                             width=36, font=("Microsoft YaHei UI", 9))
        combo["values"] = [TRIM_MODES[k] for k in TRIM_MODES]
        combo.current(0)
        combo.pack(side="left", padx=4)

        # 第 3 行：导出格式（可勾选）
        row2 = tk.Frame(self.root, bg="#f5f6f8")
        row2.pack(fill="x", padx=20, pady=(0, 8))
        tk.Label(row2, text="导出格式：", font=("Microsoft YaHei UI", 9),
                 bg="#f5f6f8").pack(side="left")
        self.var_png = tk.BooleanVar(value=False)
        self.var_jpg = tk.BooleanVar(value=True)
        tk.Checkbutton(row2, text="PNG（保留透明）", variable=self.var_png,
                       font=("Microsoft YaHei UI", 9), bg="#f5f6f8",
                       activebackground="#f5f6f8").pack(side="left", padx=(4, 10))
        tk.Checkbutton(row2, text="JPG（自动填白，质量 12）", variable=self.var_jpg,
                       font=("Microsoft YaHei UI", 9), bg="#f5f6f8",
                       activebackground="#f5f6f8").pack(side="left")
        tk.Button(row2, text="快速预设", command=self.fast_preset,
                   font=("Microsoft YaHei UI", 8)).pack(side="right")

        btns = tk.Frame(self.root, bg="#f5f6f8")
        btns.pack(fill="x", padx=20, pady=(0, 8))
        self.btn_files = tk.Button(btns, text="选择文件（可多选）", width=15,
                                   command=self.pick_files)
        self.btn_folder = tk.Button(btns, text="选择文件夹（递归）", width=15,
                                    command=self.pick_folder)
        self.btn_go = tk.Button(btns, text="开始转换", width=10,
                                command=self.start, bg="#2b6cb0", fg="white",
                                activebackground="#245a94", activeforeground="white")
        self.btn_stop = tk.Button(btns, text="停止队列", width=9,
                                  command=self.request_stop, state="disabled")
        self.btn_clear = tk.Button(btns, text="清空", width=6, command=self.clear)
        for b in (self.btn_files, self.btn_folder, self.btn_go,
                  self.btn_stop, self.btn_clear):
            b.pack(side="left", padx=(0, 8))

        progress_row = tk.Frame(self.root, bg="#f5f6f8")
        progress_row.pack(fill="x", padx=20, pady=(0, 6))
        tk.Label(progress_row, text="总进度", width=7, anchor="w",
                 font=("Microsoft YaHei UI", 9), bg="#f5f6f8").pack(side="left")
        self.progress = ttk.Progressbar(progress_row, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)

        current_row = tk.Frame(self.root, bg="#f5f6f8")
        current_row.pack(fill="x", padx=20, pady=(0, 6))
        self.current_progress_label = tk.StringVar(value="当前文件：空闲")
        tk.Label(current_row, textvariable=self.current_progress_label,
                 width=14, anchor="w", font=("Microsoft YaHei UI", 9),
                 bg="#f5f6f8").pack(side="left")
        # 空闲时使用 0 值确定进度，避免 Windows 主题在停止的动态进度条上
        # 仍显示一小块绿色，让人误以为程序卡住。
        self.current_progress = ttk.Progressbar(current_row, mode="determinate",
                                                maximum=100, value=0)
        self.current_progress.pack(side="left", fill="x", expand=True)

        self.status = tk.StringVar(value="等待文件")
        tk.Label(self.root, textvariable=self.status, anchor="w",
                 font=("Microsoft YaHei UI", 9), bg="#f5f6f8",
                 fg="#374151").pack(fill="x", padx=20)

        self.logbox = tk.Text(self.root, height=7, font=("Consolas", 9),
                              bg="#ffffff", fg="#374151", relief="solid", bd=1)
        self.logbox.pack(fill="both", expand=True, padx=20, pady=(6, 16))
        self.logbox.configure(state="disabled")

    # -- helpers ------------------------------------------------------------
    def log(self, msg: str):
        def _do():
            self.logbox.configure(state="normal")
            self.logbox.insert("end", msg + "\n")
            self.logbox.see("end")
            self.logbox.configure(state="disabled")
        self.ui(_do)

    def ui(self, fn):
        try:
            self.root.after(0, fn)
        except Exception:
            pass

    def report_callback_exception(self, exc_type, exc, tb):
        """记录 Tk 回调异常，避免拖放故障只表现为窗口消失。"""
        detail = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            path = Path(tempfile.gettempdir()) / "psdtool_gui_error.log"
            path.write_text(detail, encoding="utf-8")
        except Exception:
            pass
        try:
            self.log("! 界面回调异常，详情已写入 %TEMP%\\psdtool_gui_error.log")
            self.messagebox.showerror(APP_TITLE, str(exc))
        except Exception:
            pass

    def trim_key(self):
        label = self.trim.get()
        for k, v in TRIM_MODES.items():
            if v == label:
                return k
        return "transparent"

    def formats(self):
        out = []
        if self.var_png.get():
            out.append("png")
        if self.var_jpg.get():
            out.append("jpg")
        return tuple(out)

    def fast_preset(self):
        if self.busy:
            return
        self.trim.set(TRIM_MODES["none"])
        self.var_png.set(False)
        self.var_jpg.set(True)
        self.status.set("快速预设：不裁切，只输出 JPG")
        self.log("已启用快速预设：不裁切，只输出 JPG；文件更快，画布尺寸保持不变。")

    def open_a_plus(self):
        """打开独立的亚马逊高级 A+ 连屏切图窗口。"""
        tk = self.tk
        win = tk.Toplevel(self.root)
        win.title("亚马逊高级 A+ 连屏导图")
        win.geometry("680x480")
        win.minsize(620, 430)
        win.configure(bg="#f5f6f8")
        win.transient(self.root)

        tk.Label(win, text="亚马逊高级 A+ 连屏导图",
                 font=("Microsoft YaHei UI", 16, "bold"),
                 bg="#f5f6f8", fg="#1f2733").pack(
                     anchor="w", padx=18, pady=(14, 2))
        tk.Label(win,
                 text="新版连屏图按旧版文件名和尺寸覆盖；也可按 1464×600 / 1200×900 自动切图",
                 font=("Microsoft YaHei UI", 9), bg="#f5f6f8",
                 fg="#6b7280").pack(anchor="w", padx=18, pady=(0, 10))

        files = []
        listbox = tk.Listbox(win, height=7, font=("Microsoft YaHei UI", 9))
        listbox.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        def refresh():
            listbox.delete(0, "end")
            for path in files:
                listbox.insert("end", str(path))

        def add_images(paths):
            existing = {os.path.normcase(os.path.abspath(str(p))) for p in files}
            for raw in paths:
                path = Path(raw)
                if path.is_file() and path.suffix.lower() in IMAGE_EXTS:
                    key = os.path.normcase(os.path.abspath(str(path)))
                    if key not in existing:
                        files.append(path)
                        existing.add(key)
            refresh()

        def pick_images():
            selected = self.filedialog.askopenfilenames(
                parent=win, title="选择新版高级 A+ 连屏图",
                filetypes=[("图片", "*.jpg;*.jpeg;*.png;*.webp;*.bmp")])
            if selected:
                add_images(selected)

        if self.DND_FILES:
            try:
                listbox.drop_target_register(self.DND_FILES)
                listbox.dnd_bind("<<Drop>>", lambda event: (
                    add_images(self.root.tk.splitlist(event.data)), "copy")[1])
            except Exception:
                pass

        mode = tk.StringVar(value="reference")
        mode_row = tk.Frame(win, bg="#f5f6f8")
        mode_row.pack(fill="x", padx=18, pady=(0, 6))
        tk.Radiobutton(mode_row, text="参考旧版文件夹（推荐）", variable=mode,
                       value="reference", bg="#f5f6f8").pack(side="left")
        tk.Radiobutton(mode_row, text="固定尺寸自动切图", variable=mode,
                       value="fixed", bg="#f5f6f8").pack(side="left", padx=12)

        reference = tk.StringVar()
        ref_row = tk.Frame(win, bg="#f5f6f8")
        ref_row.pack(fill="x", padx=18, pady=(0, 7))
        tk.Label(ref_row, text="旧版/输出文件夹：", bg="#f5f6f8").pack(side="left")
        tk.Entry(ref_row, textvariable=reference).pack(
            side="left", fill="x", expand=True, padx=5)

        def pick_reference():
            selected = self.filedialog.askdirectory(
                parent=win, title="选择旧版高级 A+ 图片所在文件夹")
            if selected:
                reference.set(selected)

        tk.Button(ref_row, text="选择", command=pick_reference).pack(side="left")

        quality = tk.IntVar(value=95)
        option_row = tk.Frame(win, bg="#f5f6f8")
        option_row.pack(fill="x", padx=18, pady=(0, 8))
        tk.Label(option_row, text="JPG 质量：", bg="#f5f6f8").pack(side="left")
        tk.Spinbox(option_row, from_=80, to=100, textvariable=quality,
                   width=5).pack(side="left")
        tk.Label(option_row, text="  切图顺序：从上到下；不缩放、不留缝",
                 bg="#f5f6f8", fg="#6b7280").pack(side="left")

        status = tk.StringVar(value="请添加新版连屏图")
        tk.Label(win, textvariable=status, anchor="w", bg="#f5f6f8",
                 fg="#374151").pack(fill="x", padx=18)
        buttons = tk.Frame(win, bg="#f5f6f8")
        buttons.pack(fill="x", padx=18, pady=(8, 14))
        tk.Button(buttons, text="选择新版连屏图（可多选）",
                  command=pick_images).pack(side="left")
        tk.Button(buttons, text="清空", command=lambda: (files.clear(), refresh())).pack(
                      side="left", padx=7)

        start_button = tk.Button(buttons, text="开始导图", bg="#7c3aed", fg="white")
        start_button.pack(side="right")

        def run_worker(selected, selected_mode, folder, jpg_quality):
            succeeded, failed = 0, []
            for src in selected:
                try:
                    if selected_mode == "reference":
                        outputs = split_a_plus_reference(src, Path(folder), jpg_quality)
                    else:
                        outputs = split_a_plus_fixed(
                            src, Path(folder) if folder else None, jpg_quality)
                    succeeded += 1
                    self.log(f"√ 高级A+：{src.name} → {len(outputs)} 张")
                except Exception as exc:
                    failed.append(f"{src.name}：{exc}")

            def done():
                start_button.configure(state="normal")
                status.set(f"完成：成功 {succeeded}，失败 {len(failed)}")
                if failed:
                    self.messagebox.showwarning(
                        "高级 A+ 连屏导图", "\n".join(failed), parent=win)
                else:
                    self.messagebox.showinfo(
                        "高级 A+ 连屏导图",
                        f"已完成 {succeeded} 张连屏图的切图。", parent=win)
            self.ui(done)

        def start_a_plus():
            if not files:
                self.messagebox.showinfo(
                    "高级 A+ 连屏导图", "请先选择或拖入新版连屏图", parent=win)
                return
            selected_mode = mode.get()
            folder = reference.get().strip()
            if selected_mode == "reference" and not folder:
                self.messagebox.showinfo(
                    "高级 A+ 连屏导图", "请选择旧版图片所在文件夹", parent=win)
                return
            start_button.configure(state="disabled")
            status.set(f"正在处理 {len(files)} 张连屏图……")
            threading.Thread(
                target=run_worker,
                args=(list(files), selected_mode, folder, quality.get()),
                daemon=True).start()

        start_button.configure(command=start_a_plus)

    def add_paths(self, paths):
        if self.busy:
            return
        got = collect(paths)
        existing = {os.path.normcase(os.path.abspath(str(p))) for p in self.files}
        added = [p for p in got
                 if os.path.normcase(os.path.abspath(str(p))) not in existing]
        self.files.extend(added)
        if added:
            total = sum(p.stat().st_size for p in self.files if p.exists())
            self.status.set(f"已加入 {len(self.files)} 个文件，共 {human_size(total)}；点「开始转换」")
            for p in added[:8]:
                self.log(f"+ {p.name}  ({human_size(p.stat().st_size)})")
            if len(added) > 8:
                self.log(f"+ … 另有 {len(added) - 8} 个文件")
        elif not got:
            self.status.set("没有检测到 PSD / PSB 文件")
            self.log("! 输入里没有 .psd 或 .psb；请确认扩展名（.psb 已支持）")

    def on_drop(self, event):
        try:
            # Tk 的 splitlist 能正确解析带空格、中文和花括号包裹的路径。
            paths = list(self.root.tk.splitlist(event.data))
            self.add_paths(paths)
        except Exception:
            self.log("! 拖放数据解析失败，请改用按钮选择")
        # TkDND 回调必须返回一个动作；返回 None 可能被 Tcl 当作无效动作。
        return getattr(event, "action", "copy") or "copy"

    def pick_files(self):
        paths = self.filedialog.askopenfilenames(
            title="选择 Photoshop 文件",
            filetypes=[("Photoshop PSD / PSB", "*.psd;*.psb"),
                       ("Photoshop 大型文档", "*.psb"),
                       ("Photoshop 文档", "*.psd")])
        if paths:
            self.add_paths(paths)

    def pick_folder(self):
        d = self.filedialog.askdirectory(title="选择包含 PSD / PSB 的文件夹")
        if d:
            self.add_paths([d])

    def clear(self):
        if self.busy:
            return
        self.files = []
        self.status.set("等待文件")
        self.log("已清空列表。")

    def request_stop(self):
        if not self.busy:
            return
        self.stop_requested.set()
        self.btn_stop.configure(state="disabled")
        self.status.set("已请求停止：当前文件完成后停止后续队列")
        self.log("! 已请求停止；为保护原稿，当前 Photoshop 操作完成后再停止。")

    def update_stage(self, name, stage):
        percent, label = STAGE_INFO.get(stage, (0, stage))
        short_name = name if len(name) <= 14 else name[:12] + "…"
        self.current_progress.configure(mode="determinate", value=percent)
        self.current_progress_label.set(f"{label}：{short_name}")

    # -- conversion ---------------------------------------------------------
    def start(self):
        if self.busy:
            return
        if not self.files:
            self.messagebox.showinfo(APP_TITLE, "请先拖入或选择 PSD / PSB 文件")
            return
        formats = self.formats()
        if not formats:
            self.messagebox.showwarning(APP_TITLE, "至少勾选一种导出格式（PNG 或 JPG）")
            return

        self.busy = True
        self.stop_requested.clear()
        for b in (self.btn_files, self.btn_folder, self.btn_go, self.btn_clear):
            b.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        self.current_progress.stop()
        self.current_progress.configure(mode="determinate", value=0)
        self.current_progress_label.set("当前文件：准备中")

        files = list(self.files)
        trim_mode = self.trim_key()
        use_folder = self.output_mode.get().startswith("主副图画板")
        mode_label = "主副图画板" if use_folder else "渲染图片"
        self.log(f"— 开始：模式={mode_label}，{len(files)} 个文件，"
                 f"裁切={trim_mode}，格式={'+'.join(formats)}")

        threading.Thread(target=self.worker,
                         args=(files, trim_mode, formats, use_folder),
                         daemon=True).start()

    def worker(self, files, trim_mode, formats, use_folder=False):
        ok = 0
        failed = 0
        processed = 0
        for i, src in enumerate(files, 1):
            if self.stop_requested.is_set():
                break
            name = src.name
            try:
                size = src.stat().st_size
            except OSError:
                size = 0
            self.ui(lambda i=i, name=name, size=size: self.status.set(
                f"[{i}/{len(files)}] 正在处理 {name}  ({human_size(size)})"
                " — 大画板请耐心等待"))
            self.ui(lambda name=name: self.update_stage(name, "connecting"))
            try:
                outputs = unique_outputs(src, formats, use_folder)
                convert_one(
                    src, outputs, trim_mode, formats,
                    progress_cb=lambda stage, name=name: self.ui(
                        lambda stage=stage, name=name: self.update_stage(name, stage)))
                ok += 1
                produced = " + ".join(p.name for p in outputs.values())
                self.ui(lambda name=name, produced=produced: self.log(
                    f"√ {name} → {produced}"))
            except Exception as exc:
                failed += 1
                msg = str(exc)
                if "Photoshop" in msg or "COM" in msg.upper():
                    msg = "Photoshop 忙或无法读取文件：" + msg
                self.ui(lambda name=name, msg=msg: self.log(f"× {name}：{msg}"))
            self.ui(lambda: (
                self.current_progress.stop(),
                self.current_progress.configure(mode="determinate", value=0)))
            self.ui(lambda i=i: self.progress.configure(value=i))
            processed = i

        done_ok, done_fail = ok, failed
        stopped = processed < len(files) and self.stop_requested.is_set()
        remaining = len(files) - processed
        self.ui(lambda: self.finish(done_ok, done_fail, stopped, remaining))

    def finish(self, ok, failed, stopped=False, remaining=0):
        self.busy = False
        self.current_progress.stop()
        self.current_progress.configure(mode="determinate", value=0)
        self.current_progress_label.set("当前文件：空闲")
        self.btn_stop.configure(state="disabled")
        for b in (self.btn_files, self.btn_folder, self.btn_go, self.btn_clear):
            b.configure(state="normal")
        if stopped:
            self.status.set(f"已停止：成功 {ok}，失败 {failed}，未处理 {remaining}")
            self.log(f"— 队列已停止：成功 {ok}，失败 {failed}，未处理 {remaining}")
        else:
            self.status.set(f"完成：成功 {ok} 个，失败 {failed} 个")
            self.log(f"— 全部结束：成功 {ok}，失败 {failed}")
        if failed:
            self.messagebox.showwarning(APP_TITLE,
                                        f"完成，但有 {failed} 个文件失败。\n详情见下方日志。")

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
def run_cli(paths, trim_mode, formats):
    log_path = Path(tempfile.gettempdir()) / "psdtool_cli.log"

    def emit(message):
        text = str(message)
        try:
            if sys.stdout is not None:
                print(text, flush=True)
                return
        except (OSError, ValueError):
            pass
        with log_path.open("a", encoding="utf-8") as stream:
            stream.write(text + "\n")

    try:
        log_path.unlink(missing_ok=True)
    except OSError:
        pass
    files = collect(paths)
    if not files:
        emit(json.dumps({"error": "没有检测到 PSD / PSB 文件", "inputs": paths},
                        ensure_ascii=False))
        return 2
    emit(f"共 {len(files)} 个文件，裁切={trim_mode}，格式={'+'.join(formats)}")
    ok = failed = 0
    for i, src in enumerate(files, 1):
        try:
            outputs = unique_outputs(src, formats)
            emit(f"[{i}/{len(files)}] {src.name} ({human_size(src.stat().st_size)}) ...")
            convert_one(src, outputs, trim_mode, formats,
                        progress_cb=lambda stage: emit(
                            f"    [{STAGE_INFO.get(stage, (0, stage))[1]}]"))
            emit("    -> " + " + ".join(p.name for p in outputs.values()))
            ok += 1
        except Exception as exc:
            emit(f"    FAILED: {exc}")
            emit(traceback.format_exc())
            failed += 1
    emit(json.dumps({"ok": ok, "failed": failed}, ensure_ascii=False))
    return 0 if failed == 0 else 1


def main():
    ap = argparse.ArgumentParser(description=APP_TITLE)
    ap.add_argument("--selftest", action="store_true", help="不启动 Photoshop 的自检")
    ap.add_argument("--probe", action="store_true", help="创建 GUI 后自动退出，用于验证冻结包")
    ap.add_argument("--drop-probe", nargs="+", metavar="PATH",
                    help="模拟一次拖放回调后退出，用于验证冻结包拖放入口")
    ap.add_argument("--cli", nargs="+", metavar="PATH", help="命令行批量转换")
    ap.add_argument("--trim", choices=list(TRIM_MODES), default="transparent")
    ap.add_argument("--formats", nargs="+", choices=list(FORMATS),
                    default=list(FORMATS), help="导出格式，可多选：png jpg")
    ap.add_argument("--version", action="version", version=VERSION)
    args, unknown = ap.parse_known_args()

    if args.selftest:
        return selftest()
    if args.probe:
        try:
            import PIL
            app = App()
            app.root.after(800, app.root.destroy)
            app.run()
            result = {"gui": "OK", "dnd": app.dnd_ok, "tk": app.tk.TkVersion,
                      "pillow": PIL.__version__, "a_plus": True}
        except Exception:
            result = {"gui": "FAIL", "error": traceback.format_exc()}
        (Path(tempfile.gettempdir()) / "psdtool_probe.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False))
        return 0 if result.get("gui") == "OK" else 1
    if args.drop_probe:
        try:
            app = App()
            event = SimpleNamespace(data=app.root.tk.call("list", *args.drop_probe),
                                    action="copy")
            action = app.on_drop(event)
            result = {"drop": "OK" if len(app.files) >= 1 else "FAIL",
                      "action": action, "count": len(app.files),
                      "names": [p.name for p in app.files],
                      "dnd": app.dnd_ok}
            app.root.destroy()
        except Exception:
            result = {"drop": "FAIL", "error": traceback.format_exc()}
        (Path(tempfile.gettempdir()) / "psdtool_drop_probe.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        return 0 if result.get("drop") == "OK" else 1
    if args.cli:
        formats = tuple(dict.fromkeys(args.formats))
        if not formats:
            print("至少指定一种导出格式")
            return 2
        return run_cli(args.cli, args.trim, formats)

    # 把文件/文件夹拖到 EXE 图标上启动时，路径会从 argv 传进来
    startup = [a for a in unknown if not a.startswith("--")]
    try:
        App(startup).run()
    except Exception:
        traceback.print_exc()
        try:
            import tkinter.messagebox as mb
            mb.showerror(APP_TITLE, traceback.format_exc())
        except Exception:
            pass
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
