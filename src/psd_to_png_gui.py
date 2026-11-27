#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PSD / PSB 一键转 PNG · JPG  ·  增强版 v3.0

面向「PS 大画板 / 大画布」场景：画布很大、内容只占中间一小块、
四周大量透明或白色留白；文件常常是 1~2 GB 的 .psb。

相对原版（只认 .psd、不裁切、只出 PNG）新增：
  1. 支持 .psd 与 .psb（Photoshop 大型文档）
  2. 导出前自动裁切到内容边界（透明边 / 纯色边 / 两者 / 不裁）
  3. 导出格式可勾选：只要 PNG、只要 JPG、或两者都要
  4. 中文 / 全角括号路径、文件大小与逐张进度、单张失败不中断整批
  5. 渲染模式输出完整合成 JPG；画板模式按画板/分组名称逐张输出

仍然保持原版承诺：不覆盖同名输出、转换修改只发生在内存副本；为让
Windows 显示缩略图，PSD 会额外写回兼容合成预览（不改变图层与画面）。

用法：
    双击运行（GUI）
    psd_to_png_gui.py --selftest                       # 不启动 Photoshop 的自检
    psd_to_png_gui.py --cli <文件或目录> [...] --formats png jpg
"""

import argparse
import json
import math
import os
import re
import sys
import tempfile
import threading
import time
import traceback
from types import SimpleNamespace
from pathlib import Path

APP_TITLE = "PSD / PSB 一键转 PNG · JPG"
VERSION = "3.0"
SUPPORTED_EXTS = {".psd", ".psb"}
FORMATS = ("png", "jpg")


def resource_path(relative: str) -> Path:
    """返回开发环境或 PyInstaller 冻结包内的资源路径。"""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / relative

# ---- UI 主题配色（浅色，统一卡片与品牌层次）----
C_BG = "#eef1f6"          # 窗口底色
C_CARD = "#ffffff"        # 卡片底
C_HEADER = "#111827"      # 顶部品牌栏
C_PRIMARY = "#2563eb"     # 主色（蓝）
C_PRIMARY_D = "#1d4ed8"
C_PURPLE = "#7c3aed"      # 高级 A+ 主色（紫）
C_PURPLE_D = "#6d28d9"
C_INK = "#1f2937"         # 主文字
C_SUB = "#6b7280"         # 次文字
C_MUTED = "#9ca3af"
C_BORDER = "#e5e7eb"      # 边框
C_SOFT_BLUE = "#eff6ff"   # 浅蓝底
C_BLUE_LINE = "#bfdbfe"
C_SOFT_PURPLE = "#f5f3ff"
C_PURPLE_LINE = "#ddd6fe"

STAGE_INFO = {
    "connecting": (3, "连接 Photoshop"),
    "opening": (8, "打开大文件"),
    "preflight": (18, "预检画板"),
    "thumbnail": (15, "写入 PSD 缩略图"),
    "preparing": (30, "转换色彩/位深"),
    "layers": (38, "读取图层名称"),
    "trimming": (45, "裁切画布"),
    "saving_png": (60, "保存 PNG"),
    "saving_jpg": (80, "拼合并保存 JPG"),
    "optimizing_jpg": (90, "优化 JPG 体积"),
    "closing": (95, "关闭内存文档"),
    "done": (100, "完成"),
}


class ConversionCancelled(RuntimeError):
    """用户在一个安全检查点取消了当前转换。"""

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
                 layer_list=None, write_thumbnail=True) -> str:
    """
    生成交给 Photoshop.Application.DoJavaScript 执行的 ExtendScript。

    formats 可含 'png' / 'jpg'，任意组合：
        只 png  → 保存带透明通道的 PNG
        只 jpg  → 拼合填白后保存 JPG
        两者     → 先存 PNG（保留透明），再拼合填白存 JPG
    """
    fmts = set(formats or ())
    progress_js = "function pg(s){}"
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

    # Windows 资源管理器读取的是 PSD 内嵌的合成预览。仅对 PSD 写回兼容预览；
    # PSB 往往超过 1 GB，整稿重存代价过高，保持原样。
    thumbnail_js = ""
    if write_thumbnail and Path(psd).suffix.lower() == ".psd":
        thumbnail_js = (
            f"{stage('thumbnail')}"
            "var oldMC=null;try{"
            "oldMC=app.preferences.maximizeCompatibility;"
            "app.preferences.maximizeCompatibility=QueryStateType.ALWAYS;"
            # 轻量制造一次可撤销的脏状态，最终可见性保持不变，促使 Photoshop
            # 重新保存并嵌入合成预览，而不是因文档无变化跳过写盘。
            "var av=d.activeLayer.visible;d.activeLayer.visible=!av;"
            "d.activeLayer.visible=av;d.save();"
            "}catch(te){}finally{try{if(oldMC!==null)"
            "app.preferences.maximizeCompatibility=oldMC;}catch(re){}}")

    return (
        "app.displayDialogs=DialogModes.NO;"
        f"{progress_js}"
        "var d=null;"
        "try{"
        f"{stage('opening')}"
        f"d=app.open(new File('{js_escape(psd)}'));"
        f"{thumbnail_js}"
        # PNG 不支持 32 位浮点通道，也不支持 CMYK —— 先做内存里的转换。
        # 位深、模式、裁切和拼合只用于内存副本，最终不写回原稿。
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


def build_artboard_script(psd, output_dir, manifest, progress_file=None,
                          cancel_file=None) -> str:
    """生成“逐画板导出 JPG”的 ExtendScript，文件名直接使用画板名称。"""
    progress_js = "function pg(s){}"
    if progress_file:
        progress_js = (
            f"var pf=new File('{js_escape(progress_file)}');"
            "function pg(s){try{pf.open('a');pf.writeln(s);pf.close();}catch(e){}}")

    return (
        "app.displayDialogs=DialogModes.NO;" + progress_js +
        "var d=null;var mf=null;var cf=new File('" + js_escape(cancel_file) + "');"
        "function cleanName(s){return s.replace(/[\\\\\\/:*?\"<>|]/g,'_')"
        ".replace(/^\\s+|\\s+$/g,'')||'画板';}"
        "function isNativeArtboard(layer){var r=new ActionReference();"
        "r.putIdentifier(charIDToTypeID('Lyr '),layer.id);var x=executeActionGet(r),"
        "e=stringIDToTypeID('artboardEnabled');"
        "return x.hasKey(e)&&x.getBoolean(e);}"
        "function artboardRect(layer){"
        "var r=new ActionReference();r.putIdentifier(charIDToTypeID('Lyr '),layer.id);"
        "var x=executeActionGet(r),e=stringIDToTypeID('artboardEnabled'),"
        "k=stringIDToTypeID('artboard');"
        "if(!x.hasKey(e)||!x.getBoolean(e)||!x.hasKey(k)){"
        "if(layer.typename!='LayerSet')return null;var b=layer.bounds;"
        "return [b[0].as('px'),b[1].as('px'),b[2].as('px'),b[3].as('px')];}"
        "var a=x.getObjectValue(k),q=a.getObjectValue(stringIDToTypeID('artboardRect'));"
        "return [q.getUnitDoubleValue(stringIDToTypeID('left')),"
        "q.getUnitDoubleValue(stringIDToTypeID('top')),"
        "q.getUnitDoubleValue(stringIDToTypeID('right')),"
        "q.getUnitDoubleValue(stringIDToTypeID('bottom'))];}"
        "function overviewRect(){var u=null;for(var i=0;i<d.layers.length;i++){"
        "if(!isNativeArtboard(d.layers[i]))continue;var b=artboardRect(d.layers[i]);"
        "if(!u)u=[b[0],b[1],b[2],b[3]];else{u[0]=Math.min(u[0],b[0]);"
        "u[1]=Math.min(u[1],b[1]);u[2]=Math.max(u[2],b[2]);"
        "u[3]=Math.max(u[3],b[3]);}}return u;}"
        "try{pg('opening');d=app.open(new File('" + js_escape(psd) + "'));"
        "mf=new File('" + js_escape(manifest) + "');mf.encoding='UTF8';mf.open('w');"
        "var count=0,total=0,baseState=d.activeHistoryState;pg('preflight');"
        "for(var p=0;p<d.layers.length;p++){var pr=artboardRect(d.layers[p]);"
        "if(pr){total++;var pw=Math.max(0,Math.round(pr[2]-pr[0])),"
        "ph=Math.max(0,Math.round(pr[3]-pr[1]));"
        "mf.writeln('META:'+cleanName(d.layers[p].name)+'|'+pw+'x'+ph);"
        "if(pw===0||ph===0||"
        "(d.layers[p].typename=='LayerSet'&&d.layers[p].layers.length===0))"
        "mf.writeln('WARN:'+cleanName(d.layers[p].name)+' 可能为空画板');}}"
        "pg('preflight_done:'+total);"
        "for(var i=0;i<d.layers.length;i++){var rect=artboardRect(d.layers[i]);"
        "if(!rect)continue;"
        "if(cf.exists){mf.writeln('CANCELLED');break;}"
        "var n=cleanName(d.layers[i].name),overview=!isNativeArtboard(d.layers[i])"
        "&&/(整体|总览)/.test(n);if(overview){var ur=overviewRect();if(ur)rect=ur;}"
        "var out=new File('" +
        js_escape(output_dir) + "/'+n+'.jpg'),z=2;"
        "while(out.exists){out=new File('" + js_escape(output_dir) +
        "/'+n+'-'+z+'.jpg');z++;}"
        "pg('board:'+String(count+1)+':'+String(total)+':'+n);"
        "try{for(var j=0;j<d.layers.length;j++){"
        "if(d.layers[j].typename=='LayerSet')d.layers[j].visible="
        "(overview?isNativeArtboard(d.layers[j]):j==i);}app.refresh();"
        # 原生画板文档在拼合前调用 crop 可能继续保留整张画板工作区；
        # 先拼合内存副本，再按已记录的画板绝对坐标裁切，尺寸才准确。
        "d.flatten();d.crop(rect);"
        "if(d.mode!=DocumentMode.RGB)d.changeMode(ChangeMode.RGB);"
        "if(d.bitsPerChannel!=BitsPerChannelType.EIGHT)"
        "d.bitsPerChannel=BitsPerChannelType.EIGHT;"
        "var jo=new JPEGSaveOptions();jo.quality=12;jo.embedColorProfile=true;"
        "jo.formatOptions=FormatOptions.STANDARDBASELINE;pg('saving_jpg');"
        "d.saveAs(out,jo,true,Extension.LOWERCASE);mf.writeln('OUTPUT:'+out.fsName);count++;"
        "}catch(be){mf.writeln('BOARD_ERROR:'+n+'：'+be.toString());}"
        "finally{d.activeHistoryState=baseState;}}"
        "if(total===0)throw new Error('没有检测到 Photoshop 画板或顶层图层组');"
        "}catch(err){try{if(mf)mf.writeln('ERROR:'+err.toString()+' line='+err.line);"
        "}catch(we){}throw err;}finally{try{if(mf)mf.close();}catch(e){}pg('closing');"
        "if(d)d.close(SaveOptions.DONOTSAVECHANGES);pg('done');}"
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


def collect_aplus(paths) -> list:
    """高级 A+ 输入收集：图片与 PSD / PSB 都可以，递归文件夹、去重。

    普通导图只收 PSD / PSB；高级 A+ 连屏图的来源既可能是设计源稿，
    也可能是 Photoshop 已经导出的 JPG / PNG 连屏图，因此这里放宽类型。
    """
    allowed = IMAGE_EXTS | SUPPORTED_EXTS
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
                        if child.is_file() and child.suffix.lower() in allowed:
                            add(child)
                    except (OSError, PermissionError):
                        continue
            elif p.is_file() and p.suffix.lower() in allowed:
                add(p)
        except (OSError, ValueError):
            continue
    return found


def clean_drop_path(raw):
    """清洗 Tk DND 返回的路径，修复中文/无空格反斜杠路径被 Tcl 转义的问题。

    tkinterdnd2 把拖入路径交给 Tcl 的 splitlist；当路径不含空格时，Windows
    反斜杠可能被当作转义字符，或者被多余花括号包裹。这里统一还原成原始
    Windows 绝对路径；若清洗后仍不是有效路径，则回退到原始字符串。
    """
    text = str(raw).strip()
    if not text:
        return text
    # 去掉 splitlist 可能保留的外层花括号与成对引号
    if len(text) >= 2 and text[0] == "{" and text[-1] == "}":
        text = text[1:-1]
    if len(text) >= 2 and text[0] == text[-1] and text[0] in ("'", '"'):
        text = text[1:-1]
    candidate = text
    if Path(candidate).exists():
        return candidate
    # Tcl 把 \x 反斜杠转义吃掉的情况：把常见被吞的反斜杠补回
    # 形如 D:2026P5S... → 按盘符与已知目录结构重建失败时，回退原始值
    return str(raw).strip()


def clean_drop_list(items):
    return [clean_drop_path(item) for item in items]


def unique_outputs(src: Path, formats, use_folder=False, conflict="number"):
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
        available = out and all(not p.exists() for p in out.values())
        if conflict == "overwrite" or available:
            return out
        if conflict == "skip":
            return None
    raise RuntimeError("同名输出过多（超过 10000 个），请先清理目录")


def inherit_source_timestamp(src: Path, outputs):
    """让导出物继承源文件时间，按日期分组时仍显示在源文件附近。"""
    try:
        stat = Path(src).stat()
        times = (stat.st_atime_ns, stat.st_mtime_ns)
    except OSError:
        return
    parents = set()
    for output in outputs or ():
        path = Path(output)
        try:
            if path.is_file():
                os.utime(path, ns=times)
                parents.add(path.parent)
        except OSError:
            continue
    # 画板模式会新建同名文件夹；同步文件夹时间，避免它被分到“今天”。
    for parent in parents:
        if parent != Path(src).parent:
            try:
                os.utime(parent, ns=times)
            except OSError:
                pass


# 工作模式：key 用于持久化，display/value 用于界面，避免中文文案漂移。
MODE_DEFS = (
    ("render", "渲染图片", "JPG 放源文件旁", "渲染图片（JPG 放源文件旁）"),
    ("artboard", "主副图画板", "自动建立文件夹", "主副图画板（建立同名文件夹）"),
    ("aplus", "高级 A+", "电脑端 / 手机端", "高级 A+ 连屏（电脑端 / 手机端）"),
)
MODE_VALUES = {key: value for key, _, _, value in MODE_DEFS}
MODE_KEYS = {value: key for key, _, _, value in MODE_DEFS}

# 同名文件处理策略
CONFLICT_CHOICES = ("number", "overwrite", "skip")

# 用户设置默认值（不保存 PSD/PSB 文件列表与任何隐私内容）
DEFAULT_SETTINGS = {
    "mode": "render",
    "trim": "transparent",
    "png": False,
    "jpg": True,
    "conflict": "number",
    "aplus_quality": 95,
    "aplus_reference": "",
}

SETTINGS_PATH = (Path(os.environ.get("APPDATA", tempfile.gettempdir())) /
                 "PSDTool" / "settings.json")


def normalize_settings(data):
    """只接受已知键与合法值，其余一律回退默认；杜绝坏设置导致闪退。"""
    out = dict(DEFAULT_SETTINGS)
    if not isinstance(data, dict):
        return out
    mode = data.get("mode")
    if mode in MODE_VALUES:
        out["mode"] = mode
    elif mode in MODE_KEYS:  # 兼容早期直接存显示文本
        out["mode"] = MODE_KEYS[mode]
    trim = data.get("trim")
    if isinstance(trim, str) and trim in TRIM_MODES:
        out["trim"] = trim
    conflict = data.get("conflict")
    if isinstance(conflict, str) and conflict in CONFLICT_CHOICES:
        out["conflict"] = conflict
    for flag in ("png", "jpg"):
        value = data.get(flag)
        if isinstance(value, bool):
            out[flag] = value
    quality = data.get("aplus_quality")
    if (isinstance(quality, int) and not isinstance(quality, bool)
            and 80 <= quality <= 100):
        out["aplus_quality"] = quality
    reference = data.get("aplus_reference")
    if isinstance(reference, str):
        out["aplus_reference"] = reference
    return out


def load_settings():
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return dict(DEFAULT_SETTINGS)
    return normalize_settings(data)


def save_settings(data):
    try:
        SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS_PATH.write_text(
            json.dumps(normalize_settings(data), ensure_ascii=False, indent=2),
            encoding="utf-8")
    except OSError:
        pass


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


def _content_row_runs(image, sample_step=12, white_limit=248, bridge=4):
    """返回列图中的非白色纵向区段；允许图片内部出现少量纯白行。"""
    gray = image.convert("L")
    try:
        pixels = gray.load()
        width, height = gray.size
        active = []
        xs = list(range(0, width, max(1, sample_step)))
        if width - 1 not in xs:
            xs.append(width - 1)
        for y in range(height):
            active.append(any(pixels[x, y] < white_limit for x in xs))
    finally:
        gray.close()

    # 填平很短的白线，避免画面中的白色分隔线把一张连屏误拆开。
    index = 0
    while index < len(active):
        if active[index]:
            index += 1
            continue
        end = index
        while end < len(active) and not active[end]:
            end += 1
        if index > 0 and end < len(active) and end - index <= bridge:
            for pos in range(index, end):
                active[pos] = True
        index = end

    runs, start = [], None
    for y, value in enumerate(active + [False]):
        if value and start is None:
            start = y
        elif not value and start is not None:
            runs.append((start, y))
            start = None
    return runs


def is_a_plus_contact_sheet(src: Path):
    """识别 1464px 电脑列 + 留白 + 1200px 手机列的高级 A+ 总图。"""
    from PIL import Image
    with Image.open(src) as image:
        return image.width >= 1464 + 1200 and image.height >= 900


_CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def _aplus_name_key(name):
    """按“第N屏M”的业务顺序排序参考文件名，而不是按字典序。

    例：电脑第一屏1 / 电脑第一屏2 / 第三屏（轮播）1 / 第三屏（轮播）2 /
    第四屏（轮播）1 / 第四屏（轮播）2 / 第六屏(轮播）1 / 第六屏(轮播）2 /
    第七屏 → 顺序 1,1 / 1,2 / 3,1 / 3,2 / 4,1 / 4,2 / 6,1 / 6,2 / 7,0
    """
    import re
    stem = Path(name).stem
    screen = None
    for ch in stem:
        if ch in _CN_NUM:
            screen = _CN_NUM[ch]
            break
    if screen is None:
        m = re.search(r"(\d+)", stem)
        screen = int(m.group(1)) if m else 99
    nums = re.findall(r"(\d+)", stem)
    sub = int(nums[-1]) if nums else 0
    if len(nums) == 1 and screen == int(nums[0]):
        sub = 0
    return (screen, sub, stem)


def reference_names(reference_root, subdir, expected_height):
    """读取参考文件夹中某一列的文件名，按业务顺序排序。"""
    from PIL import Image
    d = Path(reference_root) / subdir
    if not d.is_dir():
        return []
    items = []
    for p in d.iterdir():
        if not p.is_file() or p.suffix.lower() not in IMAGE_EXTS:
            continue
        try:
            with Image.open(p) as im:
                if im.height == expected_height:
                    items.append(p.name)
        except Exception:
            continue
    items.sort(key=_aplus_name_key)
    return items


def split_a_plus_contact_sheet(src: Path, output_root=None, quality=95,
                                reference_root=None):
    """从高级 A+ 双列总图提取电脑/手机列，并按亚马逊规格连续切图。

    reference_root 不为空时，切出的文件沿用参考文件夹里的中文业务名
    （如“第三屏（轮播）1.jpg”），而不是“电脑第一屏N.jpg”序号名。
    """
    src = Path(src)
    image = _open_rgb(src)
    try:
        if image.width < 1464 + 1200:
            raise ValueError("不是电脑端+手机端双列总图")
        root = Path(output_root) if output_root else src.parent / src.stem
        specs = (
            ("电脑端", 0, 1464, 600, "电脑第一屏"),
            ("手机端", image.width - 1200, image.width, 900, "手机第一屏"),
        )
        outputs = []
        for folder_name, left, right, tile_height, prefix in specs:
            # 参考模式：沿用旧版中文业务名；数量不足时回退到序号名。
            names = (reference_names(reference_root, folder_name, tile_height)
                     if reference_root else [])
            column = image.crop((left, 0, right, image.height))
            try:
                runs = _content_row_runs(column)
                if not runs:
                    raise ValueError(f"{folder_name}列没有识别到有效画面")
                number = 1
                for top, bottom in runs:
                    run_height = bottom - top
                    # 每段内容对应若干标准屏。连屏图里每屏四周常留白，内容
                    # 区段高度会略小于屏高（例如内容 561px、屏高 600px），
                    # 因此按最接近的屏数吸附，而不是要求区段≥完整屏高。
                    screen_count = max(1, int(round(run_height / tile_height)))
                    expected = screen_count * tile_height
                    # 只排除明显是小图标/杂色的碎片区段：不足半屏直接忽略。
                    if run_height < tile_height * 0.5:
                        continue
                    # 允许的偏差放宽到半屏，覆盖四周留白造成的高度缩水。
                    if abs(run_height - expected) > tile_height * 0.5:
                        raise ValueError(
                            f"{folder_name}区段高度 {run_height}px 不能按 {tile_height}px 切分")
                    bottom = min(image.height, top + expected)
                    for y in range(top, bottom, tile_height):
                        tile = column.crop((0, y, right - left, y + tile_height))
                        if number <= len(names):
                            out = root / folder_name / names[number - 1]
                        else:
                            out = root / folder_name / f"{prefix}{number}.jpg"
                        try:
                            _save_jpeg(tile, out, quality)
                        finally:
                            tile.close()
                        outputs.append(out)
                        number += 1
            finally:
                column.close()
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


def read_artboards(src, manifest):
    """只读读取 PSD / PSB 的画板名字与坐标，写入 manifest 文件。

    每行格式：NAME|x0,y0,x1,y1。读取失败返回空列表（调用方回退到参考命名）。
    不改动源文档。
    """
    import pythoncom
    src = Path(src)
    manifest = Path(manifest)
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.unlink(missing_ok=True)
    script = (
        "app.displayDialogs=DialogModes.NO;"
        "var d=null,mf=null;"
        "function cleanName(s){return s.replace(/[\\/:*?\"<>|]/g,'_')"
        ".replace(/^\\s+|\\s+$/g,'');}"
        "function isAB(l){var r=new ActionReference();"
        "r.putIdentifier(charIDToTypeID('Lyr '),l.id);var x=executeActionGet(r),"
        "e=stringIDToTypeID('artboardEnabled');"
        "return x.hasKey(e)&&x.getBoolean(e);}"
        "function rect(l){var r=new ActionReference();"
        "r.putIdentifier(charIDToTypeID('Lyr '),l.id);var x=executeActionGet(r),"
        "e=stringIDToTypeID('artboardEnabled'),k=stringIDToTypeID('artboard');"
        "if(x.hasKey(e)&&x.getBoolean(e)&&x.hasKey(k)){"
        "var a=x.getObjectValue(k),q=a.getObjectValue(stringIDToTypeID('artboardRect'));"
        "return [q.getUnitDoubleValue(stringIDToTypeID('left')),"
        "q.getUnitDoubleValue(stringIDToTypeID('top')),"
        "q.getUnitDoubleValue(stringIDToTypeID('right')),"
        "q.getUnitDoubleValue(stringIDToTypeID('bottom'))];}"
        "if(l.typename=='LayerSet'){var b=l.bounds;return [b[0].as('px'),b[1].as('px'),"
        "b[2].as('px'),b[3].as('px')];}return null;}"
        "try{d=app.open(new File('" + js_escape(str(src)) + "'));"
        "mf=new File('" + js_escape(str(manifest)) + "');mf.encoding='UTF8';mf.open('w');"
        "for(var i=0;i<d.layers.length;i++){var l=d.layers[i];"
        "if(!isAB(l))continue;var r=rect(l);if(!r)continue;"
        "mf.writeln('BOARD:'+cleanName(l.name)+'|'+Math.round(r[0])+','+"
        "Math.round(r[1])+','+Math.round(r[2])+','+Math.round(r[3]));}"
        "}catch(err){if(mf)mf.writeln('ERROR:'+err.toString());}"
        "finally{if(mf)mf.close();if(d)d.close(SaveOptions.DONOTSAVECHANGES);}"
    )
    pythoncom.CoInitialize()
    app = None
    try:
        app = photoshop()
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
        pythoncom.CoUninitialize()

    boards = []
    if manifest.exists():
        for line in manifest.read_text(encoding="utf-8", errors="replace").splitlines():
            if not line.startswith("BOARD:"):
                continue
            body = line[len("BOARD:"):]
            name, _, coords = body.partition("|")
            try:
                x0, y0, x1, y1 = [int(v.strip()) for v in coords.split(",")]
            except ValueError:
                continue
            if name and x1 > x0 and y1 > y0:
                boards.append((name, x0, y0, x1, y1))
    return boards


def _safe_artboard_name(raw):
    """把画板名清洗成安全的 Windows 文件名：去非法字符、收敛空白、去首尾点空格。"""
    name = str(raw).strip()
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "_")
    name = re.sub(r"\s+", " ", name).rstrip(". ")
    return name or "未命名画板"


def split_a_plus_by_artboards(src, boards, output_root=None, quality=95):
    """按画板坐标与名字切图。画板高于屏高时按屏高均分并加 1/2 序号。

    健壮性处理：
      - 列分界按画板 x 坐标“最大间隙”自动判定，不写死 800px；
      - 画板名统一清洗（去非法字符/首尾空格/重复空白）并做同名去重；
      - 高度非整数倍时整屏数 + 余数超过半屏补一屏，避免 round 在 .5 处漏切。
    返回输出文件列表；画板列表无效时返回空列表，由调用方回退。
    """
    from PIL import Image
    if not boards:
        return []

    # 按 x0 排序后找最大相邻间隙作为电脑端 / 手机端的列分界。
    xs = sorted({b[1] for b in boards})
    if len(xs) < 2:
        return []
    gaps = [(xs[i + 1] - xs[i], (xs[i] + xs[i + 1]) / 2.0)
            for i in range(len(xs) - 1)]
    gap, split_x = max(gaps, key=lambda item: item[0])
    if gap < 300:  # 未形成明显两列，交给回退逻辑按参考命名处理
        return []
    pc_boards = [b for b in boards if b[1] <= split_x]
    mb_boards = [b for b in boards if b[1] > split_x]
    if not pc_boards or not mb_boards:
        return []

    image = _open_rgb(src)
    outputs = []
    try:
        for folder, items, tile_h in (
            ("电脑端", pc_boards, 600), ("手机端", mb_boards, 900)):
            # 从上到下按 y 排序，与人工浏览顺序一致。
            items = sorted(items, key=lambda b: b[2])
            used = {}
            for name, x0, y0, x1, y1 in items:
                x1 = min(x1, image.width)
                y1 = min(y1, image.height)
                height = y1 - y0
                if height <= 0:
                    continue
                # 整屏数 + 余数超过半屏补一屏：替代 round，
                # 避免 1500/600=2.5 被 round 成 2 屏而漏掉半屏内容。
                screens = max(1, int(math.floor(height / tile_h)))
                if height - screens * tile_h >= tile_h * 0.5:
                    screens += 1
                for i in range(screens):
                    top = y0 + i * tile_h
                    bottom = min(y1, top + tile_h)
                    if bottom - top < tile_h * 0.5:
                        continue
                    tile = image.crop((x0, top, x1, bottom))
                    label = _safe_artboard_name(name)
                    # 第一屏画板拆两屏时，沿用“电脑第一屏1/2”“手机第一屏1/2”。
                    if "第一屏" in label and screens > 1:
                        prefix = "电脑" if folder == "电脑端" else "手机"
                        label = "%s第一屏%d" % (prefix, i + 1)
                    elif screens > 1:
                        label = "%s%d" % (label, i + 1)
                    # 同名去重：同一文件夹内不覆盖（如两列都叫“充电”时分列已隔开；
                    # 同列真重名则追加 _2、_3）。
                    base = label
                    if base in used:
                        used[base] += 1
                        label = "%s_%d" % (base, used[base])
                    else:
                        used[base] = 1
                    out = Path(output_root) / folder / (label + ".jpg")
                    out.parent.mkdir(parents=True, exist_ok=True)
                    try:
                        _save_jpeg(tile, out, quality)
                    finally:
                        tile.close()
                    outputs.append(out)
    finally:
        image.close()
    return outputs


def render_psd_to_jpg(src, jpg_path, progress_cb=None):
    """用 Photoshop COM 把 PSD / PSB 渲染成整张连屏 JPG。

    PIL 只能读 PSD，读不了 PSB（Photoshop 大型文档）；高级 A+ 的切图
    逻辑基于 PIL，所以遇到 PSD / PSB 输入时先借 Photoshop 渲染成一张
    完整合成 JPG（不裁切、不缩放、最高质量），再交给切图函数处理。
    只写临时 JPG，绝不改动源 PSD / PSB。
    """
    import pythoncom
    src = Path(src)
    jpg_path = Path(jpg_path)
    pythoncom.CoInitialize()
    app = None
    try:
        app = photoshop()
        script = build_script(
            str(src), None, str(jpg_path), "none", ("jpg",),
            progress_file=None, layer_list=None, write_thumbnail=False)
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
        if not jpg_path.exists() or jpg_path.stat().st_size == 0:
            raise RuntimeError("Photoshop 未生成渲染图，请确认 Photoshop 可正常打开文件")
        return jpg_path
    finally:
        try:
            del app
        except Exception:
            pass
        pythoncom.CoUninitialize()


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


def convert_artboards(psd: Path, output_dir: Path, progress_cb=None,
                      cancel_event=None, info_cb=None) -> list:
    """把 PSD/PSB 顶层画板逐个导出成以画板名命名的最高质量 JPG。"""
    import pythoncom

    pythoncom.CoInitialize()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    token = f"{os.getpid()}_{threading.get_ident()}"
    marker = Path(tempfile.gettempdir()) / f"psdtool_progress_{token}.txt"
    manifest = Path(tempfile.gettempdir()) / f"psdtool_artboards_{token}.txt"
    cancel_marker = Path(tempfile.gettempdir()) / f"psdtool_cancel_{token}.txt"
    stop_monitor = threading.Event()
    monitor = None
    emitted_count = [0]
    try:
        marker.unlink(missing_ok=True)
        manifest.unlink(missing_ok=True)
        cancel_marker.unlink(missing_ok=True)
        if progress_cb:
            progress_cb("connecting")

            def watch_progress():
                while not stop_monitor.wait(0.2):
                    if cancel_event and cancel_event.is_set() and not cancel_marker.exists():
                        try:
                            cancel_marker.write_text("stop", encoding="ascii")
                        except OSError:
                            pass
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
        script = build_artboard_script(
            str(psd), str(output_dir), str(manifest),
            marker if progress_cb else None, cancel_marker)
        try:
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
            del app
        try:
            lines = [line for line in
                     manifest.read_text(encoding="utf-8-sig").splitlines()
                     if line.strip()]
        except (OSError, UnicodeError):
            lines = []
        errors = [line[6:] for line in lines if line.startswith("ERROR:")]
        if errors:
            raise RuntimeError("画板导出失败：" + errors[-1])
        notices = [line for line in lines
                   if line.startswith(("META:", "WARN:", "BOARD_ERROR:"))]
        if info_cb:
            for notice in notices:
                info_cb(notice)
        outputs = [Path(line[7:]) for line in lines if line.startswith("OUTPUT:")]
        if "CANCELLED" in lines:
            raise ConversionCancelled(f"已安全停止；当前文件已导出 {len(outputs)} 张")
        board_errors = [line[12:] for line in lines if line.startswith("BOARD_ERROR:")]
        if not outputs and board_errors:
            raise RuntimeError("所有画板导出失败：" + board_errors[-1])
        if not outputs:
            raise RuntimeError("没有检测到可导出的 Photoshop 画板")
        if progress_cb:
            progress_cb("optimizing_jpg")
        # Photoshop 的质量 12 在高细节渲染图上可能产生 10~20 MB 的 JPG。
        # 用 4:4:4、质量 99 的优化编码保留肉眼质量，同时匹配电商交付体积。
        from PIL import Image
        for exported in outputs:
            if cancel_event and cancel_event.is_set():
                raise ConversionCancelled(
                    f"已安全停止；当前文件已导出 {len(outputs)} 张")
            temp_jpg = exported.with_name(exported.stem + ".opt.tmp.jpg")
            try:
                with Image.open(exported) as image:
                    info = image.info.copy()
                    save_kwargs = {
                        "quality": 99, "optimize": True, "subsampling": 0,
                    }
                    if info.get("icc_profile"):
                        save_kwargs["icc_profile"] = info["icc_profile"]
                    if info.get("dpi"):
                        save_kwargs["dpi"] = info["dpi"]
                    image.convert("RGB").save(temp_jpg, "JPEG", **save_kwargs)
                os.replace(temp_jpg, exported)
            except Exception as exc:
                try:
                    temp_jpg.unlink(missing_ok=True)
                except OSError:
                    pass
                if info_cb:
                    info_cb(f"WARN:{exported.name} 体积优化失败，已保留原图：{exc}")
        return outputs
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
        for temp_path in (marker, manifest, cancel_marker):
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# 自检（不启动 Photoshop，不读取用户素材内容）
# ---------------------------------------------------------------------------
def selftest() -> int:
    import tempfile
    global SETTINGS_PATH

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
    check("导出修改不写回原稿", "DONOTSAVECHANGES" in js_all)
    js_psd_thumbnail = build_script(
        r"F:\主图.psd", None, r"F:\out\主图.jpg", "none", ("jpg",))
    check("PSD 自动写入资源管理器兼容缩略图",
          "maximizeCompatibility" in js_psd_thumbnail
          and "d.save()" in js_psd_thumbnail)
    check("PSB 不因缩略图整稿重存",
          "maximizeCompatibility" not in build_script(
              r"F:\大文件.psb", None, r"F:\out\x.jpg", "none", ("jpg",)))
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
        folder_outputs = unique_outputs(src, ("jpg",), use_folder=True)
        check("主副图画板模式进入同名文件夹",
              folder_outputs["jpg"].parent.name == "NTT72-D")
        o3 = unique_outputs(src, ("jpg",))
        check("渲染模式只选 JPG 时不生成图层文本",
              set(o3) == {"jpg"})

    js_layers = build_script(tricky, None, r"F:\out\x.jpg", "none", ("jpg",),
                             None, r"F:\out\图层列表.txt")
    check("递归读取图层与分组名称",
          "function dl" in js_layers and "LayerSet" in js_layers and
          "图层列表.txt" in js_layers)

    js_artboards = build_artboard_script(
        r"F:\主副图.psd", r"F:\主副图", r"F:\画板清单.txt")
    check("画板模式按画板名逐张导出 JPG",
          "artboardRect" in js_artboards and "d.layers[i].name" in js_artboards
          and "JPEGSaveOptions" in js_artboards)
    check("画板模式包含预检、逐张进度与安全停止",
          "preflight_done" in js_artboards and "pg('board:" in js_artboards
          and "CANCELLED" in js_artboards)
    check("原生画板先拼合再按画板坐标裁切",
          "d.flatten();d.crop(rect);" in js_artboards)
    check("重名画板使用短横线编号",
          "/'+n+'-'+z+'.jpg'" in js_artboards)
    check("主副图模式将全部原生画板合成为整体效果图",
          "overviewRect" in js_artboards and "(整体|总览)" in js_artboards
          and "overview?isNativeArtboard" in js_artboards)

    # --- v2.8：三种同名文件策略、PNG/JPG 共用后缀 ---
    with tempfile.TemporaryDirectory(prefix="psdtool_conflict_") as td:
        d = Path(td)
        src = d / "NTT72-D.psb"
        src.write_bytes(b"")
        # 自动编号：首次 + 冲突后成对使用同一后缀
        n1 = unique_outputs(src, ("png", "jpg"), conflict="number")
        for p in n1.values():
            p.write_bytes(b"")
        n2 = unique_outputs(src, ("png", "jpg"), conflict="number")
        check("自动编号：冲突后 PNG/JPG 共用同一后缀",
              n2["png"].stem == n2["jpg"].stem and
              n2["png"].name == "NTT72-D_PNG.png" and
              n2["jpg"].name == "NTT72-D_PNG.jpg",
              f"{n2['png'].name} / {n2['jpg'].name}")
        # 覆盖：使用原目标路径、不生成后缀
        ov = unique_outputs(src, ("png", "jpg"), conflict="overwrite")
        check("覆盖策略：使用原路径不生成后缀",
              ov["png"].name == "NTT72-D.png" and ov["jpg"].name == "NTT72-D.jpg",
              f"{ov['png'].name} / {ov['jpg'].name}")
        # 跳过：目标已存在时返回 None
        sk = unique_outputs(src, ("png", "jpg"), conflict="skip")
        check("跳过策略：目标已存在返回 None", sk is None, str(sk))
        # 不覆盖 PSD/PSB：所有输出后缀都不是源文件后缀
        check("输出不会覆盖 PSD/PSB 原稿",
              all(p.suffix.lower() not in SUPPORTED_EXTS for p in ov.values()))

    # --- v2.8：设置默认值 / 损坏回退 / 非法值忽略 ---
    saved_path = SETTINGS_PATH
    try:
        with tempfile.TemporaryDirectory(prefix="psdtool_settings_") as td:
            SETTINGS_PATH = Path(td) / "PSDTool" / "settings.json"
            fresh = load_settings()
            check("设置文件不存在时使用安全默认值",
                  fresh == DEFAULT_SETTINGS, str(fresh))
            SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
            SETTINGS_PATH.write_text("{ this is not json", encoding="utf-8")
            broken = load_settings()
            check("设置 JSON 损坏时安全回退默认值",
                  broken == DEFAULT_SETTINGS, str(broken))
            # 合法值保留，非法值被忽略
            good = normalize_settings({
                "mode": "artboard", "trim": "both", "png": True, "jpg": False,
                "conflict": "overwrite", "aplus_quality": 90,
                "aplus_reference": r"D:\参考", "secret": "ignore me"})
            check("合法设置值被保留",
                  good["mode"] == "artboard" and good["trim"] == "both" and
                  good["conflict"] == "overwrite" and good["aplus_quality"] == 90 and
                  good["png"] is True and good["jpg"] is False, str(good))
            bad = normalize_settings({
                "mode": "unknown", "trim": "xxx", "conflict": "delete",
                "png": "yes", "jpg": 1, "aplus_quality": 999,
                "aplus_reference": 123})
            check("无效设置值被忽略并回退默认", bad == DEFAULT_SETTINGS, str(bad))
            check("设置不保存未知键/隐私", "secret" not in good and
                  set(good) == set(DEFAULT_SETTINGS))
    finally:
        SETTINGS_PATH = saved_path
    # save_settings 往返（写到临时目录，不污染真实 APPDATA）
    saved_path = SETTINGS_PATH
    try:
        with tempfile.TemporaryDirectory(prefix="psdtool_save_") as td:
            SETTINGS_PATH = Path(td) / "PSDTool" / "settings.json"
            save_settings({"mode": "aplus", "conflict": "skip", "aplus_quality": 88})
            loaded = load_settings()
            check("保存设置后可正确读回",
                  loaded["mode"] == "aplus" and loaded["conflict"] == "skip" and
                  loaded["aplus_quality"] == 88, str(loaded))
    finally:
        SETTINGS_PATH = saved_path

    # --- v2.8：高级 A+ 固定切图、首页映射、排除源文件本身 ---
    try:
        from PIL import Image as _Image
        with tempfile.TemporaryDirectory(prefix="psdtool_aplus_") as td:
            d = Path(td)
            pc = d / "电脑首页.jpg"
            _Image.new("RGB", (1464, 1200), (200, 100, 50)).save(pc, "JPEG")
            pc_out = split_a_plus_fixed(pc, quality=90)
            pc_sizes = []
            for p in pc_out:
                with _Image.open(p) as im:
                    pc_sizes.append(im.size)
            check("高级A+固定电脑端切图（1464×1200→2×1464×600）",
                  len(pc_out) == 2 and pc_sizes == [(1464, 600), (1464, 600)],
                  str(pc_sizes))

            mb = d / "手机首页.jpg"
            _Image.new("RGB", (1200, 2700), (50, 100, 200)).save(mb, "JPEG")
            mb_out = split_a_plus_fixed(mb, quality=90)
            mb_sizes = []
            for p in mb_out:
                with _Image.open(p) as im:
                    mb_sizes.append(im.size)
            check("高级A+固定手机端切图（1200×2700→3×1200×900）",
                  len(mb_out) == 3 and all(s == (1200, 900) for s in mb_sizes),
                  str(mb_sizes))

            bad_h = d / "bad.jpg"
            _Image.new("RGB", (1464, 601), (0, 0, 0)).save(bad_h, "JPEG")
            try:
                split_a_plus_fixed(bad_h)
                raised = False
            except ValueError:
                raised = True
            check("高度不能整除时明确报错", raised)

            # 参考目录：电脑端/电脑第一屏1.jpg(600) + 电脑第一屏2.jpg(600)
            ref = d / "ref"
            (ref / "电脑端").mkdir(parents=True)
            _Image.new("RGB", (1464, 600), (1, 2, 3)).save(
                ref / "电脑端" / "电脑第一屏1.jpg", "JPEG")
            _Image.new("RGB", (1464, 600), (4, 5, 6)).save(
                ref / "电脑端" / "电脑第一屏2.jpg", "JPEG")
            home = d / "A+首页.jpg"
            _Image.new("RGB", (1464, 1200), (7, 8, 9)).save(home, "JPEG")
            groups = _reference_module_groups(ref, home, 1464, 1200)
            mapped = [p.name for p, _ in
                      (groups[0][1] if groups else [])]
            check("首页映射到第一屏1/2",
                  sorted(mapped) == ["电脑第一屏1.jpg", "电脑第一屏2.jpg"],
                  str(mapped))
            # 扫描参考文件夹时必须排除源连屏图本身
            (ref / "电脑端" / "A+首页.jpg").write_bytes(pc.read_bytes())
            groups2 = _reference_module_groups(ref, ref / "电脑端" / "A+首页.jpg",
                                               1464, 1200)
            included = {p.resolve() for _, items in groups2 for p, _ in items}
            check("参考扫描排除源文件本身",
                  (ref / "电脑端" / "A+首页.jpg").resolve() not in included,
                  str(len(included)))
    except ImportError:
        check("高级 A+ 图片切图需要 Pillow", False, "Pillow 不可用")

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

        # Give Windows a stable identity so the taskbar uses our icon instead
        # of grouping the window under Python/Tk.
        if os.name == "nt":
            try:
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    "PSDStudio.PSDTool.3")
            except Exception:
                pass

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
        # 启动时读取用户设置（文件不存在 / JSON 损坏 / 值非法都会安全回退）
        self.settings = load_settings()
        self.output_dirs = []
        self.aplus_quality = self.settings["aplus_quality"]
        self.aplus_reference = self.settings["aplus_reference"]

        self.root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.report_callback_exception = self.report_callback_exception
        self.root.title(APP_TITLE)
        try:
            self.root.iconbitmap(default=str(resource_path("assets/app_icon.ico")))
        except Exception:
            pass
        try:
            from PIL import ImageTk
            self.taskbar_icon = ImageTk.PhotoImage(
                Image.open(resource_path("assets/app_icon.png")))
            self.root.iconphoto(True, self.taskbar_icon)
        except Exception:
            pass
        self.root.geometry("760x720")
        self.root.minsize(700, 620)
        self.root.configure(bg=C_BG)

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

        style = ttk.Style()
        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#e5e7eb",
                        background=C_PRIMARY, borderwidth=0)
        style.configure("TCombobox", fieldbackground=C_CARD,
                        background=C_CARD, borderwidth=1)

        F = ("Microsoft YaHei UI", 9)
        FB = ("Microsoft YaHei UI", 9, "bold")

        # ---------- 品牌头部 ----------
        header = tk.Frame(self.root, bg=C_HEADER, height=88)
        header.pack(fill="x")
        header.pack_propagate(False)
        badge = tk.Frame(header, bg=C_HEADER)
        badge.pack(side="left", padx=(22, 0))
        tk.Label(badge, text="PS", font=("Segoe UI", 20, "bold"),
                 bg=C_PRIMARY, fg="white", width=3, pady=6).pack()
        ht = tk.Frame(header, bg=C_HEADER)
        ht.pack(side="left", padx=(14, 0), pady=17)
        tk.Label(ht, text="PSD Studio", font=("Segoe UI", 16, "bold"),
                 bg=C_HEADER, fg="#f9fafb").pack(anchor="w")
        tk.Label(ht, text="批量导图工具 · 让 Photoshop 精确导出，原稿零改动",
                 font=F, bg=C_HEADER, fg="#9ca3af").pack(anchor="w", pady=(3, 0))
        tk.Label(header, text=f"v{VERSION}", font=("Segoe UI", 10, "bold"),
                 bg="#1f2937", fg="#93c5fd", padx=10, pady=3).pack(
                     side="right", padx=20)

        body = tk.Frame(self.root, bg=C_BG)
        body.pack(fill="both", expand=True, padx=18, pady=(14, 12))

        # ---------- 工作模式卡片（带序号徽章）----------
        mode_row = tk.Frame(body, bg=C_BG)
        mode_row.pack(fill="x", pady=(0, 12))
        self.output_mode = tk.StringVar(
            value=MODE_VALUES.get(self.settings["mode"], MODE_VALUES["render"]))
        modes = tuple((key, heading, sub, value)
                      for key, heading, sub, value in MODE_DEFS)
        self.mode_cards = []
        for idx, (key, heading, sub, value) in enumerate(modes):
            accent = C_PURPLE if key == "aplus" else C_PRIMARY
            card = tk.Frame(mode_row, bg=C_CARD, highlightthickness=1,
                            highlightbackground=C_BORDER, bd=0)
            card.pack(side="left", fill="x", expand=True,
                      padx=(0, 10 if idx < len(modes) - 1 else 0))
            num = tk.Label(card, text=str(idx + 1), font=("Segoe UI", 9, "bold"),
                           bg=accent, fg="white", width=2)
            num.pack(anchor="w", padx=(10, 0), pady=(8, 0))
            rb = tk.Radiobutton(
                card, text=f"{heading}\n{sub}", variable=self.output_mode,
                value=value, indicatoron=False, command=self.sync_mode_ui,
                font=FB, justify="left", anchor="w", padx=12, pady=10,
                bd=0, relief="flat", bg=C_CARD, fg=C_INK, selectcolor=C_CARD,
                activebackground=C_SOFT_BLUE, activeforeground=C_PRIMARY,
                highlightthickness=0)
            rb.pack(fill="both", expand=True)
            self.mode_cards.append((card, num, rb, value, accent))

        # ---------- 拖放区 ----------
        self.drop = tk.Label(
            body,
            text="↓\n把 PSD / PSB 文件或文件夹拖到这里\n"
                 "默认导出到源文件旁边，不修改原稿",
            font=("Microsoft YaHei UI", 12), bg=C_CARD, fg=C_PRIMARY,
            relief="flat", bd=0, justify="center", padx=16, pady=22)
        self.drop.pack(fill="both", expand=True, pady=(0, 12))
        self.drop.configure(highlightthickness=2, highlightbackground=C_BLUE_LINE)

        # ---------- 设置卡片：裁切 / 格式 / 同名 ----------
        settings = tk.Frame(body, bg=C_CARD, highlightthickness=1,
                            highlightbackground=C_BORDER)
        settings.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(settings, bg=C_CARD)
        inner.pack(fill="x", padx=16, pady=(12, 10))

        row1 = tk.Frame(inner, bg=C_CARD)
        row1.pack(fill="x", pady=(0, 8))
        tk.Label(row1, text="裁切方式", width=9, anchor="w",
                 font=FB, bg=C_CARD, fg=C_SUB).pack(side="left")
        self.trim = tk.StringVar(
            value=TRIM_MODES.get(self.settings["trim"], TRIM_MODES["transparent"]))
        combo = ttk.Combobox(row1, textvariable=self.trim, state="readonly",
                             width=38, font=F)
        combo["values"] = [TRIM_MODES[k] for k in TRIM_MODES]
        if self.trim.get() in combo["values"]:
            combo.current(list(combo["values"]).index(self.trim.get()))
        else:
            combo.current(0)
        combo.pack(side="left", padx=4)

        row2 = tk.Frame(inner, bg=C_CARD)
        row2.pack(fill="x", pady=(0, 8))
        tk.Label(row2, text="导出格式", width=9, anchor="w",
                 font=FB, bg=C_CARD, fg=C_SUB).pack(side="left")
        self.var_png = tk.BooleanVar(value=self.settings["png"])
        self.var_jpg = tk.BooleanVar(value=self.settings["jpg"])
        tk.Checkbutton(row2, text="PNG（保留透明）", variable=self.var_png,
                       font=F, bg=C_CARD, fg=C_INK, activebackground=C_CARD,
                       selectcolor=C_SOFT_BLUE, relief="flat").pack(side="left", padx=(4, 12))
        tk.Checkbutton(row2, text="JPG（最高质量 12/12）", variable=self.var_jpg,
                       font=F, bg=C_CARD, fg=C_INK, activebackground=C_CARD,
                       selectcolor=C_SOFT_BLUE, relief="flat").pack(side="left")
        tk.Button(row2, text="快速预设", command=self.fast_preset,
                  font=("Microsoft YaHei UI", 8), bg=C_CARD, fg=C_SUB,
                  bd=1, relief="solid", padx=10, pady=2,
                  activebackground="#f9fafb").pack(side="right")

        # 高级 A+ 参考模式固定覆盖，不受本行影响。
        row3 = tk.Frame(inner, bg=C_CARD)
        row3.pack(fill="x")
        self.conflict_row = row3
        tk.Label(row3, text="同名文件", width=9, anchor="w",
                 font=FB, bg=C_CARD, fg=C_SUB).pack(side="left")
        self.conflict = tk.StringVar(value=self.settings["conflict"])
        for text, key in (("自动编号", "number"), ("覆盖", "overwrite"),
                          ("跳过", "skip")):
            tk.Radiobutton(row3, text=text, variable=self.conflict, value=key,
                           font=F, bg=C_CARD, fg=C_INK, activebackground=C_CARD,
                           selectcolor=C_SOFT_BLUE, relief="flat",
                           activeforeground=C_PRIMARY).pack(side="left", padx=(4, 10))

        # ---------- 操作按钮 ----------
        btns = tk.Frame(body, bg=C_BG)
        btns.pack(fill="x", pady=(0, 12))
        light_btn = dict(font=F, bg=C_CARD, fg=C_INK, bd=1, relief="solid",
                         highlightthickness=0, padx=8, pady=4,
                         activebackground="#f9fafb")
        self.btn_files = tk.Button(btns, text="选择文件（可多选）", width=16,
                                   command=self.pick_files, **light_btn)
        self.btn_folder = tk.Button(btns, text="选择文件夹（递归）", width=16,
                                    command=self.pick_folder, **light_btn)
        self.btn_go = tk.Button(btns, text="开始转换", width=12,
                                command=self.start, bg=C_PRIMARY, fg="white",
                                font=FB, bd=0, padx=10, pady=6,
                                activebackground=C_PRIMARY_D, activeforeground="white")
        self.btn_stop = tk.Button(btns, text="停止队列", width=9,
                                  command=self.request_stop, state="disabled",
                                  **light_btn)
        self.btn_clear = tk.Button(btns, text="清空", width=6,
                                   command=self.clear, **light_btn)
        self.btn_open_dir = tk.Button(btns, text="打开输出目录", width=13,
                                      command=self.open_output_dir, state="disabled",
                                      **light_btn)
        for b in (self.btn_files, self.btn_folder, self.btn_go,
                  self.btn_stop, self.btn_clear):
            b.pack(side="left", padx=(0, 8))
        self.btn_open_dir.pack(side="right")
        self.sync_mode_ui()

        # ---------- 进度与状态（卡片）----------
        progress_card = tk.Frame(body, bg=C_CARD, highlightthickness=1,
                                 highlightbackground=C_BORDER)
        progress_card.pack(fill="x", pady=(0, 10))
        pc = tk.Frame(progress_card, bg=C_CARD)
        pc.pack(fill="x", padx=16, pady=(10, 4))
        tk.Label(pc, text="总进度", width=9, anchor="w",
                 font=F, bg=C_CARD, fg=C_SUB).pack(side="left")
        self.progress = ttk.Progressbar(pc, mode="determinate",
                                        style="Accent.Horizontal.TProgressbar")
        self.progress.pack(side="left", fill="x", expand=True)

        current_row = tk.Frame(progress_card, bg=C_CARD)
        current_row.pack(fill="x", padx=16, pady=(4, 4))
        self.current_progress_label = tk.StringVar(value="当前文件：空闲")
        tk.Label(current_row, textvariable=self.current_progress_label,
                 width=22, anchor="w", font=F, bg=C_CARD,
                 fg=C_SUB).pack(side="left")
        # 空闲用确定值 0，避免某些 Windows 主题在停止的动态进度条上残留绿块。
        self.current_progress = ttk.Progressbar(
            current_row, mode="determinate", maximum=100, value=0,
            style="Accent.Horizontal.TProgressbar")
        self.current_progress.pack(side="left", fill="x", expand=True)

        self.status = tk.StringVar(value="等待文件")
        tk.Label(progress_card, textvariable=self.status, anchor="w",
                 font=F, bg=C_CARD, fg=C_PRIMARY).pack(
                     fill="x", padx=16, pady=(4, 10))

        # ---------- 日志（卡片 + 标题栏）----------
        log_card = tk.Frame(body, bg=C_CARD, highlightthickness=1,
                            highlightbackground=C_BORDER)
        log_card.pack(fill="both", expand=True)
        log_head = tk.Frame(log_card, bg=C_CARD)
        log_head.pack(fill="x", padx=16, pady=(9, 0))
        tk.Label(log_head, text="运行日志", font=FB, bg=C_CARD,
                 fg=C_SUB).pack(side="left")
        tk.Label(log_head, text="单张失败不中断整批", font=F, bg=C_CARD,
                 fg=C_MUTED).pack(side="right")
        self.logbox = tk.Text(log_card, height=6, font=("Consolas", 9),
                              bg="#fbfcfe", fg=C_INK, relief="flat", bd=0,
                              highlightthickness=0, padx=10, pady=8)
        self.logbox.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        self.logbox.configure(state="disabled")

    # -- helpers ------------------------------------------------------------
    def sync_mode_ui(self):
        """让主按钮、拖放说明与模式卡片高亮跟随当前业务模式。"""
        mode = self.output_mode.get()
        if not hasattr(self, "btn_go"):
            return
        # 卡片选中态：高亮边框 + 序号徽章配色 + 标题着色。
        if hasattr(self, "mode_cards"):
            for card, num, rb, value, accent in self.mode_cards:
                selected = value == mode
                card.configure(
                    highlightbackground=accent if selected else C_BORDER,
                    highlightthickness=2 if selected else 1,
                    bg=C_SOFT_PURPLE if (selected and accent == C_PURPLE)
                    else (C_SOFT_BLUE if selected else C_CARD))
                num.configure(bg=accent if selected else C_MUTED)
                rb.configure(bg=(C_SOFT_PURPLE if accent == C_PURPLE else C_SOFT_BLUE)
                             if selected else C_CARD,
                             fg=accent if selected else C_INK)
        # 高级 A+ 不经过普通转换的同名策略，置灰避免误导。
        if hasattr(self, "conflict_row"):
            conflict_state = "disabled" if mode.startswith("高级 A+") else "normal"
            for child in self.conflict_row.winfo_children():
                try:
                    child.configure(state=conflict_state)
                except Exception:
                    pass
        if mode.startswith("高级 A+"):
            self.btn_go.configure(text="打开高级 A+", bg=C_PURPLE,
                                  activebackground=C_PURPLE_D)
            self.drop.configure(
                text="高级 A+ 连屏导图\n点击“打开高级 A+”进入专用工作区\n"
                     "支持电脑端、手机端和旧版文件夹匹配",
                fg=C_PURPLE, highlightbackground=C_PURPLE_LINE)
        elif mode.startswith("主副图画板"):
            self.btn_go.configure(text="开始转换", bg=C_PRIMARY,
                                  activebackground=C_PRIMARY_D)
            self.drop.configure(
                text="↓\n把主副图 PSD / PSB 拖到这里\n"
                     "自动建立源文件同名文件夹",
                fg=C_PRIMARY, highlightbackground=C_BLUE_LINE)
        else:
            self.btn_go.configure(text="开始转换", bg=C_PRIMARY,
                                  activebackground=C_PRIMARY_D)
            self.drop.configure(
                text="↓\n把渲染 PSD / PSB 文件或文件夹拖到这里\n"
                     "JPG 默认保存在源文件旁边",
                fg=C_PRIMARY, highlightbackground=C_BLUE_LINE)

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

    def persist_settings(self):
        """把当前界面状态写入设置文件；只存偏好，不存文件列表/隐私。"""
        if not hasattr(self, "conflict"):
            return
        mode = MODE_KEYS.get(self.output_mode.get(), "render")
        data = {
            "mode": mode,
            "trim": self.trim_key(),
            "png": bool(self.var_png.get()),
            "jpg": bool(self.var_jpg.get()),
            "conflict": self.conflict.get(),
            "aplus_quality": self.aplus_quality,
            "aplus_reference": self.aplus_reference,
        }
        self.settings = data
        save_settings(data)

    def on_close(self):
        self.persist_settings()
        try:
            self.root.destroy()
        except Exception:
            pass

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
        existing = getattr(self, "_aplus_win", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.deiconify()
                    existing.lift()
                    self.log("高级 A+ 工作区已打开，请直接在其中拖入文件")
                    return
            except Exception:
                pass
        tk = self.tk
        F = ("Microsoft YaHei UI", 9)
        FB = ("Microsoft YaHei UI", 9, "bold")
        win = tk.Toplevel(self.root)
        self._aplus_win = win
        win.title("亚马逊高级 A+ 连屏导图")
        win.geometry("820x680")
        win.minsize(720, 600)
        win.configure(bg=C_BG)
        win.transient(self.root)
        win.protocol("WM_DELETE_WINDOW", lambda: self._on_aplus_close(win))

        # 品牌头（紫），与主界面呼应
        ap_header = tk.Frame(win, bg=C_PURPLE, height=72)
        ap_header.pack(fill="x")
        ap_header.pack_propagate(False)
        apb = tk.Frame(ap_header, bg=C_PURPLE)
        apb.pack(side="left", padx=(18, 0))
        tk.Label(apb, text="A+", font=("Segoe UI", 16, "bold"),
                 bg=C_PURPLE_D, fg="white", width=3, pady=6).pack()
        apt = tk.Frame(ap_header, bg=C_PURPLE)
        apt.pack(side="left", padx=(12, 0), pady=12)
        tk.Label(apt, text="亚马逊高级 A+ 连屏导图",
                 font=("Microsoft YaHei UI", 14, "bold"),
                 bg=C_PURPLE, fg="white").pack(anchor="w")
        tk.Label(apt, text="按旧版文件名和尺寸原位覆盖，或按 1464×600 / 1200×900 自动切图",
                 font=("Microsoft YaHei UI", 9), bg=C_PURPLE,
                 fg="#ede9fe").pack(anchor="w", pady=(3, 0))

        # 把主界面已经拖入的连屏图直接带入高级 A+ 列表，避免用户点了
        # “高级 A+ 开始转换”后还得重新选一遍（表现为按钮“没反应”）。
        aplus_exts = IMAGE_EXTS | SUPPORTED_EXTS
        files = [Path(p) for p in getattr(self, "files", [])
                 if Path(p).is_file() and Path(p).suffix.lower() in aplus_exts]
        listbox = tk.Listbox(win, height=6)

        def refresh():
            listbox.delete(0, "end")
            for path in files:
                listbox.insert("end", str(path))

        # 高级 A+ 的输入既可以是已导出的图片，也可以直接是 PSD / PSB 连屏源图。
        # PSB 会先用 Photoshop 渲染成临时 JPG 再切图（见 run_worker）。
        APLUS_INPUT_EXTS = aplus_exts

        def add_images(paths):
            existing = {os.path.normcase(os.path.abspath(str(p))) for p in files}
            added = 0
            for raw in paths:
                path = Path(clean_drop_path(raw))
                if path.is_file() and path.suffix.lower() in APLUS_INPUT_EXTS:
                    key = os.path.normcase(os.path.abspath(str(path)))
                    if key not in existing:
                        files.append(path)
                        existing.add(key)
                        added += 1
            refresh()
            if added:
                # 拖入后立刻给出可见反馈，避免用户以为“没反应”。
                status.set(f"已添加 {added} 个连屏图，共 {len(files)} 个；点「开始转换」执行")
            return added

        def pick_images():
            selected = self.filedialog.askopenfilenames(
                parent=win, title="选择新版高级 A+ 连屏图（图片或 PSD / PSB）",
                filetypes=[("图片 / PSD / PSB",
                            "*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.psd;*.psb")])
            if selected:
                add_images(selected)

        # 暴露给主界面：拖入后可直接补进已打开的工作区，不会重复开窗。
        self._aplus_add_images = add_images
        if self.DND_FILES:
            drop_cb = lambda event: (
                add_images(clean_drop_list(self.root.tk.splitlist(event.data))), "copy")[1]
            # 整个窗口都可拖入，避免用户拖到标题栏/设置区时“没反应”。
            for widget in (win, listbox):
                try:
                    widget.drop_target_register(self.DND_FILES)
                    widget.dnd_bind("<<Drop>>", drop_cb)
                except Exception:
                    pass

        ap_body = tk.Frame(win, bg=C_BG)
        ap_body.pack(fill="both", expand=True, padx=18, pady=(12, 12))

        # 切图设置卡片
        cfg = tk.Frame(ap_body, bg=C_CARD, highlightthickness=1,
                       highlightbackground=C_BORDER)
        cfg.pack(fill="x")
        cfi = tk.Frame(cfg, bg=C_CARD)
        cfi.pack(fill="x", padx=16, pady=(12, 10))

        mode = tk.StringVar(value="reference")
        mode_row = tk.Frame(cfi, bg=C_CARD)
        # 最近一次参考目录作为默认值带入；用户改动后在导图时保存。
        mode_row.pack(fill="x", pady=(0, 9))
        tk.Label(mode_row, text="切图方式", width=9, anchor="w",
                 font=FB, bg=C_CARD, fg=C_SUB).pack(side="left")
        tk.Radiobutton(mode_row, text="参考旧版文件夹（推荐）", variable=mode,
                       value="reference", bg=C_CARD, font=F, fg=C_INK,
                       activebackground=C_CARD, selectcolor=C_SOFT_PURPLE,
                       relief="flat").pack(side="left")
        tk.Radiobutton(mode_row, text="固定尺寸自动切图", variable=mode,
                       value="fixed", bg=C_CARD, font=F, fg=C_INK,
                       activebackground=C_CARD, selectcolor=C_SOFT_PURPLE,
                       relief="flat").pack(side="left", padx=12)

        reference = tk.StringVar(value=self.aplus_reference)
        ref_row = tk.Frame(cfi, bg=C_CARD)
        ref_row.pack(fill="x", pady=(0, 9))
        tk.Label(ref_row, text="参考文件夹", width=9, anchor="w",
                 font=FB, bg=C_CARD, fg=C_SUB).pack(side="left")
        tk.Entry(ref_row, textvariable=reference, font=F, relief="solid",
                 bd=1).pack(side="left", fill="x", expand=True, padx=(0, 6))

        def pick_reference():
            selected = self.filedialog.askdirectory(
                parent=win, title="选择旧版高级 A+ 图片所在文件夹")
            if selected:
                reference.set(selected)

        tk.Button(ref_row, text="选择", command=pick_reference, font=F,
                  bg=C_CARD, fg=C_INK, bd=1, relief="solid",
                  activebackground="#f9fafb", padx=10).pack(side="left")

        quality = tk.IntVar(value=self.aplus_quality)
        option_row = tk.Frame(cfi, bg=C_CARD)
        option_row.pack(fill="x")
        tk.Label(option_row, text="JPG 质量", width=9, anchor="w",
                 font=FB, bg=C_CARD, fg=C_SUB).pack(side="left")
        tk.Spinbox(option_row, from_=80, to=100, textvariable=quality,
                   width=5, font=F, relief="solid", bd=1).pack(side="left")
        tk.Label(option_row, text="从上到下切图 · 不缩放 · 不留缝",
                 font=F, bg=C_CARD, fg=C_MUTED).pack(side="left", padx=12)

        # 图片列表（拖放目标）——白卡片 + 标题，空列表也清晰可见
        list_card = tk.Frame(ap_body, bg=C_CARD, highlightthickness=1,
                             highlightbackground=C_BORDER)
        list_card.pack(in_=ap_body, fill="both", expand=True, pady=(12, 4))
        lh = tk.Frame(list_card, bg=C_CARD)
        lh.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(lh, text="连屏图列表（窗口任意位置可拖入）", font=FB,
                 bg=C_CARD, fg=C_SUB).pack(side="left")
        tk.Label(lh, text="JPG / PNG / PSD / PSB", font=F, bg=C_CARD,
                 fg=C_MUTED).pack(side="right")
        listbox.configure(bg=C_CARD, fg=C_INK, relief="flat",
                          highlightthickness=0, font=("Consolas", 9))
        listbox.pack(in_=list_card, fill="both", expand=True, padx=8, pady=(4, 8))

        status = tk.StringVar(value="请添加新版连屏图")
        tk.Label(ap_body, textvariable=status, anchor="w", bg=C_BG,
                 fg=C_PRIMARY, font=F).pack(fill="x", pady=(4, 0))
        buttons = tk.Frame(ap_body, bg=C_BG)
        buttons.pack(fill="x", pady=(10, 0))
        lbtn = dict(font=F, bg=C_CARD, fg=C_INK, bd=1, relief="solid",
                    highlightthickness=0, padx=10, pady=5,
                    activebackground="#f9fafb")
        tk.Button(buttons, text="选择新版连屏图（可多选）",
                  command=pick_images, **lbtn).pack(side="left")
        tk.Button(buttons, text="清空", command=lambda: (files.clear(), refresh()),
                  **lbtn).pack(side="left", padx=7)

        start_button = tk.Button(buttons, text="开始转换", bg=C_PURPLE, fg="white",
                                 font=("Microsoft YaHei UI", 11, "bold"),
                                 bd=0, padx=24, pady=9,
                                 activebackground=C_PURPLE_D, activeforeground="white")
        start_button.pack(side="right")

        def run_worker(selected, selected_mode, folder, jpg_quality):
            succeeded, failed = 0, []
            # PSD / PSB 不能被 PIL 读取，先用 Photoshop 渲染成临时 JPG 再切图。
            tmp_dir = Path(tempfile.gettempdir()) / f"psdtool_aplus_{os.getpid()}"
            for src in selected:
                temp_jpg = None
                source_psd = None
                try:
                    work = Path(src)
                    if work.suffix.lower() in SUPPORTED_EXTS:
                        source_psd = work
                        self.ui(lambda n=work.name: status.set(f"正在渲染 {n}（Photoshop 大文件请稍候）…"))
                        tmp_dir.mkdir(parents=True, exist_ok=True)
                        temp_jpg = tmp_dir / f"aplus_{abs(hash(str(work)))}{work.stem}.jpg"
                        render_psd_to_jpg(work, temp_jpg)
                        work = temp_jpg

                    if is_a_plus_contact_sheet(work):
                        # 3648px 等宽的电脑端+手机端组合总图：自动识别两列和白色间隔。
                        ref = Path(folder) if (selected_mode == "reference" and folder) else None
                        outputs = []
                        # 优先读取 PSD / PSB 的画板名字来命名（画板名即业务名）。
                        if source_psd is not None:
                            try:
                                self.ui(lambda: status.set("正在读取画板名字…"))
                                mf = tmp_dir / f"artboards_{abs(hash(str(source_psd)))}.txt"
                                boards = read_artboards(source_psd, mf)
                                if boards:
                                    outputs = split_a_plus_by_artboards(
                                        work, boards, ref, jpg_quality)
                                    if outputs:
                                        self.log(f"√ 高级A+：{src.name} 使用画板名字命名（{len(boards)} 个画板）")
                            except Exception as be:
                                self.log(f"! 画板名字读取失败，回退参考命名：{be}")
                                outputs = []
                        if not outputs:
                            # 读不到画板则沿用参考文件夹的中文业务名。
                            outputs = split_a_plus_contact_sheet(
                                work, ref, jpg_quality, reference_root=ref)
                    elif selected_mode == "reference":
                        outputs = split_a_plus_reference(work, Path(folder), jpg_quality)
                    else:
                        outputs = split_a_plus_fixed(work, None, jpg_quality)
                    succeeded += 1
                    self.log(f"√ 高级A+：{src.name} → {len(outputs)} 张")
                except Exception as exc:
                    failed.append(f"{src.name}：{exc}")
                finally:
                    if temp_jpg is not None:
                        try:
                            temp_jpg.unlink(missing_ok=True)
                        except OSError:
                            pass

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
            # 记住质量与最近参考目录（参考模式即用户明确同意覆盖旧版分图）
            self.aplus_quality = quality.get()
            self.aplus_reference = folder
            self.persist_settings()
            start_button.configure(state="disabled")
            status.set(f"正在处理 {len(files)} 张连屏图……")
            threading.Thread(
                target=run_worker,
                args=(list(files), selected_mode, folder, quality.get()),
                daemon=True).start()

        start_button.configure(command=start_a_plus)

        # 显示从主界面带入的 PSD / PSB；若有文件则提示用户直接点开始。
        refresh()
        if files:
            status.set(f"已带入主界面的 {len(files)} 个连屏源文件，可直接点「开始转换」")
            self.log(f"高级 A+：已从主界面带入 {len(files)} 个 PSD / PSB 源文件")

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
            # Tk 的 splitlist 能解析带空格/中文的路径；再经 clean_drop_list
            # 兜底修复“无空格反斜杠路径被 Tcl 转义”的情况。
            paths = clean_drop_list(self.root.tk.splitlist(event.data))
            aplus_mode = self.output_mode.get().startswith("高级 A+")
            if aplus_mode:
                # 高级 A+：图片与 PSD / PSB 都收；拖入后立即进入专用工作区，
                # 让用户在列表里看到文件，而不是“拖进去没反应”。
                existing = {os.path.normcase(os.path.abspath(str(p)))
                            for p in self.files}
                added = [p for p in collect_aplus(paths)
                         if os.path.normcase(os.path.abspath(str(p)))
                         not in existing]
                self.files.extend(added)
                if added:
                    self.status.set(
                        f"已添加 {len(added)} 个 A+ 连屏图，正在打开高级 A+ 工作区…")
                    self.log(f"高级 A+：拖入 {len(added)} 个文件，打开专用工作区")
                    self.root.after(120, lambda: self._open_aplus_with(added))
                else:
                    self.status.set("未识别到可用文件（支持 JPG / PNG / PSD / PSB）")
                    self.log("! 高级 A+ 拖入没有可用文件；支持 JPG / PNG / PSD / PSB")
                return getattr(event, "action", "copy") or "copy"

            # 每次新的拖放代表一批新任务。清掉上次完成后仍保留的队列，
            # 避免先拖橙色、再拖蓝色时把旧橙色再次一起转换。
            if not self.busy and self.files:
                previous = len(self.files)
                self.files = []
                self.log(f"— 已替换上次队列（{previous} 个文件），仅处理本次拖入内容")
            self.add_paths(paths)
            # 普通导图模式默认“拖入即转换”。
            if self.files and not self.busy:
                self.root.after(120, self.start)
        except Exception:
            self.log("! 拖放数据解析失败，请改用按钮选择")
        # TkDND 回调必须返回一个动作；返回 None 可能被 Tcl 当作无效动作。
        return getattr(event, "action", "copy") or "copy"

    def _open_aplus_with(self, paths=None):
        """打开（或激活）高级 A+ 工作区，并补入本次拖入的文件。"""
        win = getattr(self, "_aplus_win", None)
        if win is not None:
            try:
                if win.winfo_exists():
                    if paths and getattr(self, "_aplus_add_images", None):
                        self._aplus_add_images(paths)
                    win.deiconify()
                    win.lift()
                    return
            except Exception:
                pass
        self.open_a_plus()
        if paths and getattr(self, "_aplus_add_images", None):
            self._aplus_add_images(paths)

    def _on_aplus_close(self, win):
        """关闭高级 A+ 工作区时清理引用，下次可重新打开。"""
        try:
            win.destroy()
        finally:
            if getattr(self, "_aplus_win", None) is win:
                self._aplus_win = None
                self._aplus_add_images = None

    def pick_files(self):
        if self.output_mode.get().startswith("高级 A+"):
            paths = self.filedialog.askopenfilenames(
                title="选择高级 A+ 连屏图（图片或 PSD / PSB）",
                filetypes=[("图片 / PSD / PSB",
                            "*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.psd;*.psb")])
            if paths:
                existing = {os.path.normcase(os.path.abspath(str(p)))
                            for p in self.files}
                added = [p for p in collect_aplus(paths)
                         if os.path.normcase(os.path.abspath(str(p)))
                         not in existing]
                self.files.extend(added)
                if added:
                    self.status.set(
                        f"已添加 {len(added)} 个 A+ 连屏图，正在打开高级 A+ 工作区…")
                    self.root.after(120, lambda: self._open_aplus_with(added))
            return
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
        if stage.startswith("preflight_done:"):
            total = stage.rsplit(":", 1)[-1]
            self.current_progress.configure(mode="determinate", value=22)
            self.current_progress_label.set(f"预检完成：检测到 {total} 个画板")
            self.log(f"✓ {name}：检测到 {total} 个可导出画板/顶层分组")
            return
        if stage.startswith("board:"):
            parts = stage.split(":", 3)
            if len(parts) == 4:
                try:
                    current, total = int(parts[1]), max(1, int(parts[2]))
                except ValueError:
                    current, total = 1, 1
                percent = 22 + round(68 * current / total)
                board_name = parts[3]
                self.current_progress.configure(mode="determinate", value=percent)
                self.current_progress_label.set(
                    f"导出画板 {current}/{total}：{board_name[:22]}")
                return
        percent, label = STAGE_INFO.get(stage, (0, stage))
        short_name = name if len(name) <= 14 else name[:12] + "…"
        self.current_progress.configure(mode="determinate", value=percent)
        self.current_progress_label.set(f"{label}：{short_name}")

    # -- conversion ---------------------------------------------------------
    def start(self):
        if self.busy:
            return
        if self.output_mode.get().startswith("高级 A+"):
            self.open_a_plus()
            return
        if not self.files:
            self.messagebox.showinfo(APP_TITLE, "请先拖入或选择 PSD / PSB 文件")
            return
        formats = self.formats()
        use_folder = self.output_mode.get().startswith("主副图画板")
        if use_folder:
            # 画板模式的业务输出固定为：每个画板一张同名最高质量 JPG。
            formats = ("jpg",)
        if not formats:
            self.messagebox.showwarning(APP_TITLE, "至少勾选一种导出格式（PNG 或 JPG）")
            return

        self.busy = True
        self.stop_requested.clear()
        self.output_dirs = []
        self.btn_open_dir.configure(state="disabled")
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
        conflict = self.conflict.get()
        if conflict not in CONFLICT_CHOICES:
            conflict = "number"
        mode_label = "主副图画板" if use_folder else "渲染图片"
        conflict_label = {"number": "自动编号", "overwrite": "覆盖",
                          "skip": "跳过"}[conflict]
        self.log(f"— 开始：模式={mode_label}，{len(files)} 个文件，"
                 f"裁切={trim_mode}，格式={'+'.join(formats)}，同名={conflict_label}")
        self.persist_settings()

        threading.Thread(target=self.worker,
                         args=(files, trim_mode, formats, use_folder, conflict),
                         daemon=True).start()

    def worker(self, files, trim_mode, formats, use_folder=False,
               conflict="number"):
        ok = 0
        failed = 0
        skipped = 0
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
                if use_folder:
                    directory = src.parent / src.stem
                    produced_paths = convert_artboards(
                        src, directory,
                        progress_cb=lambda stage, name=name: self.ui(
                            lambda stage=stage, name=name: self.update_stage(name, stage)),
                        cancel_event=self.stop_requested,
                        info_cb=lambda notice, name=name: self.ui(
                            lambda notice=notice, name=name:
                            self.log_artboard_notice(name, notice)))
                    outputs = {f"artboard_{n}": p
                               for n, p in enumerate(produced_paths, 1)}
                else:
                    outputs = unique_outputs(src, formats, False, conflict)
                if outputs is None:
                    # skip 策略：目标图片已存在，明确记为跳过，不计入失败。
                    skipped += 1
                    self.ui(lambda name=name: self.log(
                        f"— 已跳过：{name}：目标文件已存在"))
                else:
                    if not use_folder:
                        convert_one(
                            src, outputs, trim_mode, formats,
                            progress_cb=lambda stage, name=name: self.ui(
                                lambda stage=stage, name=name: self.update_stage(name, stage)))
                    inherit_source_timestamp(src, outputs.values())
                    ok += 1
                    # 记录实际输出目录，供“打开输出目录”使用。
                    dirs = {p.parent for p in outputs.values()}
                    for directory in dirs:
                        if directory not in self.output_dirs:
                            self.output_dirs.append(directory)
                    produced = " + ".join(p.name for p in outputs.values())
                    output_location = str(next(iter(outputs.values())).parent)
                    self.ui(lambda name=name, produced=produced,
                            output_location=output_location: self.log(
                        f"√ {name} → {produced}\n  保存位置：{output_location}"))
            except ConversionCancelled as exc:
                self.ui(lambda msg=str(exc): self.log("! " + msg))
                processed = i
                break
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

        done_ok, done_fail, done_skip = ok, failed, skipped
        stopped = self.stop_requested.is_set()
        remaining = len(files) - processed
        self.ui(lambda: self.finish(done_ok, done_fail, stopped, remaining,
                                    done_skip))

    def log_artboard_notice(self, filename, notice):
        if notice.startswith("META:"):
            name, _, size = notice[5:].partition("|")
            self.log(f"  · 预检：{name}（{size or '尺寸未知'}）")
        elif notice.startswith("WARN:"):
            self.log(f"! {filename}：{notice[5:]}")
        elif notice.startswith("BOARD_ERROR:"):
            self.log(f"× {filename}：{notice[12:]}")

    def finish(self, ok, failed, stopped=False, remaining=0, skipped=0):
        self.busy = False
        self.current_progress.stop()
        self.current_progress.configure(mode="determinate", value=0)
        self.current_progress_label.set("当前文件：空闲")
        self.btn_stop.configure(state="disabled")
        for b in (self.btn_files, self.btn_folder, self.btn_go, self.btn_clear):
            b.configure(state="normal")
        # 有成功输出才启用“打开输出目录”，且只打开一个资源管理器窗口。
        if self.output_dirs and ok > 0:
            self.btn_open_dir.configure(state="normal")
        skip_txt = f"，跳过 {skipped}" if skipped else ""
        if stopped:
            self.status.set(f"已停止：成功 {ok}，失败 {failed}，未处理 {remaining}{skip_txt}")
            self.log(f"— 队列已停止：成功 {ok}，失败 {failed}，未处理 {remaining}{skip_txt}")
        else:
            self.status.set(f"完成：成功 {ok}，失败 {failed}{skip_txt}")
            self.log(f"— 全部结束：成功 {ok}，失败 {failed}{skip_txt}")
        if failed:
            self.messagebox.showwarning(APP_TITLE,
                                        f"完成，但有 {failed} 个文件失败。\n详情见下方日志。")

    def open_output_dir(self):
        """只打开一个资源管理器：多文件同目录开共同目录，不同目录开第一个。"""
        if not self.output_dirs:
            return
        target = Path(self.output_dirs[0])
        unique = {os.path.normcase(os.path.abspath(str(d))) for d in self.output_dirs}
        if len(unique) == 1:
            target = Path(self.output_dirs[0])
        else:
            target = Path(self.output_dirs[0])  # 不同目录时只打开第一个
        try:
            os.startfile(str(target))
        except OSError as exc:
            self.log(f"! 无法打开输出目录：{exc}")

    def run(self):
        self.root.mainloop()


LegacyApp = App


class ModernApp(LegacyApp):
    """CustomTkinter 现代界面；转换、拖放和任务逻辑沿用稳定后端。"""

    def _build_ui(self):
        import customtkinter as ctk
        from tkinter import ttk

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.root.geometry("1040x720")
        self.root.minsize(920, 640)
        self.root.configure(bg="#eef2f7")
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        font = ("Microsoft YaHei UI", 13)
        small = ("Microsoft YaHei UI", 11)
        title_font = ("Microsoft YaHei UI", 25, "bold")

        # 左侧导航栏
        sidebar = ctk.CTkFrame(self.root, width=218, corner_radius=0,
                               fg_color="#111827")
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        try:
            from PIL import Image
            self.brand_icon = ctk.CTkImage(
                light_image=Image.open(resource_path("assets/app_icon.png")),
                dark_image=Image.open(resource_path("assets/app_icon.png")),
                size=(58, 58))
            ctk.CTkLabel(sidebar, text="", image=self.brand_icon,
                         width=58, height=58).pack(anchor="w", padx=20,
                                                   pady=(22, 8))
        except Exception:
            ctk.CTkLabel(sidebar, text="PS", width=52, height=52,
                         corner_radius=14, fg_color="#2563eb", text_color="white",
                         font=("Segoe UI", 22, "bold")).pack(
                             anchor="w", padx=22, pady=(25, 10))
        ctk.CTkLabel(sidebar, text="PSD Studio", text_color="#f9fafb",
                     font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=22)
        ctk.CTkLabel(sidebar, text="电商设计批量导图工具", text_color="#9ca3af",
                     font=small).pack(anchor="w", padx=22, pady=(2, 28))

        self.output_mode = self.tk.StringVar(
            value=MODE_VALUES.get(self.settings["mode"], MODE_VALUES["render"]))
        self.mode_buttons = []
        mode_items = (
            ("渲染图片", "JPG 放在源文件旁", MODE_VALUES["render"], "▣"),
            ("主副图画板", "自动建立同名文件夹", MODE_VALUES["artboard"], "▦"),
            ("高级 A+", "电脑端 / 手机端", MODE_VALUES["aplus"], "✦"),
        )

        def choose_mode(value):
            self.output_mode.set(value)
            self.sync_mode_ui()

        for heading, subtitle, value, icon in mode_items:
            button = ctk.CTkButton(
                sidebar, text=f"{icon}   {heading}\n      {subtitle}", anchor="w",
                height=62, corner_radius=12, font=small,
                fg_color="transparent", hover_color="#1f2937",
                text_color="#d1d5db", command=lambda v=value: choose_mode(v))
            button.pack(fill="x", padx=14, pady=4)
            self.mode_buttons.append((button, value))

        ctk.CTkLabel(sidebar, text=f"版本  {VERSION}", text_color="#6b7280",
                     font=("Segoe UI", 10)).pack(side="bottom", anchor="w",
                                                   padx=22, pady=20)

        # 主工作区
        main = ctk.CTkFrame(self.root, corner_radius=0, fg_color="#eef2f7")
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        top = ctk.CTkFrame(main, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=28, pady=(20, 10))
        self.page_title = ctk.CTkLabel(top, text="渲染图片",
                                      text_color="#111827", font=title_font)
        self.page_title.pack(anchor="w")
        self.page_hint = ctk.CTkLabel(
            top, text="拖入即转换 · 自动补写 PSD 缩略图",
            text_color="#64748b", font=small)
        self.page_hint.pack(anchor="w", pady=(2, 0))

        self.drop = self.tk.Label(
            main, text="⇧\n将 PSD / PSB 文件或文件夹拖到这里\n也可以点击下方按钮选择",
            font=("Microsoft YaHei UI", 16, "bold"), bg="#eff6ff", fg="#1d4ed8",
            justify="center", relief="flat", bd=0, cursor="hand2")
        self.drop.grid(row=1, column=0, sticky="ew", padx=28, pady=(0, 14),
                       ipady=36)
        self.drop.configure(highlightthickness=2, highlightbackground="#60a5fa",
                            highlightcolor="#2563eb")
        self.drop.bind("<Button-1>", lambda _event: self.pick_files())

        content = ctk.CTkFrame(main, fg_color="transparent")
        content.grid(row=2, column=0, sticky="nsew", padx=28)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(2, weight=1)

        settings = ctk.CTkFrame(content, corner_radius=16, fg_color="#ffffff",
                                border_width=1, border_color="#e2e8f0")
        settings.grid(row=0, column=0, sticky="ew")
        settings.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(settings, text="导出设置", font=("Microsoft YaHei UI", 14, "bold"),
                     text_color="#1e293b").grid(row=0, column=0, columnspan=4,
                                                 sticky="w", padx=18, pady=(14, 8))
        ctk.CTkLabel(settings, text="裁切方式", font=small,
                     text_color="#64748b").grid(row=1, column=0, sticky="w",
                                                 padx=(18, 8), pady=8)
        self.trim = self.tk.StringVar(
            value=TRIM_MODES.get(self.settings["trim"], TRIM_MODES["transparent"]))
        ctk.CTkOptionMenu(settings, variable=self.trim,
                          values=list(TRIM_MODES.values()), height=34,
                          corner_radius=9, fg_color="#f1f5f9",
                          button_color="#e2e8f0", button_hover_color="#cbd5e1",
                          text_color="#334155", dropdown_fg_color="white",
                          dropdown_text_color="#334155", font=small).grid(
                              row=1, column=1, sticky="ew", padx=(0, 16), pady=8)

        self.var_png = self.tk.BooleanVar(value=self.settings["png"])
        self.var_jpg = self.tk.BooleanVar(value=self.settings["jpg"])
        self.cb_png = ctk.CTkCheckBox(settings, text="PNG 透明", variable=self.var_png,
                                      font=small, corner_radius=5)
        self.cb_png.grid(row=1, column=2, padx=8, pady=8)
        self.cb_jpg = ctk.CTkCheckBox(settings, text="JPG 最高质量",
                                      variable=self.var_jpg,
                                      font=small, corner_radius=5)
        self.cb_jpg.grid(row=1, column=3, padx=(8, 18), pady=8)

        ctk.CTkLabel(settings, text="同名文件", font=small,
                     text_color="#64748b").grid(row=2, column=0, sticky="w",
                                                 padx=(18, 8), pady=(6, 16))
        self.conflict = self.tk.StringVar(value=self.settings["conflict"])
        self.conflict_controls = []
        conflict_frame = ctk.CTkFrame(settings, fg_color="transparent")
        conflict_frame.grid(row=2, column=1, columnspan=3, sticky="w",
                            pady=(6, 16))
        for label, value in (("自动编号", "number"), ("覆盖", "overwrite"),
                             ("跳过", "skip")):
            item = ctk.CTkRadioButton(conflict_frame, text=label,
                                      variable=self.conflict, value=value,
                                      font=small, radiobutton_width=18,
                                      radiobutton_height=18)
            item.pack(side="left", padx=(0, 20))
            self.conflict_controls.append(item)

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", pady=12)
        self.btn_files = ctk.CTkButton(actions, text="选择文件", width=105,
                                       height=38, corner_radius=10,
                                       fg_color="#ffffff", hover_color="#f1f5f9",
                                       border_width=1, border_color="#cbd5e1",
                                       text_color="#334155", command=self.pick_files)
        self.btn_folder = ctk.CTkButton(actions, text="选择文件夹", width=115,
                                        height=38, corner_radius=10,
                                        fg_color="#ffffff", hover_color="#f1f5f9",
                                        border_width=1, border_color="#cbd5e1",
                                        text_color="#334155", command=self.pick_folder)
        self.btn_clear = ctk.CTkButton(actions, text="清空", width=72, height=38,
                                       corner_radius=10, fg_color="transparent",
                                       hover_color="#e2e8f0", text_color="#64748b",
                                       command=self.clear)
        self.btn_stop = ctk.CTkButton(actions, text="停止", width=72, height=38,
                                      corner_radius=10, fg_color="#fee2e2",
                                      hover_color="#fecaca", text_color="#b91c1c",
                                      command=self.request_stop, state="disabled")
        self.btn_go = ctk.CTkButton(actions, text="开始转换", width=126, height=40,
                                    corner_radius=11, font=("Microsoft YaHei UI", 12, "bold"),
                                    command=self.start)
        self.btn_open_dir = ctk.CTkButton(
            actions, text="打开输出目录", width=118, height=38, corner_radius=10,
            fg_color="#ffffff", hover_color="#f1f5f9", border_width=1,
            border_color="#cbd5e1", text_color="#334155",
            command=self.open_output_dir, state="disabled")
        for widget in (self.btn_files, self.btn_folder, self.btn_clear, self.btn_stop):
            widget.pack(side="left", padx=(0, 8))
        self.btn_go.pack(side="right")
        self.btn_open_dir.pack(side="right", padx=(0, 8))

        status_card = ctk.CTkFrame(content, corner_radius=16, fg_color="#ffffff",
                                   border_width=1, border_color="#e2e8f0")
        status_card.grid(row=2, column=0, sticky="nsew", pady=(0, 18))
        status_card.grid_columnconfigure(0, weight=1)
        status_card.grid_rowconfigure(4, weight=1)
        self.status = self.tk.StringVar(value="等待文件")
        ctk.CTkLabel(status_card, textvariable=self.status, anchor="w",
                     font=small, text_color="#334155").grid(
                         row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
        style = ttk.Style()
        style.configure("Modern.Horizontal.TProgressbar", troughcolor="#e2e8f0",
                        background="#2563eb", borderwidth=0)
        self.progress = ttk.Progressbar(status_card, mode="determinate",
                                        style="Modern.Horizontal.TProgressbar")
        self.progress.grid(row=1, column=0, sticky="ew", padx=16)
        self.current_progress_label = self.tk.StringVar(value="当前文件：空闲")
        ctk.CTkLabel(status_card, textvariable=self.current_progress_label,
                     anchor="w", font=("Microsoft YaHei UI", 10),
                     text_color="#64748b").grid(row=2, column=0, sticky="ew",
                                                 padx=16, pady=(8, 4))
        self.current_progress = ttk.Progressbar(
            status_card, mode="determinate", maximum=100, value=0,
            style="Modern.Horizontal.TProgressbar")
        self.current_progress.grid(row=3, column=0, sticky="ew", padx=16)
        self.logbox = ctk.CTkTextbox(status_card, corner_radius=10,
                                     fg_color="#f8fafc", text_color="#334155",
                                     font=("Consolas", 10), border_width=0)
        self.logbox.grid(row=4, column=0, sticky="nsew", padx=16, pady=12)
        self.logbox.configure(state="disabled")
        self.sync_mode_ui()

    def sync_mode_ui(self):
        mode = self.output_mode.get()
        if not hasattr(self, "btn_go"):
            return
        for button, value in self.mode_buttons:
            selected = value == mode
            is_aplus = value.startswith("高级 A+")
            active = "#7c3aed" if is_aplus else "#2563eb"
            button.configure(fg_color=active if selected else "transparent",
                             hover_color="#6d28d9" if is_aplus else "#1f2937",
                             text_color="white" if selected else "#d1d5db")
        aplus = mode.startswith("高级 A+")
        artboard = mode.startswith("主副图画板")
        for item in self.conflict_controls:
            item.configure(state="disabled" if (aplus or artboard) else "normal")
        if hasattr(self, "cb_png"):
            self.cb_png.configure(state="disabled" if (aplus or artboard) else "normal")
            self.cb_jpg.configure(state="disabled" if (aplus or artboard) else "normal")
        if artboard:
            self.var_png.set(False)
            self.var_jpg.set(True)
        if aplus:
            self.page_title.configure(text="高级 A+ 连屏")
            self.page_hint.configure(text="自动生成电脑端与手机端分图")
            self.drop.configure(text="✦\n高级 A+ 使用独立工作区\n点击“打开高级 A+”开始",
                                bg="#f5f3ff", fg="#6d28d9",
                                highlightbackground="#a78bfa")
            self.btn_go.configure(text="高级 A+ 开始转换", fg_color="#7c3aed",
                                  hover_color="#6d28d9")
        elif mode.startswith("主副图画板"):
            self.page_title.configure(text="主副图画板")
            self.page_hint.configure(text="每个画板单独导出 JPG，并使用画板名称命名")
            self.drop.configure(text="⇧\n将主副图 PSD / PSB 拖到这里\n自动建立同名输出文件夹",
                                bg="#eff6ff", fg="#1d4ed8",
                                highlightbackground="#60a5fa")
            self.btn_go.configure(text="开始转换", fg_color="#2563eb",
                                  hover_color="#1d4ed8")
        else:
            self.page_title.configure(text="渲染图片")
            self.page_hint.configure(text="拖入即转换 · JPG 存在源文件旁 · 自动补写 PSD 缩略图")
            self.drop.configure(text="⇧\n将渲染 PSD / PSB 文件或文件夹拖到这里\n也可以点击下方按钮选择",
                                bg="#eff6ff", fg="#1d4ed8",
                                highlightbackground="#60a5fa")
            self.btn_go.configure(text="开始转换", fg_color="#2563eb",
                                  hover_color="#1d4ed8")


# 默认启用现代界面；LegacyApp 保留作故障回退和代码对照。
App = ModernApp


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
            inherit_source_timestamp(src, outputs.values())
            emit("    -> " + " + ".join(p.name for p in outputs.values()))
            emit("    保存位置：" + str(next(iter(outputs.values())).parent))
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
