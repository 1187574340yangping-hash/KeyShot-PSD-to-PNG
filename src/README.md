# src/ · 源码说明

本目录有**两套源码**，对应两个版本，请勿混用：

| 文件 | 对应版本 | 性质 |
| --- | --- | --- |
| `psd_to_png_gui.py` | **v2（当前维护，推荐）** | 完整源码，Python + tkinter |
| `PsdToPng.cs` | v1 | 依据 `dist/PSD一键转PNG.exe` **重建的行为等价参考实现**，不是随原工具分发的原始源码 |
| `build.bat` | v1 | C# 编译脚本（Windows 自带 `csc.exe`） |

两个版本共用同一条 Photoshop COM 链路，但实现语言、GUI 框架、功能范围都不同。逐条设计理由见 [`docs/DESIGN.md`](../docs/DESIGN.md)。

---

## v2 · `psd_to_png_gui.py`

### 运行环境

| 项 | 要求 |
| --- | --- |
| 操作系统 | Windows（COM + 桌面 GUI，不能跑 macOS / Linux） |
| Python | 3.13（开发态；分发版已由 PyInstaller 冻结，用户无需安装） |
| 第三方包 | `pywin32`（`win32com.client` / `pythoncom`）、`tkinterdnd2`（稳定窗口拖放）、`PyInstaller`（仅打包时） |
| 标准库 | `tkinter`、`argparse`、`pathlib`、`threading`、`json` |
| 外部依赖 | **Adobe Photoshop 桌面完整版**，且 `Photoshop.Application` COM 已注册 |

`tkinterdnd2` 不可用时会自动降级为「选择文件 / 选择文件夹」按钮，程序仍可正常使用。旧版使用的 `windnd` 在部分 64 位 Python/Tk 环境会触发 `ucrtbase.dll / 0xc0000409`，因此不再使用。

### 三种运行方式

```bat
:: 1) GUI（默认）
python src/psd_to_png_gui.py

:: 2) 开发态自检：不启动 Photoshop，22 项纯逻辑校验
python src/psd_to_png_gui.py --selftest

:: 3) 命令行批量（可挂到脚本里）
python src/psd_to_png_gui.py --cli "F:\杨平-产品图片\...\NTT72" --formats png jpg --trim transparent
```

其它参数：

| 参数 | 说明 |
| --- | --- |
| `--cli PATH [PATH ...]` | 文件或文件夹，文件夹会递归 |
| `--formats png jpg` | 导出格式，可只给一个；两个都不给会报错 |
| `--trim {transparent,topleft,both,none}` | 裁切策略，默认 `transparent` |
| `--probe` | 创建 GUI 后 800ms 自动退出，写 `%TEMP%\psdtool_probe.json`，用于验证冻结包能否正常起窗口 |
| `--selftest` | 纯逻辑自检 |
| `--version` | 输出 `2.0` |

> 把文件 / 文件夹拖到 **EXE 图标**上启动时，路径从 `argv` 传入（代码里 `parse_known_args()` 收集 `unknown`），会自动开始转换。

### 打包

打包脚本在项目根目录 `build_exe.py`（不在 `src/`）：

```bat
cd psd-to-png
python build_exe.py
```

产物：`dist/PSD-PSB一键转PNG-JPG.exe`（约 11 MB，PyInstaller `--onefile --windowed`）。

脚本要点：

- 隐藏导入 `win32com.client` / `pythoncom` / `tkinter.*`，并用 `--collect-all tkinterdnd2` 收集 TkDND 二进制与脚本。
- 排除 `PySide6` / `PyQt5` / `PIL` / `numpy` / `matplotlib` / `scipy` / `pandas` / `pytest`，避免把几百 MB 无关依赖打进去。
- **不删旧产物**：把 `build/work`、`build/out`、`build/spec` 挪到 `%TEMP%\old_build\<时间戳>`，再重建。受限环境下批量删除会触发保护；同时固定 `--workpath`、不加 `--clean`。
- 中间产物 `build/`（约 27 MB）已被 `.gitignore` 排除，只有最终 EXE 入库。

### 源码结构

| 函数 / 类 | 职责 |
| --- | --- |
| `js_escape(s)` | Windows 路径 → JS 字符串字面量（`\` → `\\`，`'` → `\'`） |
| `build_script(psd, png, jpg, trim_mode, formats)` | 拼出完整 ExtendScript，只给勾中的格式生成保存代码 |
| `collect(paths)` | 文件 / 文件夹混合收集，递归、过滤 `.psd` / `.psb`、大小写不敏感去重 |
| `unique_outputs(src, formats)` | 成对分配 PNG / JPG 输出路径，整组避开同名，冲突加 `_PNG` / `_PNG_2` |
| `human_size(n)` | 大文件友好体积显示（1~2 GB 时给进度提示） |
| `photoshop()` | `win32com.client.Dispatch("Photoshop.Application")` |
| `convert_one(...)` | 单张转换：每线程 `CoInitialize()` → `DoJavaScript`；Photoshop 忙时自动重试 → `CoUninitialize()` |
| `selftest()` | 22 项纯逻辑校验，输出 JSON |
| `App` | tkinter GUI：拖放、勾选、进度条、日志、后台 worker、`after(0, fn)` 回主线程 |
| `run_cli(...)` | 命令行批量，复用 `convert_one()`；窗口版无 stdout 时写 `%TEMP%\psdtool_cli.log` |
| `main()` | 参数解析与四种入口分派 |

`--selftest` 的 22 项覆盖：扩展名收集（含大写）、排除非 PS 文件、路径转义、`build_script` 的 PNG/JPG/trim 各分支、未选格式不写入脚本、输出成对且不覆盖、只选单格式时只占单格式名额等。全部通过会输出 `"selftest": "PASS"`。

---

## v1 · `PsdToPng.cs`（历史参考）

`PsdToPng.cs` 是**依据 `dist/PSD一键转PNG.exe` 重建的行为等价参考实现**，不是随原工具分发的原始源码。

原程序只发布了一个 13,312 字节的 .NET Framework 4.0 WinForms EXE。本文件根据以下证据重建：

| 证据 | 来源 |
| --- | --- |
| 目标框架 `v4.0.30319`、程序集引用 | CLI 元数据头与 `#Strings` 堆 |
| `MainForm` / `AcceptPaths` / `StartConversion` / `ConvertFiles` / `UniquePng` / `JsEscape` / `DoJavaScript` / `FinalReleaseComObject` | 元数据方法名 |
| 全部界面文案、文件过滤器、`Photoshop.Application`、ExtendScript 模板 | `#US` 用户字符串堆 |
| `STAThread`、`Thread.SetApartmentState`、`BeginInvoke`、`COMException`、`Activator.CreateInstance`、`Directory.GetFiles(..., AllDirectories)`、`Distinct(OrdinalIgnoreCase)` | IL 引用的 API |

原始字节证据见 [`docs/REVERSE_ENGINEERING.md`](../docs/REVERSE_ENGINEERING.md)。

### 构建

```bat
cd src
build.bat
```

`build.bat` 调用 Windows 自带的 .NET Framework 4.x `csc.exe`，不需要 Visual Studio、SDK 或 NuGet。产物写到 `dist/PsdToPng.rebuilt.exe`。

要点：

- `/target:winexe` — WinForms 窗口程序，不弹控制台
- `/r:Microsoft.CSharp.dll` — `dynamic ps = app;` 调用 `DoJavaScript` 需要运行时绑定器（原程序同样引用了 `Microsoft.CSharp.RuntimeBinder`）
- `/codepage:65001` — 源文件是 UTF-8，里面有中文界面文案，不加会用系统 ANSI 读取导致乱码
- 脚本内路径全为 ASCII，避免 `.bat` 的 ANSI 编码把中文路径写坏

也可以用 Visual Studio 新建「Windows 窗体应用程序 (.NET Framework)」，把 `PsdToPng.cs` 加进去，添加 `Microsoft.CSharp` 引用后直接生成。

### 与原 EXE 的关系

- **行为对齐**：三种输入入口、递归收集、大小写不敏感去重、输出命名、Photoshop JS 脚本、不保存原 PSD、COM 释放、错误文案。
- **实现差异**：UI 控件坐标与配色是近似还原；异常分支改为「单张失败继续下一张」，并在 `finally` 里统一恢复按钮状态；`AcceptPaths` 增加了无权限路径的跳过保护。
- **原版优先**：日常使用请直接运行分发版 EXE；v1 重建源码主要用于理解原理、修改规则和二次开发。

> 重建源码未在本次环境中执行编译验证（环境策略不允许调用任意 C# 编译器）。如编译报错，请以 `docs/REVERSE_ENGINEERING.md` 中的原始 JS 模板与文案为准修正。
