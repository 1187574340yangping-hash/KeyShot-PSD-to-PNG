# PSD / PSB 一键转 PNG · JPG

> 大画板、四周全是留白？拖进去，自动裁到内容边界，顺手出 PNG 和 JPG。原稿一个字节都不动。

`Windows` · `Photoshop COM / ExtendScript` · `Python + tkinter（v2）` · `MIT`

一个绿色小工具：把 **PSD 和 PSB（Photoshop 大型文档）批量导出成 PNG / JPG**，并且在导出前**自动裁掉画布四周的大片空白**。它不自己解析文件，而是让本机已安装的 Adobe Photoshop 打开、裁切、另存——图层、混合模式、智能对象、蒙版、色彩配置的视觉效果与你在 Photoshop 里手工操作完全一致。

## 版本

| | 程序 | 能力 |
| --- | --- | --- |
| **v2（推荐）** | [`dist/PSD-PSB一键转PNG-JPG.exe`](dist/PSD-PSB一键转PNG-JPG.exe) | **支持 .psd + .psb**、自动裁切到内容边界、PNG / JPG 可勾选、大文件进度提示 |
| **v2.1 目录修复版（本机推荐）** | [`dist/PSDTool-onedir-fixed.zip`](dist/PSDTool-onedir-fixed.zip) | 加固窗口拖入、增加当前文件动态进度与总进度、修复无控制台 CLI 报错；解压后运行 `PSD一键转PNG工具/PSDTool.exe` |
| v1（原版，保留） | [`dist/PSD一键转PNG.exe`](dist/PSD一键转PNG.exe)（13 KB .NET） | 只认 .psd、不裁切、只出 PNG |

> v1 的原始二进制与逆向证据完整保留在仓库里，见 [`docs/REVERSE_ENGINEERING.md`](docs/REVERSE_ENGINEERING.md)。

---

## 你遇到的那个问题

`F:\...\NTT72\NTT72-D.psb`（1.65 GB）拖不进 v1、也导不出 PNG —— **不是文件坏了，是 v1 只认 `.psd`**：文件对话框过滤器、拖入时的扩展名判断、文件夹递归三处都写死 `*.psd`。而 PSB 是 Photoshop 在文档超过 2 GB 或 30,000 像素时自动使用的「大型文档格式」，渲染输出、大画板拼版几乎都是它。

同一个目录里就有 5 个 1~2 GB 的 `.psb`。v2 把扩展名放宽到 `.psd` + `.psb`，并按你的实际场景补了两件事：

1. **自动裁切到内容边界** —— 大画板里内容只占中间一小块、四周大量透明/白色留白时，导出的图不再带着一整圈空白。
2. **导出格式可勾选** —— 只要 PNG（保留透明）、只要 JPG（自动填白）、或两者都要；两种格式共用同一个文件名，成对出现。

---

## 30 秒上手

1. 下载 [`dist/PSD-PSB一键转PNG-JPG.exe`](dist/PSD-PSB一键转PNG-JPG.exe)，双击打开（无需安装、无需联网、无配置文件）。
2. 把 **PSD / PSB 文件或整个文件夹**拖进窗口；也可以点「选择文件」多选或「选择文件夹」递归。
3. 选**裁切方式**（默认「裁掉透明边」）和**导出格式**（默认 PNG + JPG 都勾），点「开始转换」。

> 更快：把文件 / 文件夹直接拖到 EXE **图标**上启动，会自动开始转换。

**前置条件**：Windows + 已安装 Adobe Photoshop 桌面完整版。详见[运行环境](#运行环境与前置条件)。

---

## 适用什么客户

| 谁 | 具体痛点 | 本工具给什么 |
| --- | --- | --- |
| **产品渲染 / 大画板拼版**（本次场景） | 一张 PS 大画板里排了多张详情图，四周留白巨大；文件动辄 1~2 GB 存成 `.psb`，普通转换工具直接不认 | 认 `.psb`，自动把画布裁到实际内容边界，一次出 PNG + JPG |
| **3D 渲染师 / KeyShot 用户** | 渲染输出、分层 PSD 要变成贴图或预览 PNG；手动另存打断渲染流程 | 一次拖放批量出图，尺寸与 alpha 原样保留（不裁切时），路径与源文件同目录 |
| **电商视觉 / 产品图团队** | 大量白底图、透明底 PSD 要批量出 PNG 上架，平台不收 PSD；同时还要一份 JPG 给详情页 | 格式可勾选，一次导出两种；输出自动避开同名，重复跑不覆盖已上架的图 |
| **设计助理 / 美工** | 重复的「打开 → 图像 → 裁切 → 另存为 PNG → 再另存为 JPG」机械劳动 | 裁切 + 双格式一次完成，进度条显示 `3/20` 与当前文件名、文件大小 |
| **摄影后期 / 修图工作室** | 客户要 PNG 预览，但原始分层 PSD 必须保持可编辑 | `DONOTSAVECHANGES` 保证原稿绝不被写回 |
| **独立设计师 / 小团队** | 不想为「批量转格式 + 裁边」买插件、装 Action 包或搭脚本环境 | 单 EXE、免安装、只依赖已有 Photoshop |

> 不适合：脱离 Photoshop 的服务器 / CI、macOS / Linux、AI / EPS / TIFF 等其它格式。见[能力边界](#能力边界与已知限制)。

---

## 工作流图解

### 1. 端到端：从大画板到可用的两张图

```mermaid
flowchart LR
    A["PS 大画板<br/>PSD / PSB 1~2 GB"] --> B["拖入 v2"]
    C["四周大片透明 / 白色留白"] --> B
    B --> D["Photoshop.Application<br/>DoJavaScript"]
    D --> E["trim 裁到内容边界"]
    E --> F1["PNG 保留透明"]
    E --> F2["JPG 自动填白"]
    F1 --> G1["贴图 / 设计交付"]
    F2 --> G2["详情页 / 商品图 / 预览"]
    A -.-> H["原稿不保存更改"]
```

### 2. 四种入口收敛成一条流水线

```mermaid
flowchart TD
    S1["拖入窗口"] --> C
    S2["拖到 EXE 图标启动"] --> C
    S3["选择文件（可多选）"] --> C
    S4["选择文件夹（递归）"] --> C
    C["collect：文件与文件夹混合收集"] --> D["过滤 .psd / .psb<br/>大小写不敏感去重"]
    D --> E{"检测到文件?"}
    E -- 否 --> F["状态栏：没有检测到 PSD / PSB 文件"]
    E -- 是 --> G["读取裁切方式与导出格式"]
    G --> H["后台线程逐张转换<br/>显示文件名 + 体积 + 大画板提示"]
    H --> I["主线程更新状态与进度条"]
    I --> J["恢复按钮，显示成功 / 失败数"]
```

### 3. 单张在 Photoshop 里实际发生什么

```mermaid
flowchart TD
    A["app.open(new File(psd/psb))"] --> B["32 位/通道 → 16 位"]
    B --> C["非 RGB → changeMode(RGB)"]
    C --> D{"裁切方式"}
    D -- 透明边 --> E1["trim(TRANSPARENT)"]
    D -- 纯色边 --> E2["trim(TOPLEFT)"]
    D -- 两者 --> E3["两次 trim"]
    D -- 不裁 --> E4["保留原画布"]
    E1 --> F{"导出格式"}
    E2 --> F
    E3 --> F
    E4 --> F
    F -- 含 PNG --> G1["saveAs PNG（保留 alpha）"]
    F -- 含 JPG --> G2["flatten 填白 → 8 位 → saveAs JPG 质量 12"]
    G1 --> H["close(DONOTSAVECHANGES)"]
    G2 --> H
    G1 -.-> I["原稿始终未被写回"]
```

### 4. 输出命名：成对且永不覆盖

```mermaid
flowchart TD
    A["源文件 D:/NTT72/NTT72-D.psb"] --> B["候选 1：NTT72-D.png + NTT72-D.jpg"]
    B --> C{"任一已存在?"}
    C -- 否 --> Z["写入并结束"]
    C -- 是 --> D["候选 2：NTT72-D_PNG.png + .jpg"]
    D --> E{"任一已存在?"}
    E -- 否 --> Z
    E -- 是 --> F["NTT72-D_PNG_2.png/.jpg → 递增"]
    F --> G{"任一已存在?"}
    G -- 否 --> Z
    G -- 是 --> F
```

---

## 方便什么

- **PSB 直接拖**：不用先把 1~2 GB 的大文档另存成 PSD，也不用改扩展名。
- **不用手动「图像 → 裁切」**：四周空白自动裁掉，四种裁切策略可选；不想要也能关掉。
- **一次出两种格式**：PNG 保留透明给设计/贴图，JPG 自动填白给详情页/上架，勾选即可，不用跑两遍。
- **不会毁稿**：原稿以「不保存更改」关闭；已有输出不会被覆盖，重复跑自动加 `_PNG`。
- **大文件有反馈**：状态栏显示文件名、体积和 `3/20` 进度，日志逐张记录成功/失败；1 GB 以上会明确提示耐心等待。
- **中文 / 全角括号路径可用**：`F:\杨平-产品图片\01-产品图片（产品系列分类）\...` 这类路径已做转义处理。
- **零部署**：单 EXE，不用装 Python、.NET 6/8，不用联网、账号、注册表或配置文件。
- **错误说人话**：Photoshop 被占用提示「Photoshop 忙或无法读取文件」，没装提示「请先确认 Photoshop 已安装」，单张失败继续下一张。

---

## 为什么这样做

| # | 决策 | 原因 |
| --- | --- | --- |
| 1 | **调 Photoshop COM + ExtendScript，不自己解析 PSD/PSB** | 图层、混合模式、智能对象、蒙版、32 位 HDR、1~2 GB 大文档，自己写解析器必然丢效果也吃不下；让 Photoshop 本人操作才是视觉 1:1 |
| 2 | **放宽到 `.psb`** | PSB 是 Photoshop 对超大画布/超大文件的原生格式，渲染输出与拼版 commonplace；扩展名不同但 `app.open` 一视同仁 |
| 3 | `displayDialogs = DialogModes.NO` | 批处理最怕弹窗卡死：缺失字体、颜色配置文件、增效工具提示会让无人值守的转换停住 |
| 4 | **`Document.trim()` 放在保存之前** | 大画板留白在文档层面裁掉，输出的 PNG/JPG 才不会带一整圈空白；且因 `DONOTSAVECHANGES`，裁的是内存副本，磁盘原稿不变 |
| 5 | **提供 4 种裁切策略** | 透明留白用 `TRANSPARENT`、白底留白用 `TOPLEFT`、两者叠加用 `both`、完全不动用 `none`；一刀切会误裁或漏裁 |
| 6 | **格式可勾选而非固定双输出** | 有时只要 PNG 贴图，有时只要 JPG 上架；强行两个都出既费时又占盘 |
| 7 | **先存 PNG 再 `flatten()` 存 JPG** | PNG 必须保留 alpha；JPG 不支持透明，拼合时由 Photoshop 用背景色（默认白）填充，顺序反了 PNG 就没透明了 |
| 8 | JPG 前降到 8 位/通道、`quality=12` | JPG 不支持 16 位；质量 12 是 Photoshop 最高档，商品图不糊 |
| 9 | 32 位 HDR → 16 位；非 RGB → RGB | PNG 不支持 32 位浮点通道和 CMYK，不转换会直接保存失败 |
| 10 | `saveAs(..., true, Extension.LOWERCASE)` | 以副本另存、扩展名统一小写，避免 `.PNG` 在大小写敏感的引擎/CI 里读不到 |
| 11 | `finally { d.close(DONOTSAVECHANGES) }` | 无论成功失败都关闭且不保存：既保证原稿不被改，也避免批量转换后 Photoshop 里堆几十个大文档吃满内存 |
| 12 | 路径 `js_escape`（`\`→`\\`、`'`→`\'`） | Windows 反斜杠、中文 / 全角括号 / 带单引号的目录名会破坏 JS 字符串字面量，表现为「点了没反应」 |
| 13 | 输出写在源文件旁边、成对共用同一个 stem | 少一次「选输出目录」的决策；PNG 与 JPG 同名便于配对管理 |
| 14 | 同名自动 `_PNG` / `_PNG_2` | 反复重跑、增量补图时不会把上一版成果覆盖掉 |
| 15 | 文件夹递归 + 大小写不敏感去重 | 一次选根目录覆盖整个项目；混合输入容易把同一文件计两次 |
| 16 | 后台线程转换 + 每线程 `CoInitialize()` | COM 必须在初始化过的单元里调用；大文件转换不能冻结窗口；UI 只由主线程更新 |
| 17 | 单张失败继续下一张 | 20 张里有一张被占用/损坏，不该让整批白跑 |

逐条展开见 [`docs/DESIGN.md`](docs/DESIGN.md)，v1 原始字节证据见 [`docs/REVERSE_ENGINEERING.md`](docs/REVERSE_ENGINEERING.md)。

---

## 输出规则

以 `NTT72-D.psb` 为例（默认 PNG + JPG 都勾）：

| 情况 | 输出 |
| --- | --- |
| 首次 | `NTT72-D.png` + `NTT72-D.jpg` |
| 任一已存在 | `NTT72-D_PNG.png` + `NTT72-D_PNG.jpg` |
| 再跑一次 | `NTT72-D_PNG_2.png` + `NTT72-D_PNG_2.jpg` … 递增 |
| 只勾 PNG | 只生成 `NTT72-D.png` |
| 只勾 JPG | 只生成 `NTT72-D.jpg` |

- PNG 保留 alpha（透明背景）；JPG 拼合后用白底填充透明区域、质量 12。
- 裁切后画布变小，输出尺寸 = 内容边界尺寸；选「不裁切」则保持原始画布。
- 不生成日志文件或中间文件，不改动源 PSD/PSB。
- 中文、空格、全角括号、单引号路径均可处理（已转义）。

---

## 运行环境与前置条件

| 依赖 | 要求 |
| --- | --- |
| 操作系统 | Windows（Win10 / Win11）；COM + 桌面 GUI 决定它不能跑在 macOS / Linux |
| **Adobe Photoshop** | **必须已安装桌面完整版**，且 COM 组件 `Photoshop.Application` 可用 |
| v2 运行时 | 已打包进 EXE，不需要另装 Python / .NET 6/8 / VC++ 运行库 |
| v1 运行时 | .NET Framework 4.0+（Win10/11 自带 4.8） |
| 不需要 | 网络连接、管理员权限、注册表写入、配置文件 |

> **最容易忽略的一点**：工具不附带 Photoshop，也不能在没装 Photoshop 的机器上转换。这是第 1 条决策换来的准确性代价。首次使用建议先手动打开一次 Photoshop，把许可 / 初始化 / 字体等一次性弹窗处理掉。

---

## 能力边界与已知限制

- **大画板很吃资源**：1~2 GB 的 PSB 需要 Photoshop 真实加载，可能数分钟、占用大量内存与暂存盘；转换期间不要操作 Photoshop。
- **裁切依赖留白类型**：`TRANSPARENT` 只裁透明像素；**如果留白是白色/纯色而非透明，请改选「裁掉纯色边」或「先透明后纯色」**，否则看起来没裁动。这是策略差异，不是故障。
- **JPG 必然丢失透明**：这是格式本身限制，工具用白底填充（Photoshop 默认背景色）。
- **不提供 PNG 压缩 / 位深 / ICC 细调**：要精细控制请用 Photoshop 动作或另存对话框。
- **Photoshop 被占用会失败**：模态框或另一个脚本占用时 `DoJavaScript` 抛 COM 异常，提示后继续下一张。
- **未做数字签名**：PyInstaller 打包的 EXE 部分杀软可能误报；源码在 [`src/`](src/)，可自行打包核对。
- **真·多画板（Artboard）未专项验证**：工具按「一整张画布 + 四周留白」处理；若 PSD 内含多个 Artboard，建议先导出为单画布再试。
- **退出时弹 `Failed to remove temporary directory`（onefile 版）**：PyInstaller 单文件模式每次运行都要往临时目录解压，退出时删除被安全软件锁住的 DLL 就会弹这个模态框，把进程卡死。这是打包方式的问题，**不是转换失败**。根治办法是用 `--onedir` 模式（见下方[从源码构建](#从源码构建)）——目录版不解压临时文件，从物理上不会出现这个弹窗，且退出干净、启动更快。仓库里的单文件版保留给没有此问题的机器；本机已改用目录版。

---

## 从源码构建

v2 是 Python（tkinter + pywin32 + tkinterdnd2），用 PyInstaller 打包：

```bat
pip install -r requirements.txt
```

```bat
cd psd-to-png
python build_exe.py
```

产物写到 `dist/PSD-PSB一键转PNG-JPG.exe`（onefile，单文件分发用）。

**本机推荐 `--onedir`（根治临时目录弹窗）**：

```bat
python build_exe.py --onedir
```

产物是 `桌面\PSD一键转PNG工具\` 目录（内含 `PSDTool.exe` + `_internal`），然后把桌面/任务栏快捷方式指向 `PSDTool.exe`。目录版**不往 `%TEMP%` 解压**，因此不会弹 `Failed to remove temporary directory`，退出码干净为 0、启动也更快。实测 `rc=0`、`_MEI` 目录数前后不变。

> 2026-11-18 修复：窗口拖放从 `windnd` 改为 `tkinterdnd2/TkDND`，解决部分 64 位 Python/Tk 环境拖入文件时 `ucrtbase.dll / 0xc0000409` 进程级崩溃；窗口版 CLI 同时改为在无控制台时写 `%TEMP%\psdtool_cli.log`，并在 Photoshop 返回“应用程序正忙”时自动重试。

> 2026-11-19 v2.1：拖放只注册一个 TkDND 目标并明确返回复制动作，增加 `%TEMP%\psdtool_gui_error.log` 回调异常记录；界面新增“当前文件”动态进度和批量“总进度”。

> 注意：onedir 的程序名必须用 ASCII（`PSDTool`），中文名会让 bootloader 报 `Could not load PyInstaller's embedded PKG archive`——构建脚本里已按此处理。onefile 版仍用固定运行时临时目录 `C:\Users\Administrator\PSDToolRuntime`，避免每次在系统 `%TEMP%` 留下随机 `_MEI`。

开发态自检（不启动 Photoshop）：

```bat
python src/psd_to_png_gui.py --selftest
```

命令行批量（适合挂到脚本里）：

```bat
python src/psd_to_png_gui.py --cli "F:\杨平-产品图片\...\NTT72" --formats png jpg --trim transparent
```

`src/PsdToPng.cs` 是 **v1 原版 EXE 的行为等价重建**（C# / .NET 4.0），保留作历史参考；`src/build.bat` 可用 Windows 自带 `csc.exe` 编译它。两者关系见 [src/README.md](src/README.md)。

---

## 项目结构

```
psd-to-png/
├── README.md                          ← 本详情页
├── LICENSE                            ← MIT
├── build_exe.py                       ← v2 打包脚本（PyInstaller）
├── dist/
│   ├── PSD-PSB一键转PNG-JPG.exe        ← v2 推荐（支持 PSB / 裁切 / 双格式）
│   └── PSD一键转PNG.exe                ← v1 原版（13,312 bytes，未改动）
├── src/
│   ├── psd_to_png_gui.py              ← v2 完整源码
│   ├── PsdToPng.cs                    ← v1 行为等价重建（历史参考）
│   ├── build.bat                      ← v1 C# 编译脚本
│   └── README.md                      ← 源码说明与构建参数
└── docs/
    ├── DESIGN.md                      ← 17 条设计决策逐条展开
    └── REVERSE_ENGINEERING.md          ← v1 元数据 / 字符串 / JS 模板原始证据
```

---

## FAQ

**为什么 `.psb` 拖进去没反应？**
你用的是 v1（只认 `.psd`）。换 `dist/PSD-PSB一键转PNG-JPG.exe`，`.psd` 与 `.psb` 都支持。

**留白是白色的，默认裁不掉？**
默认「裁掉透明边」只处理透明像素。在「裁切方式」里改选「裁掉纯色边（按左上角像素颜色）」或「先裁透明边，再裁纯色边」。

**1~2 GB 的 PSB 会不会卡死？**
不是卡死，是 Photoshop 在真实加载大文档。状态栏会显示文件名与体积；请耐心等待，期间不要在 Photoshop 里点任何弹窗。内存/暂存盘不足时该张会失败并记入日志，不影响后续文件。

**会改坏我的 PSD / PSB 吗？**
不会。脚本以 `close(SaveOptions.DONOTSAVECHANGES)` 收尾，且所有修改（位深、模式、裁切、拼合）都发生在内存文档上，以副本方式另存。

**PNG 还有透明背景吗？**
有（选「不裁切」或任意裁切策略都保留 alpha）。JPG 没有透明，会用白底填充。

**能只出一种格式吗？**
能。取消勾选另一种即可；两个都取消会提示至少选一个。

**重跑会覆盖上次的图吗？**
不会，自动加 `_PNG`、`_PNG_2`… 后缀，且 PNG 与 JPG 共用同一后缀保持成对。

**提示「没有找到 Photoshop」？**
确认装的是桌面完整版而非看图软件；COM 组件需注册。修复安装或先手动启动一次 Photoshop 通常可解决。

**能脱离 Photoshop 用吗？**
不能，见「为什么这样做」第 1 条。改用 PSD 解析库会牺牲复杂图层效果的还原度，也处理不了 1~2 GB 的 PSB。

---

## License & 声明

[MIT](LICENSE) · Copyright (c) 2026 1187574340yangping-hash

Adobe、Photoshop、KeyShot 分别为其各自所有者的商标。本工具不随附、不修改、不分发 Photoshop，仅通过系统已注册的 COM 接口调用用户本机安装的软件；转换结果的准确性由本机 Photoshop 决定。使用本工具产生的图片请自行确认符合素材与平台规范。

v1 原始 EXE 校验值：`SHA-256 7e39aa6949d2275b3b32ab80f3951161701b017ebbc5326ef855850632dff2b5`
