# 设计说明 · 为什么这样做

本文把 [README 里的 17 条决策](../README.md#为什么这样做) 逐条展开。结论先说：

> **这个工具只做一件事：让 Photoshop 本人把 PSD / PSB 打开、裁到内容边界、另存成你要的格式。其余所有设计，都是为了保证这件事在批量、无人值守、中文路径、1~2 GB 大文档、原稿安全的前提下不翻车。**

版本关系：v1（`dist/PSD一键转PNG.exe`，13 KB .NET）只认 `.psd`、不裁切、只出 PNG；v2（`dist/PSD-PSB一键转PNG-JPG.exe`）在同一条 COM 链路上补齐了 **PSB 支持 / 自动裁切 / 格式可勾选**。v1 的原始字节证据见 [`REVERSE_ENGINEERING.md`](REVERSE_ENGINEERING.md)，本文讨论的是 v2 的当前实现（`src/psd_to_png_gui.py`）。

---

## 1. 为什么调 Photoshop COM + ExtendScript，而不是自己解析 PSD / PSB

最关键、也最容易被误解的取舍。

PSD / PSB 不是扁平位图：可能有数十个图层、混合模式、智能对象、调整图层、蒙版、剪贴蒙版、文本 / 矢量图层、图层效果，还可能是 16 位或 32 位每通道、带 ICC 配置。开源解析器通常只能取到**合成预览图**，遇到智能对象或复杂混合就会和 Photoshop 里看到的不一样。

到了 v2 面对的 PSB，这条路更走不通：单个文件 1~2 GB，画布可能超过 30,000 像素，纯 Python 解析要么吃不下内存，要么直接不支持。

```python
app = win32com.client.Dispatch("Photoshop.Application")
app.DoJavaScript(script)
```

`DoJavaScript` 执行 Photoshop 原生 ExtendScript，等于让 Photoshop 用**它自己的渲染管线**打开文档、用它自己的编码器写出文件。合成结果与人工「另存为」逐像素一致，这是任何第三方解析器给不了的保证。

代价明确写进 README：**必须有 Photoshop、必须有 Windows、速度取决于 Photoshop 打开文档的速度**。对目标用户（本来就用 Photoshop / KeyShot 的设计师）代价为零；对服务器 / CI 则不可接受，因此直接列为不适用场景，而不是含糊地说「推荐安装」。

## 2. 为什么把扩展名放宽到 `.psb`

用户现场：`F:\杨平-产品图片\...\NTT72\` 里 `NTT72-D` 其实是 `NTT72-D.psb`（1.65 GB），同目录还有 5 个 1~2 GB 的 `.psb`。v1 拖不进去、也导不出 PNG，**不是文件坏了**，而是三处都写死 `*.psd`：文件对话框过滤器、拖入时的扩展名判断、文件夹递归。

PSB（Photoshop Large Document Format）是 Photoshop 在文档超过 2 GB 或 30,000 像素时自动使用的原生格式，渲染输出、大画板拼版几乎都是它。对 `app.open(new File(...))` 而言 PSD 与 PSB 一视同仁，区别只在扩展名。

v2 统一用 `SUPPORTED_EXTS = {".psd", ".psb"}`，收集时把后缀 `lower()` 后比对，大小写不敏感。

## 3. 为什么第一句 JS 就是 `displayDialogs = DialogModes.NO`

批处理的头号杀手不是崩溃，是**挂起**。Photoshop 打开文档时可能弹：缺失字体、颜色配置文件与工作空间不匹配、增效工具数据无法读取、最大化文件兼容性确认……任意一个都会让 `DoJavaScript` 一直不返回，用户看到的就是「卡在第 3 张」。大画板 PSB 更容易触发这些弹窗，所以这一句在 v2 里更不可省。

## 4. 为什么 `trim()` 放在保存之前，而不是导出后再裁

大画板里内容只占中间一小块、四周一大圈透明或白色留白。如果按原画布导出，得到的 PNG 里 80% 是空白，贴图、商品图、详情页都用不了。

两种做法：

- **导出后再用图像库裁**：要先把 1~2 GB 的合成图读进内存，且「什么是留白」得自己判断，容易把浅色产品边缘一起裁掉。
- **在文档层面 `Document.trim()`**：Photoshop 自己判断边界，裁完再编码，输出尺寸直接就是内容边界尺寸，零额外解码。

```javascript
d.trim(TrimType.TRANSPARENT, true, true, true, true);
```

更关键的是：所有修改都发生在内存文档上，文档最终以 `DONOTSAVECHANGES` 关闭，**磁盘原稿一个字节都不动**。

## 5. 为什么提供 4 种裁切策略，而不是固定一种

留白的性质不一样，一刀切必然翻车：

| 策略 | ExtendScript | 适用 |
| --- | --- | --- |
| `transparent` | `trim(TRANSPARENT)` | 四周透明（默认，渲染输出、抠图稿常见） |
| `topleft` | `trim(TOPLEFT)` | 白底 / 纯色留白，按左上角像素颜色裁 |
| `both` | 先 `TRANSPARENT` 再 `TOPLEFT` | 透明边 + 纯色边叠加，最彻底 |
| `none` | 不调用 `trim` | 需要保留原始画布尺寸（贴图对齐、版式定位） |

`TRANSPARENT` 只认透明像素；**白底留白用默认策略看起来「没裁动」，这是策略差异不是故障**——所以 README 的 FAQ 与「能力边界」里都专门点名了这一点，并给出改选建议。工具做的是像素级裁边，不是主体内容识别。

## 6. 为什么格式可勾选，而不是固定同时出两份

最初的误解是「保留一份 JPG」，一度做成了固定双输出。实际场景是：有时只要 PNG 贴图，有时只要 JPG 上架，有时两者都要。固定两个都出既费时又占盘——对 1~2 GB 的 PSB，重复编码一次就是几分钟。

因此 v2 把格式做成 `FORMATS = ("png", "jpg")` 的勾选，`build_script()` 只为勾中的格式生成保存代码，`unique_outputs()` 也只为勾中的格式占位。两个都不勾则直接提示至少选一个。

## 7. 为什么先存 PNG，再 `flatten()` 存 JPG

顺序不能反：

```javascript
var po = new PNGSaveOptions();
po.interlaced = false;
d.saveAs(new File(pngPath), po, true, Extension.LOWERCASE);
d.flatten();   // 这一步之后透明就没了
var jo = new JPEGSaveOptions();
jo.quality = 12;
d.saveAs(new File(jpgPath), jo, true, Extension.LOWERCASE);
```

PNG 必须保留 alpha；JPG 不支持透明，`flatten()` 会用 Photoshop 背景色（默认白）填充。若先拼合，PNG 就变成白底了。

## 8. 为什么 JPG 前要降到 8 位/通道、`quality=12`

- **位深**：JPG 规范只支持 8 位每通道，而文档可能刚从 32 位降到 16 位（见第 9 条），不降会保存失败。
- **质量**：`quality=12` 是 Photoshop 最高档，商品图、详情页放大看不糊；同时设 `embedColorProfile=true`、`formatOptions=STANDARDBASELINE` 保证色彩与兼容性。

## 9. 为什么 32 位 → 16 位、非 RGB → RGB

```javascript
if (d.bitsPerChannel == BitsPerChannelType.THIRTYTWO)
    d.bitsPerChannel = BitsPerChannelType.SIXTEEN;
if (d.mode != DocumentMode.RGB)
    d.changeMode(ChangeMode.RGB);
```

HDR 渲染输出常是 32 位浮点通道，PNG 只支持 8/16 位整数；CMYK / 灰度 / Lab 模式下 PNG 与 JPG 的编码也会失败或偏色。v1 只处理了位深，v2 补了 `changeMode(RGB)`——大画板拼版常从印刷稿转来，这一步能避免「点了没反应」。同样，改的都是内存文档，不写回原稿。

## 10. 为什么 `saveAs(..., true, Extension.LOWERCASE)`

- **`true`（asCopy）**：以副本另存，当前文档仍指向原始 PSD/PSB，不会把文档「变成」 PNG 再被后续操作误保存。
- **`LOWERCASE`**：强制小写 `.png` / `.jpg`。Windows 上无所谓，但素材常同步到 Linux 构建机、CI、KeyShot 渲染农场或对象存储，`.PNG` 与 `.png` 在那里是两个文件。
- **不交错**：`interlaced=false`，Adam7 只对网页渐进加载有意义，对贴图 / 商品预览反而增大体积、拖慢解码。

## 11. 为什么用 `finally` 包住 `close(DONOTSAVECHANGES)`

```javascript
var d = null;
try {
    d = app.open(new File(srcPath));
    ...
} finally {
    if (d) d.close(SaveOptions.DONOTSAVECHANGES);
}
```

两个目的：

1. **原稿安全**：`open` 之后做了位深、模式、裁切、拼合等一系列修改，若以 `SAVECHANGES` 关闭就等于把用户 1~2 GB 的原稿改了。
2. **内存与稳定性**：`finally` 保证 `open` 成功但 `saveAs` 抛错时文档仍被关闭。否则批量 20 张里失败几张，Photoshop 里就积着几个 GB 级的打开文档，后续越来越慢直至暂存盘爆掉。

## 12. 为什么必须有 `js_escape`

工具把 Windows 路径拼进 JS 字符串字面量：

```javascript
d = app.open(new File('F:\杨平-产品图片\01-产品图片（产品系列分类）\NTT72\NTT72-D.psb'));
```

`\杨` 不是合法转义，轻则路径错、重则整段脚本语法错误；而 COM 调用往往只回一个笼统异常，用户看到的是「点了没反应」。中文、空格、全角括号、`O'Brien` 这类带单引号的目录名同样会打断字符串。

```python
def js_escape(s):
    return str(s).replace("\\", "\\\\").replace("'", "\\'")
```

两条替换覆盖 Windows 路径场景。这也是为什么不能图省事用 f-string 直接拼脚本。

## 13. 为什么每个源文件使用同名输出文件夹

**同名文件夹**：3D 贴图、商品图仍与源文件放在同一个项目位置，但每个 PSD/PSB 的 JPG、可选 PNG 和 `图层列表.txt` 集中在独立目录，避免批量任务把源目录铺满图片。

**图层清单**：Photoshop 递归读取图层组和普通图层名称，按缩进层级写入文本；图片仍是整张画布的合成结果，不逐层导出。

## 14. 为什么同名自动 `_PNG` / `_PNG_2`，而不是覆盖

批量任务天然会被重跑：补渲了几张、改了一版颜色、上次中途失败。直接覆盖会把上一版已交付的图悄悄弄丢。

```
NTT72-D.png/.jpg → NTT72-D_PNG.png/.jpg → NTT72-D_PNG_2.png/.jpg → ...
```

判定是**整组**的：只要候选组里任一文件已存在，就整组跳到下一档，保证 PNG 与 JPG 永远共用同一后缀、不会错配。上限 10000 档，超了明确报错让用户清理目录，而不是无限循环。

## 15. 为什么文件夹递归 + 大小写不敏感去重

- **递归**：项目常按 `项目/材质/`、`项目/输出/`、`项目/版本2/` 分目录，只扫一层等于没做批量。
- **去重**：四种入口（拖窗口 / 拖 EXE 图标 / 多选文件 / 选文件夹）可能混用；Windows 文件系统大小写不敏感，`A.PSD` 与 `a.psd` 不能被 Photoshop 打开两遍。收集后按解析后的绝对路径归一化去重。
- **容错**：单路径读取包 `try/except`，无权限目录、断开的网络盘不该让整批中断。

## 16. 为什么后台线程转换 + 每线程 `CoInitialize()`

三个约束叠在一起：

| 约束 | 后果 |
| --- | --- |
| Photoshop COM 是单元线程模型 | 每个调用线程都要自己 `pythoncom.CoInitialize()` / `CoUninitialize()` |
| 打开 1~2 GB 的 PSB 可能数分钟 | 放主线程整个窗口假死 |
| tkinter 控件非线程安全 | 状态文字、进度条只能用 `root.after(0, fn)` 投递回主线程 |

结构：主线程收集文件 → 禁用按钮 → 起 `threading.Thread(daemon=True)` 跑 worker → worker 里每张先 `CoInitialize()`、结束 `CoUninitialize()` → 所有 UI 更新用 `after(0, ...)` → 完成后恢复按钮。`daemon=True` 保证用户直接关窗口不留孤儿进程。CLI 模式走同一套 `convert_one()`，逻辑不重复。

## 17. 为什么单张失败继续下一张

20 张里有一张被占用、损坏或暂存盘不足，不该让整批白跑。worker 对每张单独 `try/except`，失败记入日志与失败计数，继续下一张；`convert_one()` 的 `finally` 里释放 COM 引用并 `CoUninitialize()`，保证失败后线程状态干净。错误文案区分「Photoshop 忙或无法读取文件」「请先确认 Photoshop 已安装」「转换失败」，用户能直接知道该去点掉哪个弹窗。

---

## v1 → v2 改了什么

| | v1（.NET / C#） | v2（Python / tkinter） |
| --- | --- | --- |
| 输入 | 只 `.psd` | `.psd` + `.psb` |
| 裁切 | 无 | 4 种策略，保存前 `trim()` |
| 输出 | 只 PNG | PNG / JPG 可勾选，成对 |
| 模式 / 位深 | 32 位 → 16 位 | 再加非 RGB → RGB、JPG 前降 8 位 |
| GUI | WinForms + `BeginInvoke` + STA 线程 | tkinter + `after(0, ...)` + 后台线程 |
| 拖放 | WinForms 原生 | `tkinterdnd2` / TkDND，不可用时降级为按钮选择；避免 `windnd` 在部分 64 位环境触发原生崩溃 |
| 分发 | 13 KB 单 EXE（.NET 4.0） | ~11 MB 单 EXE（PyInstaller onefile，自带运行时） |

v1 的原始二进制、逆向证据与行为等价重建源码（`src/PsdToPng.cs`）全部保留，方便对照与二次开发。

## 与相邻工具的关系

```
WorkBuddy 图生 3D (hy-3d)
        ↓ OBJ + PBR 贴图
本工具：PSD / PSB ──→ PNG / JPG（裁到内容边界）
        ↓
KeyShot-PBR-Texture-Tool：重整 OBJ/MTL/贴图 → KeyShot 可直接导入
        ↓
KeyShot：运行脚本批量生成 .bip
```

本工具只负责**格式转换 + 裁边这一小段**，不做材质重命名、不生成 MTL、不碰 `.bip`。职责收窄是刻意的：每一环都能单独替换或跳过。

## 明确的边界

- **离线 ≠ 独立**：不联网，但强依赖本机 Photoshop。
- **准确 ≠ 快速**：视觉 1:1 的代价是每张都要真实打开文档；1~2 GB 的 PSB 可能数分钟。
- **裁切是策略，不是内容识别**：`trim()` 按透明 / 左上角像素颜色裁，不做主体检测；真·多画板（Artboard）未专项验证。
- **不改原稿 ≠ 不留痕**：会在源目录新建与 PSD/PSB 同名的输出文件夹，里面保存 JPG、可选 PNG 和图层清单。
- **不做 AI 判定、不做图像内容分析、不上传任何数据**。
