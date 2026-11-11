# 逆向证据 · 原始 EXE 里到底有什么

本文件记录从 `dist/PSD一键转PNG.exe` 直接读出的字节证据，用于核对 [README](../README.md)、[DESIGN.md](DESIGN.md) 与 [重建源码](../src/PsdToPng.cs) 的每一条结论。

> 说明：只做了**静态读取**（PE 头、CLI 元数据、字符串堆、方法名与引用 API），未运行程序、未连接 Photoshop、未反编译第三方代码。原始 EXE 未作任何修改。

---

## 1. 文件指纹

| 项 | 值 |
| --- | --- |
| 路径 | 桌面 `PSD一键转PNG.exe`，仓库内 `dist/PSD一键转PNG.exe` |
| 大小 | 13,312 bytes |
| SHA-256 | `7e39aa6949d2275b3b32ab80f3951161701b017ebbc5326ef855850632dff2b5` |
| DOS/PE 头 | `MZ` … `PE\0\0`，Machine `0x8664` |
| 运行时 | CLI 元数据版本串 `v4.0.30319` → .NET Framework 4.0 / C# WinForms |

只有 `.text` 与 `.rsrc` 两个节，没有 `.reloc` 之外的大量数据段；文件里没有 PyInstaller `MEI` cookie、没有内嵌 Python、没有打包的第三方 DLL —— 这直接排除了"Python 打包工具"的可能，也解释了为什么只有 13 KB。

```
PE off 0x80   machine 0x8664
  .text    VA=0x2000 VSZ=0x2aec ROFF=0x200  RSZ=0x2c00
  .rsrc    VA=0x6000 VSZ=0x4f0  ROFF=0x2e00 RSZ=0x600

Metadata root at 0xf08, version string: v4.0.30319
  stream #~         off=0xf74  size=2228
  stream #Strings   off=0x1828 size=2840
  stream #US        off=0x2340 size=1460
  stream #GUID      off=0x28f4 size=16
  stream #Blob      off=0x2904 size=1000
```

## 2. `#US` 用户字符串堆（完整导出）

这是最直接的证据：C# 里所有字符串字面量都放在 `#US` 堆，以压缩长度 + UTF-16LE 存储。

```
[  39] Photoshop PSD|*.psd
[  43] 选择包含 KeyShot PSD 的文件夹
[  11] *.psd
[  23] PSD 一键转 PNG
[  37] Microsoft YaHei UI
[  39] KeyShot PSD 一键转 PNG
[  69] 拖入一个或多个 PSD；PNG 保存在原文件旁边，不修改原 PSD。
[  51] ↓\r\n拖入 PSD 文件\r\n自动保留尺寸和透明背景
[  19] 选择 PSD 文件
[  11] 选择文件夹
[  13] 等待拖入文件
[   9] .psd
[  27] 没有检测到 PSD 文件。
[   3] \
[   5] \\
[   3] '
[   5] \'
[   9] .png
[  17] _PNG.png
[  11] _PNG_
[  11] 正在转换 
[   3] /
[   3] ：
[  21] （大文件请耐心等待）
[  15] 完成：已导出 
[  53]  个 PNG。若同名文件已存在，会自动加 _PNG。
[  39] Photoshop 忙或无法读取文件：
[  11] 转换失败：
[  43] Photoshop.Application
[  69] 没有找到 Photoshop。请先确认 Photoshop 已安装。
[ 145] app.displayDialogs=DialogModes.NO; var d=null; try{d=app.open(new File('
[ 327] '));if(d.bitsPerChannel==BitsPerChannelType.THIRTYTWO)d.bitsPerChannel=BitsPerChannelType.SIXTEEN;var o=new PNGSaveOptions();o.interlaced=false;d.saveAs(new File('
[ 169] '),o,true,Extension.LOWERCASE);}finally{if(d)d.close(SaveOptions.DONOTSAVECHANGES);}
[  25] DoJavaScript
[  43] FinalReleaseComObject
```

由此可直接确认：

- 程序标题是 **「KeyShot PSD 一键转 PNG」**，文件夹选择框文案是 **「选择包含 KeyShot PSD 的文件夹」** → 定位就是 KeyShot / 3D 贴图工作流。
- 提示语明确承诺 **「PNG 保存在原文件旁边，不修改原 PSD」**、**「自动保留尺寸和透明背景」**。
- 输出命名用 `.png` / `_PNG.png` / `_PNG_` 三级阶梯。
- 转义表只有 `\`→`\\` 和 `'`→`\'` 两条。
- 两种错误文案分别对应 `COMException` 与通用异常。
- 依赖的 ProgID 是 `Photoshop.Application`，调用方法是 `DoJavaScript`，收尾是 `FinalReleaseComObject`。

## 3. 还原出的 ExtendScript（三段 `#US` 拼接）

```javascript
app.displayDialogs = DialogModes.NO;
var d = null;
try {
    d = app.open(new File('<PSD 路径，已转义>'));
    if (d.bitsPerChannel == BitsPerChannelType.THIRTYTWO)
        d.bitsPerChannel = BitsPerChannelType.SIXTEEN;
    var o = new PNGSaveOptions();
    o.interlaced = false;
    d.saveAs(new File('<PNG 路径，已转义>'), o, true, Extension.LOWERCASE);
} finally {
    if (d) d.close(SaveOptions.DONOTSAVECHANGES);
}
```

DESIGN.md 的第 2–6 条决策全部来自这一段，没有额外推测。

## 4. 元数据里出现的方法名 / 类名

```
<Module>  MainForm  Program(隐含)
dropText  status  dropPanel  ProgressBar  progress
Button  chooseFiles  chooseFolder
DragEventArgs  OnDragEnter  OnDragDrop
IEnumerable`1  AcceptPaths  List`1  StartConversion
Action  JsEscape  UniquePng  ConvertFiles  startupPaths
<>c__DisplayClass6 / 9 / 18 / 1a / 1c / 1e
<StartConversion>b__8   <ConvertFiles>b__10 … b__15
```

- `startupPaths` + `add_Shown` → 支持**命令行传入路径**（把文件拖到 EXE 图标上启动）。
- `<>c__DisplayClass*` 与 `<ConvertFiles>b__*` → 闭包捕获循环变量，用 `BeginInvoke` 回主线程更新 UI。
- `AcceptPaths` / `StartConversion` / `ConvertFiles` / `UniquePng` / `JsEscape` → 重建源码里的同名方法即由此而来。

## 5. 引用的 API（由 `#Strings` / MemberRef 推断调用链）

```
System.Windows.Forms      Form / Label / Button / ProgressBar / OpenFileDialog /
                          FolderBrowserDialog / DragEventArgs / DataFormats.FileDrop /
                          Application.EnableVisualStyles / SetBounds / Dock / TextAlign
System.Drawing            Color.FromArgb / Font / FontStyle / ContentAlignment
System.IO                 Directory.GetFiles(path, "*.psd", SearchOption.AllDirectories)
                          Path.GetExtension / GetFileNameWithoutExtension /
                          GetDirectoryName / GetFileName / Combine
System.Linq               Enumerable.Distinct / ToList
System.Collections.Generic StringComparer.OrdinalIgnoreCase / IEqualityComparer`1
System.Threading          ThreadStart / Thread / IsBackground /
                          SetApartmentState(ApartmentState.STA)
System.Runtime.InteropServices  COMException / Marshal.FinalReleaseComObject /
                          STAThreadAttribute
System                    Type.GetTypeFromProgID / Activator.CreateInstance /
                          Delegate.BeginInvoke / Control.IsDisposed / IsHandleCreated
Microsoft.CSharp.RuntimeBinder  CSharpBinderFlags / InvokeMember / CallSite
                          → dynamic 调用 DoJavaScript
```

`IEnumerable` + `Distinct(StringComparer.OrdinalIgnoreCase)` 确认了去重规则；`SearchOption` + `GetFiles` 确认了文件夹递归；`SetApartmentState` 确认了 STA；`Microsoft.CSharp.RuntimeBinder` 确认了用 `dynamic` 而不是手写反射。

## 6. 由证据到结论的对应关系

| 结论 | 证据 |
| --- | --- |
| 不是自解析 PSD，而是驱动 Photoshop | `Photoshop.Application` + `DoJavaScript` + ExtendScript 模板 |
| 关掉全部弹窗 | `displayDialogs=DialogModes.NO` |
| 处理 32 位 HDR | `BitsPerChannelType.THIRTYTWO → SIXTEEN` |
| 不修改原稿 | `DONOTSAVECHANGES` + 提示语「不修改原 PSD」 |
| 输出同目录、不覆盖 | `.png` / `_PNG.png` / `_PNG_` + 提示语「保存在原文件旁边」 |
| 中文/单引号路径可用 | `#US` 里 `\`、`\\`、`'`、`\'` 四条转义片段 |
| 三种（四种）输入入口 | `OnDragDrop` / `OpenFileDialog(Multiselect)` / `FolderBrowserDialog(AllDirectories)` / `startupPaths` |
| 后台线程不卡界面 | `Thread` + `SetApartmentState` + `BeginInvoke` |
| 错误分类提示 | 两条不同的错误文案 + `COMException` 引用 |

## 7. 未做的推断

为免过度解读，以下内容**没有**写进详情页：

- 原始窗口的精确控件坐标、配色数值、字体大小（只知用了 `Microsoft YaHei UI` 和 `Color.FromArgb`）。
- 原版在单张失败后是否继续下一张（重建源码选择了继续，并在 [src/README.md](../src/README.md) 标注为差异）。
- 原版的具体编译选项、是否签名、确切的 C# 编译器版本（只能确认目标框架 `v4.0.30319`）。
- 作者、构建时间、原始源码所在地。

原始 EXE 保持字节不变，可随时用任意 .NET 反编译器（ILSpy / dnSpy）复核。
