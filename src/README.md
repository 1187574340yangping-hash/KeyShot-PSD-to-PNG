# src/ · 源码说明

## 这里是什么

`PsdToPng.cs` 是**依据 `dist/PSD一键转PNG.exe` 重建的行为等价参考实现**，不是随原工具分发的原始源码。

原程序只发布了一个 13,312 字节的 .NET Framework 4.0 WinForms EXE。本文件根据以下证据重建：

| 证据 | 来源 |
| --- | --- |
| 目标框架 `v4.0.30319`、程序集引用 | CLI 元数据头与 `#Strings` 堆 |
| `MainForm` / `AcceptPaths` / `StartConversion` / `ConvertFiles` / `UniquePng` / `JsEscape` / `DoJavaScript` / `FinalReleaseComObject` | 元数据方法名 |
| 全部界面文案、文件过滤器、`Photoshop.Application`、ExtendScript 模板 | `#US` 用户字符串堆 |
| `STAThread`、`Thread.SetApartmentState`、`BeginInvoke`、`COMException`、`Activator.CreateInstance`、`Directory.GetFiles(..., AllDirectories)`、`Distinct(OrdinalIgnoreCase)` | IL 引用的 API |

`docs/DESIGN.md` 记录了逐条设计决策；`docs/REVERSE_ENGINEERING.md` 记录了原始字节证据，便于核对。

## 构建

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

## 与原 EXE 的关系

- **行为对齐**：三种输入入口、递归收集、大小写不敏感去重、输出命名、Photoshop JS 脚本、不保存原 PSD、COM 释放、错误文案。
- **实现差异**：UI 控件坐标与配色是近似还原；异常分支改为「单张失败继续下一张」，并在 `finally` 里统一恢复按钮状态；`AcceptPaths` 增加了无权限路径的跳过保护。
- **原版优先**：日常使用请直接运行 `dist/PSD一键转PNG.exe`，那是经过实际使用的原版。重建源码主要用于理解原理、修改规则和二次开发。

> 重建源码未在本次环境中执行编译验证（环境策略不允许调用任意 C# 编译器）。如编译报错，请以 `docs/REVERSE_ENGINEERING.md` 中的原始 JS 模板与文案为准修正。
