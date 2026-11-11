# 设计说明 · 为什么这样做

本文把 [README 里的 14 条决策](../README.md#为什么这样做) 逐条展开。结论先说：

> **这个工具的唯一目标是「让 Photoshop 本人替你把 PSD 另存成 PNG」，其余所有设计都是在保证这件事在批量、无人值守、中文路径、原稿安全的前提下不翻车。**

---

## 1. 为什么调 Photoshop COM，而不是自己解析 PSD

这是最关键、也最容易被误解的取舍。

自己解析 PSD 听起来更"独立"，但 PSD 不是一种扁平位图格式：它可能有数十个图层、混合模式、智能对象、调整图层、图层蒙版、剪贴蒙版、文本图层、矢量形状、图层效果，还可能是 16 位或 32 位/通道、带 ICC 配置。开源解析器通常只能取到**合成预览图**，遇到智能对象或复杂混合就会和 Photoshop 里看到的不一样。

工具的做法是：

```csharp
Type t = Type.GetTypeFromProgID("Photoshop.Application");
object app = Activator.CreateInstance(t);
dynamic ps = app;
ps.DoJavaScript(script);
```

`DoJavaScript` 执行的是 Photoshop 原生 ExtendScript，等于让 Photoshop 用它自己的渲染管线打开文档、用它自己的 PNG 编码器写出文件。**合成结果与人工「另存为 PNG」逐像素一致**，这是任何第三方解析器都给不了的保证。

代价也明确写进了文档：**必须有 Photoshop、必须有 Windows、转换速度取决于 Photoshop 打开文档的速度**。对目标用户（本来就用 Photoshop / KeyShot 的设计师）这个代价为零，对服务器批量场景则不可接受——所以 README 里直接把它列为不适用场景，而不是含糊地说"推荐安装"。

## 2. 为什么第一句 JS 就是 `displayDialogs = DialogModes.NO`

批处理的头号杀手不是崩溃，是**挂起**。Photoshop 打开 PSD 时可能弹：

- 「缺失字体，是否替换」
- 「嵌入的颜色配置文件与工作空间不匹配」
- 「此文档包含无法读取的增效工具数据」
- 「最大化文件兼容性」另存确认

任意一个弹窗都会让 `DoJavaScript` 一直不返回，用户看到的就是"卡在第 3 张"。`DialogModes.NO` 把全部对话框压成默认行为，是无人值守的前提。

## 3. 为什么 32 位/通道要降到 16 位

HDR 渲染输出（KeyShot、V-Ray、Blender 线性 EXR 转存等）常是 32 位浮点通道。PNG 规范只支持 8/16 位整数，直接 `saveAs` 会失败或静默截断。脚本里的判断：

```javascript
if (d.bitsPerChannel == BitsPerChannelType.THIRTYTWO)
    d.bitsPerChannel = BitsPerChannelType.SIXTEEN;
```

注意这一步改的是**内存里的文档对象**，而文档最终以 `DONOTSAVECHANGES` 关闭，所以磁盘上的 HDR 原稿完全不受影响。

## 4–5. 为什么 `interlaced=false` 与 `Extension.LOWERCASE`

```javascript
var o = new PNGSaveOptions();
o.interlaced = false;
d.saveAs(new File(pngPath), o, true, Extension.LOWERCASE);
```

- **不交错**：交错（Adam7）只对网页渐进加载有意义，对贴图、材质、商品预览没有收益，反而增大体积、拖慢部分 3D 引擎和 CDN 的解码。
- **`true`（asCopy）**：以副本方式另存，当前文档仍指向原始 PSD，不会把文档"变成" PNG 再被后续操作误保存。
- **`LOWERCASE`**：强制小写 `.png`。Windows 上无所谓，但素材常常要同步到 Linux 构建机、CI、KeyShot 渲染农场或对象存储，`.PNG` 与 `.png` 在那里是两个文件，统一小写能省掉一类极难排查的"贴图找不到"。

## 6. 为什么用 `finally` 包住 `close(DONOTSAVECHANGES)`

```javascript
var d = null;
try {
    d = app.open(new File(psdPath));
    ...
    d.saveAs(...);
} finally {
    if (d) d.close(SaveOptions.DONOTSAVECHANGES);
}
```

两个目的：

1. **原稿安全**：`open` 之后做了位深修改，如果最后以 `SAVECHANGES` 关闭就等于把用户 PSD 改了。`DONOTSAVECHANGES` 明确拒绝写回。
2. **内存与稳定性**：`finally` 保证 `open` 成功但 `saveAs` 抛错时文档仍被关闭。否则批量转换 50 张里失败 10 张，Photoshop 里就积 10 个打开的大文档，后续转换越来越慢直至内存耗尽。

## 7. 为什么必须有 `JsEscape`

工具把 Windows 路径拼进 JS 字符串字面量：

```javascript
d = app.open(new File('D:\素材\面包.psd'));
```

在 JavaScript 里 `\素` 不是合法转义，轻则路径错、重则整段脚本语法错误——而 COM 调用往往只回一个笼统异常，用户看到的是"点了没反应"。中文目录、空格目录、`O'Brien` 这类带单引号的目录名同样会打断字符串。

```csharp
s.Replace("\\", "\\\\").Replace("'", "\\'")
```

两条替换覆盖全部 Windows 路径场景。这也是为什么不能图省事用字符串插值直接拼脚本。

## 8–9. 为什么输出在原目录、且永不覆盖

**同目录**：3D 贴图和商品图的工作习惯是"素材和它的导出图放一起"。让用户每次选输出目录，等于给一个 20 张图的任务加 20 次决策，还容易把图导到别处导致 KeyShot 断链。

**不覆盖**：批量任务天然会被重跑——补渲了几张、改了一版颜色、上次中途失败。如果直接覆盖，用户上一版交付的图就悄悄没了。命名阶梯：

```
面包.png  →  面包_PNG.png  →  面包_PNG_2.png  →  ...
```

`_PNG` 这个后缀也出现在完成提示里（「若同名文件已存在，会自动加 _PNG」），让用户知道文件去哪了，而不是以为转换失败。

## 10. 为什么递归 + 大小写不敏感去重

- **递归**：KeyShot 项目常按 `项目/材质/`、`项目/输出/`、`项目/版本2/` 分目录。只扫一层等于没做批量。
- **去重**：四种入口可能混合使用（先拖了几个文件，又选了根目录），Windows 文件系统大小写不敏感，`Distinct(StringComparer.OrdinalIgnoreCase)` 保证 `A.PSD` 和 `a.psd` 不会被 Photoshop 打开两遍、生成两个 PNG。

`AcceptPaths` 对单路径的读取包了 `try/catch`：无权限目录、断开的网络盘不应该让整批任务中断。

## 11. 为什么是后台 STA 线程

三个约束叠在一起：

| 约束 | 后果 |
| --- | --- |
| Photoshop COM 是单元线程模型 | 必须在 `ApartmentState.STA` 线程创建和调用 |
| 打开大 PSD 可能几十秒 | 放主线程会整个窗口假死 |
| WinForms 控件非线程安全 | 状态文字、进度条只能用 `BeginInvoke` 回主线程改 |

所以结构是：主线程收集文件 → 禁用按钮 → 起后台 STA 线程跑 `ConvertFiles` → 线程里每段状态更新都用 `BeginInvoke` 投递 → 完成后 `finally` 恢复按钮。`Thread.IsBackground = true` 保证用户直接关窗口不会留下孤儿进程。

## 12. 为什么 `COMException` 要单独捕获

`DoJavaScript` 最常见的失败不是文件坏，而是 **Photoshop 正忙**：用户在前台点了某个菜单、有个模态框没关、另一个脚本正在跑、或者增效工具弹了窗。这时返回的是 `COMException`，笼统提示「转换失败」对用户毫无帮助。

单独分支给出「Photoshop 忙或无法读取文件：…」，用户立刻知道去把 Photoshop 窗口里的弹窗点掉。其余 `Exception` 走「转换失败：…」，两者都**继续处理下一张**，不让一张坏图终止整批。

## 13. 为什么显式 `FinalReleaseComObject`

`Activator.CreateInstance` 拿到的 RCW 若只靠 GC 回收，Photoshop 进程可能在转换结束后继续驻留。`Marshal.FinalReleaseComObject(app)` 在 `finally` 里立即释放引用，配合第 6 条的文档关闭，保证一轮批量结束后 Photoshop 回到干净状态。

## 14. 为什么做成 13 KB 的单 EXE

目标用户是设计师，不是运维。这个体量意味着：

- 只引用系统程序集（`System` / `System.Core` / `System.Drawing` / `System.Windows.Forms` / `Microsoft.CSharp`），没有第三方 DLL 要一起分发；
- 不写注册表、不装服务、不建配置目录，复制到 U 盘或共享盘就能给同事；
- .NET Framework 4.0 是 Win10/11 系统组件，不需要引导安装 .NET 6/8 运行时；
- 无网络请求，素材不出本机——对未发布的商品图、客户稿很重要。

---

## 与相邻工具的关系

```
WorkBuddy 图生 3D (hy-3d)
        ↓ OBJ + PBR 贴图（PSD/PNG 混合）
本工具：PSD ──→ PNG（保留尺寸与 alpha）
        ↓
KeyShot-PBR-Texture-Tool：重整 OBJ/MTL/贴图 → KeyShot 可直接导入
        ↓
KeyShot：运行脚本批量生成 .bip
```

本工具只负责**格式转换这一小段**，不做材质重命名、不生成 MTL、不碰 `.bip`。职责收窄是刻意的：每一环都能单独替换或跳过。

## 明确的边界

- **离线 ≠ 独立**：不联网，但强依赖本机 Photoshop。
- **准确 ≠ 快速**：视觉 1:1 的代价是每张都要真实打开文档。
- **不改原稿 ≠ 不留痕**：会在源目录新建 PNG，这是用户预期内的产物。
- **启发式之外**：本工具不做任何 AI 判定、不做图像内容分析、不上传任何数据。
