// PsdToPng.cs
// -----------------------------------------------------------------------------
// 行为等价的参考实现（reconstructed reference implementation）
//
// 说明：本文件不是随工具一起分发的原始源码。原始程序只有一个 13 KB 的
//       .NET Framework 4.0 WinForms EXE（见 ../dist/）。本文件依据该 EXE 的
//       CLI 元数据、方法名、#US 字符串堆（界面文案、Photoshop JS 模板）与
//       引用的 API 重建，用于说明工作原理和方便二次开发。
//
// 构建：见同目录 build.bat（需要 .NET Framework 4.0+ 自带的 csc.exe）。
// 运行：需要 Windows + 已安装且已注册 COM 的 Adobe Photoshop。
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace PsdToPng
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(args));
        }
    }

    internal sealed class MainForm : Form
    {
        private readonly Label dropPanel;
        private readonly Label status;
        private readonly ProgressBar progress;
        private readonly Button chooseFiles;
        private readonly Button chooseFolder;
        private readonly string[] startupPaths;

        public MainForm(string[] startupPaths)
        {
            this.startupPaths = startupPaths ?? new string[0];

            Text = "KeyShot PSD 一键转 PNG";
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(560, 280);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(245, 246, 248);
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            AllowDrop = true;

            dropPanel = new Label();
            dropPanel.Dock = DockStyle.Top;
            dropPanel.Height = 142;
            dropPanel.BorderStyle = BorderStyle.FixedSingle;
            dropPanel.BackColor = Color.White;
            dropPanel.ForeColor = Color.FromArgb(70, 80, 95);
            dropPanel.TextAlign = ContentAlignment.MiddleCenter;
            dropPanel.Text = "↓\n拖入 PSD 文件\n自动保留尺寸和透明背景";

            status = new Label();
            status.SetBounds(16, 156, 528, 24);
            status.ForeColor = Color.FromArgb(70, 80, 95);
            status.Text = "等待拖入文件";

            progress = new ProgressBar();
            progress.SetBounds(16, 186, 528, 18);

            chooseFiles = new Button();
            chooseFiles.SetBounds(16, 220, 132, 32);
            chooseFiles.Text = "选择 PSD 文件";

            chooseFolder = new Button();
            chooseFolder.SetBounds(156, 220, 132, 32);
            chooseFolder.Text = "选择文件夹";

            Controls.AddRange(new Control[]
            {
                dropPanel, status, progress, chooseFiles, chooseFolder
            });

            chooseFiles.Click += delegate { PickFiles(); };
            chooseFolder.Click += delegate { PickFolder(); };
            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;

            // 支持把 PSD / 文件夹直接拖到 EXE 图标上启动。
            Shown += delegate
            {
                if (this.startupPaths.Length > 0)
                    StartConversion(AcceptPaths(this.startupPaths));
            };
        }

        // ------------------------------------------------------------------
        // 输入：三种入口统一收敛到「一组 PSD 绝对路径」
        // ------------------------------------------------------------------
        private void OnDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null) return;
            StartConversion(AcceptPaths(paths));
        }

        private void PickFiles()
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Photoshop PSD|*.psd";
                dlg.Multiselect = true;
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    StartConversion(AcceptPaths(dlg.FileNames));
            }
        }

        private void PickFolder()
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = "选择包含 KeyShot PSD 的文件夹";
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    StartConversion(AcceptPaths(new string[] { dlg.SelectedPath }));
            }
        }

        /// <summary>
        /// 文件与文件夹混合输入；文件夹递归收集 *.psd；大小写不敏感去重。
        /// </summary>
        private static List<string> AcceptPaths(IEnumerable<string> paths)
        {
            List<string> list = new List<string>();
            foreach (string p in paths)
            {
                if (string.IsNullOrEmpty(p)) continue;
                try
                {
                    if (Directory.Exists(p))
                    {
                        list.AddRange(Directory.GetFiles(
                            p, "*.psd", SearchOption.AllDirectories));
                    }
                    else if (File.Exists(p) &&
                             string.Equals(Path.GetExtension(p), ".psd",
                                           StringComparison.OrdinalIgnoreCase))
                    {
                        list.Add(p);
                    }
                }
                catch
                {
                    // 无权限 / 断开的网络路径：跳过，不让单个坏路径中断整批。
                }
            }
            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // ------------------------------------------------------------------
        // 后台转换：COM 必须在 STA 线程上创建，UI 只能由主线程更新
        // ------------------------------------------------------------------
        private void StartConversion(List<string> files)
        {
            if (files == null || files.Count == 0)
            {
                status.Text = "没有检测到 PSD 文件。";
                return;
            }

            chooseFiles.Enabled = false;
            chooseFolder.Enabled = false;
            progress.Maximum = files.Count;
            progress.Value = 0;

            Thread worker = new Thread(delegate() { ConvertFiles(files); });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        private void Ui(Action action)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try { BeginInvoke(action); }
            catch { /* 窗口正在关闭 */ }
        }

        private void ConvertFiles(List<string> files)
        {
            int succeeded = 0;
            object app = null;

            try
            {
                Type psType = Type.GetTypeFromProgID("Photoshop.Application");
                if (psType == null)
                {
                    Ui(delegate
                    {
                        status.Text = "没有找到 Photoshop。请先确认 Photoshop 已安装。";
                    });
                    return;
                }

                app = Activator.CreateInstance(psType);
                dynamic ps = app; // 与原版一致：用 dynamic 调 DoJavaScript

                for (int i = 0; i < files.Count; i++)
                {
                    string psd = files[i];
                    int index = i;
                    string fileName = Path.GetFileName(psd);

                    Ui(delegate
                    {
                        status.Text = "正在转换 " + (index + 1) + "/" + files.Count
                                    + "：" + fileName + "（大文件请耐心等待）";
                    });

                    try
                    {
                        string png = UniquePng(psd);
                        ps.DoJavaScript(BuildJavaScript(psd, png));
                        succeeded++;
                    }
                    catch (COMException ex)
                    {
                        // Photoshop 正忙、弹了模态框、或文件被占用。
                        Ui(delegate
                        {
                            status.Text = "Photoshop 忙或无法读取文件：" + ex.Message;
                        });
                    }
                    catch (Exception ex)
                    {
                        Ui(delegate { status.Text = "转换失败：" + ex.Message; });
                    }

                    Ui(delegate
                    {
                        if (!progress.IsDisposed) progress.Value = index + 1;
                    });
                }
            }
            catch (Exception ex)
            {
                Ui(delegate { status.Text = "转换失败：" + ex.Message; });
            }
            finally
            {
                if (app != null)
                {
                    try { Marshal.FinalReleaseComObject(app); }
                    catch { }
                }

                int done = succeeded;
                Ui(delegate
                {
                    status.Text = "完成：已导出 " + done
                                + " 个 PNG。若同名文件已存在，会自动加 _PNG。";
                    chooseFiles.Enabled = true;
                    chooseFolder.Enabled = true;
                });
            }
        }

        // ------------------------------------------------------------------
        // 输出命名：写在原 PSD 旁边，永不覆盖已有 PNG
        // ------------------------------------------------------------------
        private static string UniquePng(string psd)
        {
            string dir = Path.GetDirectoryName(psd);
            if (string.IsNullOrEmpty(dir)) dir = ".";
            string name = Path.GetFileNameWithoutExtension(psd);

            string first = Path.Combine(dir, name + ".png");
            if (!File.Exists(first)) return first;

            string second = Path.Combine(dir, name + "_PNG.png");
            if (!File.Exists(second)) return second;

            for (int i = 2; ; i++)
            {
                string candidate = Path.Combine(dir, name + "_PNG_" + i + ".png");
                if (!File.Exists(candidate)) return candidate;
            }
        }

        // ------------------------------------------------------------------
        // 交给 Photoshop 执行的 ExtendScript
        // ------------------------------------------------------------------
        private static string BuildJavaScript(string psd, string png)
        {
            return "app.displayDialogs=DialogModes.NO; var d=null; try{d=app.open(new File('"
                 + JsEscape(psd)
                 + "'));if(d.bitsPerChannel==BitsPerChannelType.THIRTYTWO)d.bitsPerChannel=BitsPerChannelType.SIXTEEN;var o=new PNGSaveOptions();o.interlaced=false;d.saveAs(new File('"
                 + JsEscape(png)
                 + "'),o,true,Extension.LOWERCASE);}finally{if(d)d.close(SaveOptions.DONOTSAVECHANGES);}";
        }

        /// <summary>
        /// Windows 路径里的反斜杠和单引号会破坏 JS 字符串字面量，
        /// 中文 / 带空格 / 带引号的目录名都要先转义。
        /// </summary>
        private static string JsEscape(string s)
        {
            return s.Replace("\\", "\\\\").Replace("'", "\\'");
        }
    }
}
