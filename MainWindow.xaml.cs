using DocTransform.ViewModels;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.ApplicationModel.Resources;
using Windows.Graphics;
using Windows.Storage;
using WinRT.Interop;

namespace DocTransform
{
    public sealed partial class MainWindow : Window
    {
        // 注释掉 ViewModel 相关代码，避免初始化错误
        public MainViewModel ViewModel { get; }

        private Microsoft.UI.Windowing.AppWindow appWindow;

        private static ResourceLoader _resourceLoader = null;

        private static ResourceLoader GetResourceLoaderInstance()
        {
            if (_resourceLoader == null)
            {
                try
                {
                    // 尝试获取与当前视图无关的 ResourceLoader，适用于后台任务或库代码
                    // 对于UI代码，ResourceLoader.GetForCurrentView() 也可以，但GetForViewIndependentUse更通用
                    _resourceLoader = ResourceLoader.GetForViewIndependentUse();
                }
                catch
                {
                    // 如果上述方法失败（例如在某些非标准上下文中），尝试默认构造函数
                    _resourceLoader = new ResourceLoader();
                }
            }
            return _resourceLoader;
        }

        /// <summary>
        /// 从资源文件获取本地化字符串。
        /// </summary>
        /// <param name="resourceKey">资源文件中的键名 (Name)。</param>
        /// <returns>本地化后的字符串，如果找不到则返回标记过的键名。</returns>
        public static string GetLocalizedString(string resourceKey)
        {
            try
            {
                return GetResourceLoaderInstance().GetString(resourceKey);
            }
            catch (Exception ex) // 更具体地捕获可能的异常，例如资源未找到
            {
                System.Diagnostics.Debug.WriteLine($"资源字符串 '{resourceKey}' 未找到。错误: {ex.Message}");
                return $"!{resourceKey}!"; // 返回一个清晰的标记，表示资源缺失
            }
        }

        public MainWindow()
        {
            this.InitializeComponent();

            try
            {
                // 获取 AppWindow 并配置标题栏
                appWindow = GetAppWindowForCurrentWindow();

                if (appWindow != null)
                {
                    var titleBar = appWindow.TitleBar;
                    titleBar.ExtendsContentIntoTitleBar = true; // 自定义标题栏
                    appWindow.TitleBar.PreferredHeightOption = TitleBarHeightOption.Tall;

                    titleBar.ButtonBackgroundColor = Colors.Transparent;
                    titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;

                    // 设置窗口初始大小
                    appWindow.Resize(new SizeInt32(950, 680));

                    // 居中窗口
                    CenterWindow();

                    if (this.CustomTitleBar != null)
                    {
                        this.SetTitleBar(this.CustomTitleBar);
                        Debug.WriteLine("CustomTitleBar 已成功设置为标题栏内容区域。");
                    }
                    else
                    {
                        Debug.WriteLine("错误: XAML中的 CustomTitleBar 未找到或为null。");
                    }
                }
            }

            catch (Exception ex)
            {
                // 如果标题栏设置失败，至少确保窗口能正常显示
                System.Diagnostics.Debug.WriteLine($"标题栏设置失败: {ex.Message}");
            }
        }

        private Microsoft.UI.Windowing.AppWindow GetAppWindowForCurrentWindow()
        {
            try
            {
                IntPtr hWnd = WindowNative.GetWindowHandle(this);
                var wndId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
                return Microsoft.UI.Windowing.AppWindow.GetFromWindowId(wndId);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取AppWindow失败: {ex.Message}");
                return null;
            }
        }

        private void CenterWindow()
        {
            try
            {
                if (appWindow != null)
                {
                    var displayArea = DisplayArea.GetFromWindowId(appWindow.Id, DisplayAreaFallback.Nearest);
                    if (displayArea != null)
                    {
                        var centeredPosition = appWindow.Position;
                        centeredPosition.X = (displayArea.WorkArea.Width - appWindow.Size.Width) / 2;
                        centeredPosition.Y = (displayArea.WorkArea.Height - appWindow.Size.Height) / 2;
                        appWindow.Move(centeredPosition);
                    }
                }
            }

            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"居中窗口失败: {ex.Message}");
            }
        }

        private void RootGrid_DragOver(object sender, DragEventArgs e)
        {
            try
            {
                // 当拖拽物包含文件时，向系统表明我们接受复制操作
                e.AcceptedOperation = DataPackageOperation.Copy;
            }

            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"拖拽悬停处理失败: {ex.Message}");
            }
        }

        private async void RootGrid_Drop(object sender, DragEventArgs e)
        {
            try
            {
                // 检查拖拽物中是否包含文件（StorageItems）
                if (e.DataView.Contains(StandardDataFormats.StorageItems))
                {
                    // 异步获取所有被拖入的文件
                    var items = await e.DataView.GetStorageItemsAsync();
                    if (items.Any())
                    {
                        // 提取所有文件的完整路径
                        var filePaths = items.OfType<StorageFile>().Select(i => i.Path).ToArray();

                        // 暂时只输出文件路径，等 ViewModel 准备好后再处理
                        System.Diagnostics.Debug.WriteLine($"拖入文件: {string.Join(", ", filePaths)}");

                        // 如果 ViewModel 和 Command 存在，则执行命令
                        // if (ViewModel?.HandleDroppedFilesCommand != null)
                        // {
                        //     ViewModel.HandleDroppedFilesCommand.Execute(filePaths);
                        // }
                    }
                }
            }

            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"文件拖拽处理失败: {ex.Message}");
            }
        }
    }
}