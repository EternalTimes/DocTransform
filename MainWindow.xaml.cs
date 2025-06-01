using DocTransform.ViewModels;
using DocTransform.Pages;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Navigation;
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
                        // 在CustomTitleBar加载和窗口大小改变时调整其Padding
                        this.CustomTitleBar.Loaded += CustomTitleBar_LoadedOrSizeChanged;
                        appWindow.Changed += AppWindow_Changed;

                        // SetTitleBar仍然重要，它定义了可拖动区域
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


            if (this.ContentFrame != null)
            {
                // 确保 HomePage 类存在于正确的命名空间下
                // 如果 HomePage 在 DocTransform.Views 命名空间下，则用 typeof(DocTransform.Views.HomePage)
                this.ContentFrame.Navigate(typeof(DocTransform.Pages.HomePage)); // 假设 HomePage 在 DocTransform 命名空间
                System.Diagnostics.Debug.WriteLine("初始导航到 HomePage。");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("错误: MainWindow.xaml 中未找到名为 ContentFrame 的 Frame 控件，无法进行初始导航。");
            }

        }

        private void CustomTitleBar_LoadedOrSizeChanged(object sender, RoutedEventArgs e)
        {
            AdjustCustomTitleBarPadding();
        }

        private void AppWindow_Changed(object sender, AppWindowChangedEventArgs args)
        {
            // 仅当窗口大小改变时才调整，避免不必要的重复计算
            if (args.DidSizeChange)
            {
                AdjustCustomTitleBarPadding();
            }
        }

        private void AdjustCustomTitleBarPadding()
        {
            if (appWindow != null && this.CustomTitleBar != null && CustomTitleBar.XamlRoot != null)
            {
                // RightInset 是以物理像素为单位的。我们需要将其转换回有效的像素（XAML单位）。
                double rightPadding = appWindow.TitleBar.RightInset / CustomTitleBar.XamlRoot.RasterizationScale;

                if (rightPadding < 0) rightPadding = 0; // 确保 padding 不是负数

                this.CustomTitleBar.Padding = new Thickness(0, 0, rightPadding, 0);
                Debug.WriteLine($"CustomTitleBar Padding updated. Right: {rightPadding}");
            }
        }

        private void SettingsButtonOnAppTitle_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("SettingsButtonOnAppTitle_Click Fired.");

            if (this.ContentFrame == null)
            {
                Debug.WriteLine("错误: this.ContentFrame is NULL. Navigation cannot proceed.");
                return;
            }

            if (this.ContentFrame.XamlRoot == null)
            {
                Debug.WriteLine("错误: this.ContentFrame.XamlRoot is NULL. The Frame might not be in the live visual tree. Navigation cannot proceed.");
                return;
            }

            Debug.WriteLine($"ContentFrame found: {this.ContentFrame.Name}, XamlRoot is not null.");

            try
            {
                // 【恢复】实际导航到 SettingsPage 的代码
                Type settingsPageType = typeof(DocTransform.Pages.SettingsPage);
                if (settingsPageType == null)
                {
                    Debug.WriteLine("错误: typeof(DocTransform.SettingsPage) resolved to null. Check class name and namespace.");
                    return;
                }
                Debug.WriteLine($"Attempting to navigate to {settingsPageType.FullName}");
                this.ContentFrame.Navigate(settingsPageType);
                Debug.WriteLine("成功导航到设置页面。");
            }
            catch (NullReferenceException nre)
            {
                Debug.WriteLine($"导航时捕获到 NullReferenceException: {nre.Message}");
                Debug.WriteLine($"Stack Trace: {nre.StackTrace}");
                // You can add more specific error handling or logging here
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"导航时捕获到其他异常: {ex.Message}");
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
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

    }
}