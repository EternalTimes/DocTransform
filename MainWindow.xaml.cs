using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.Storage;
using WinRT.Interop;
using Microsoft.UI.Xaml.Input;
using DocTransform.ViewModels;

namespace DocTransform
{
    public sealed partial class MainWindow : Window
    {
        // 注释掉 ViewModel 相关代码，避免初始化错误
        public MainViewModel ViewModel { get; }

        private Microsoft.UI.Windowing.AppWindow appWindow;
        private Microsoft.UI.Windowing.OverlappedPresenter presenter;

        private Windows.Foundation.Point _lastPointerScreenPosition;

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
                    appWindow.TitleBar.PreferredHeightOption = TitleBarHeightOption.Collapsed;

                    /*
                    // 1. 通用按钮背景色
                    titleBar.ButtonBackgroundColor = Colors.Transparent;
                    titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
                    titleBar.ButtonHoverBackgroundColor = Colors.Transparent;
                    titleBar.ButtonPressedBackgroundColor = Colors.Transparent;

                    // 2. 通用按钮前景色 (控制图标颜色) - 您第二个文件已包含大部分
                    titleBar.ButtonForegroundColor = Colors.Transparent;
                    titleBar.ButtonInactiveForegroundColor = Colors.Transparent;
                    titleBar.ButtonHoverForegroundColor = Colors.Transparent;
                    titleBar.ButtonPressedForegroundColor = Colors.Transparent;
                    */

                    // 设置自定义标题栏为可拖动区域
                    SetCustomTitleBar();

                    // 获取窗口状态管理器
                    presenter = appWindow.Presenter as Microsoft.UI.Windowing.OverlappedPresenter;
                     
                    // 设置窗口初始大小
                    appWindow.Resize(new SizeInt32(950, 680));

                    // 居中窗口
                    CenterWindow();

                    if (this.CustomTitleBar != null)
                    {
                        SetupTitleBarDragSupport();
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("错误: CustomTitleBar XAML 元素未找到，无法设置拖动。");
                    }
                }
            }
            catch (Exception ex)
            {
                // 如果标题栏设置失败，至少确保窗口能正常显示
                System.Diagnostics.Debug.WriteLine($"标题栏设置失败: {ex.Message}");
            }
        }

        private void SetCustomTitleBar()
        {
            // 将 XAML 中定义的 CustomTitleBar 作为标题栏
            this.SetTitleBar(CustomTitleBar);
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

        // 最小化按钮事件
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                presenter?.Minimize();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"最小化失败: {ex.Message}");
            }
        }

        // 最大化/还原按钮事件
        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (presenter != null)
                {
                    if (presenter.State == Microsoft.UI.Windowing.OverlappedPresenterState.Maximized)
                        presenter.Restore();
                    else
                        presenter.Maximize();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"最大化/还原失败: {ex.Message}");
            }
        }

        // 关闭按钮事件
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SetupTitleBarDragSupport()
        {
            try
            {
                // 为标题栏添加拖动事件处理
                CustomTitleBar.PointerPressed += CustomTitleBar_PointerPressed;
                CustomTitleBar.PointerMoved += CustomTitleBar_PointerMoved;
                CustomTitleBar.PointerReleased += CustomTitleBar_PointerReleased;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"设置标题栏拖动支持失败: {ex.Message}");
            }
        }

        private bool _isDragging = false;
        // _lastPointerScreenPosition 字段已移至类级别声明 (见文件顶部附近)

        private void CustomTitleBar_PointerPressed(object sender, PointerRoutedEventArgs e)
        {
            var originalSource = e.OriginalSource as FrameworkElement;
            if (originalSource is Button ||
                (originalSource?.Parent is Button))
            {
                _isDragging = false;
                return;
            }

            try
            {
                var pointerPoint = e.GetCurrentPoint(null);
                if (pointerPoint.Properties.IsLeftButtonPressed)
                {
                    _isDragging = true;
                    // 确保使用正确的字段名: _lastPointerScreenPosition (对应行号 179)
                    _lastPointerScreenPosition = pointerPoint.Position;
                    (sender as UIElement).CapturePointer(e.Pointer);

                    CustomTitleBar.PointerMoved += CustomTitleBar_PointerMoved;
                    CustomTitleBar.PointerReleased += CustomTitleBar_PointerReleased;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"标题栏按下事件失败: {ex.Message}");
                if (_isDragging)
                {
                    _isDragging = false;
                    (sender as UIElement)?.ReleasePointerCapture(e.Pointer);
                    CustomTitleBar.PointerMoved -= CustomTitleBar_PointerMoved;
                    CustomTitleBar.PointerReleased -= CustomTitleBar_PointerReleased;
                }
            }
        }

        private void CustomTitleBar_PointerMoved(object sender, PointerRoutedEventArgs e)
        {
            try
            {
                if (_isDragging && appWindow != null)
                {
                    var currentPointerScreenPosition = e.GetCurrentPoint(null).Position;
                    // 确保使用正确的字段名: _lastPointerScreenPosition (对应行号 207, 208)
                    var deltaX = currentPointerScreenPosition.X - _lastPointerScreenPosition.X;
                    var deltaY = currentPointerScreenPosition.Y - _lastPointerScreenPosition.Y;

                    var currentAppWindowPosition = appWindow.Position;
                    appWindow.Move(new Windows.Graphics.PointInt32(
                        currentAppWindowPosition.X + (int)deltaX,
                        currentAppWindowPosition.Y + (int)deltaY));

                    // 确保使用正确的字段名: _lastPointerScreenPosition (对应行号 216)
                    _lastPointerScreenPosition = currentPointerScreenPosition;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"标题栏移动事件失败: {ex.Message}");
            }
        }

        private void CustomTitleBar_PointerReleased(object sender, PointerRoutedEventArgs e)
        {
            try
            {
                if (_isDragging)
                {
                    _isDragging = false;
                    (sender as UIElement).ReleasePointerCapture(e.Pointer);

                    CustomTitleBar.PointerMoved -= CustomTitleBar_PointerMoved;
                    CustomTitleBar.PointerReleased -= CustomTitleBar_PointerReleased;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"标题栏释放事件失败: {ex.Message}");
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