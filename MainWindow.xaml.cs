using DocTransform.ViewModels;
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
// using Windows.UI.WindowManagement;
using WinRT.Interop;

namespace DocTransform
{

    public sealed partial class MainWindow : Window
    {
        public MainViewModel ViewModel { get; }
        private AppWindow _appWindow;
        private bool _isMaximized = false;
        private OverlappedPresenter? _presenter;

        private Microsoft.UI.Windowing.AppWindow appWindow;
        private Microsoft.UI.Windowing.OverlappedPresenter presenter;

        public MainWindow()
        {
            this.InitializeComponent();

            // 获取 AppWindow 并配置标题栏
            appWindow = GetAppWindowForCurrentWindow();
            var titleBar = appWindow.TitleBar;
            titleBar.ExtendsContentIntoTitleBar = true; // 自定义标题栏
            titleBar.ButtonBackgroundColor = Colors.Transparent;
            titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;

            // 设置自定义标题栏为可拖动区域
            this.SetTitleBar(CustomTitleBar);


            // 设置自定义标题栏为可拖动区域
            this.SetTitleBar(CustomTitleBar);

            // 获取窗口状态管理器
            presenter = appWindow.Presenter as Microsoft.UI.Windowing.OverlappedPresenter;


            /*
            // Temporarily comment out or remove ViewModel creation and other initializations

            // Pass the window's DispatcherQueue to the ViewModel's constructor
            ViewModel = new MainViewModel(this.DispatcherQueue);

            // Initialize window properties
            InitializeWindow();

            // Set up custom title bar
            ExtendsContentIntoTitleBar = true;
            SetTitleBar(AppTitleBar);

            // Add drag and drop handlers
            SetupDragAndDrop();

            // Add window state change handlers
            SetupWindowStateHandlers();

            if (ViewModel != null)
            {
                ViewModel.PropertyChanged += ViewModel_PropertyChanged;
            }
            */
        }

        private void InitializeWindow()
        {
            /*
            // Get AppWindow for advanced window management
            IntPtr hWnd = WindowNative.GetWindowHandle(this);
            WindowId windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            _appWindow = AppWindow.GetFromWindowId(windowId);
            */

            // Get the presenter for window state management
            _presenter = _appWindow.Presenter as OverlappedPresenter;

            // Set window properties
            this.Title = "Excel到Word数据映射工具";
            // 获取AppWindow并设置初始大小
            
            if (_appWindow != null)
            {
                _appWindow.Resize(new SizeInt32(950, 680)); // 设置初始宽度和高度
            }

            // Center the window
            CenterWindow();
        }

        private Microsoft.UI.Windowing.AppWindow GetAppWindowForCurrentWindow()
        {
            IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            var wndId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            return Microsoft.UI.Windowing.AppWindow.GetFromWindowId(wndId);
        }

        // 最小化按钮事件
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            presenter?.Minimize();
        }

        // 最大化/还原按钮事件
        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (presenter != null)
            {
                if (presenter.State == Microsoft.UI.Windowing.OverlappedPresenterState.Maximized)
                    presenter.Restore();
                else
                    presenter.Maximize();
            }
        }

        // 关闭按钮事件
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


        private void CenterWindow()
        {
            var displayArea = DisplayArea.GetFromWindowId(_appWindow.Id, DisplayAreaFallback.Nearest);
            if (displayArea is not null)
            {
                var centeredPosition = _appWindow.Position;
                centeredPosition.X = (displayArea.WorkArea.Width - _appWindow.Size.Width) / 2;
                centeredPosition.Y = (displayArea.WorkArea.Height - _appWindow.Size.Height) / 2;
                _appWindow.Move(centeredPosition);
            }
        }

        private void SetupDragAndDrop()
        {
            // Enable drag and drop on the MainContentGrid
            // We'll set this up after the grid is loaded
        }

        private void SetupWindowStateHandlers()
        {
            // We'll track window state manually through button clicks
            // Alternative: Use a timer to periodically check state if needed
        }

        

        private async void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            // 检查是否是 IsProcessing 属性发生了变化
            if (e.PropertyName == nameof(ViewModel.IsProcessing))
            {
                // 当 IsProcessing 从 true 变为 false，并且有结果文本时，显示对话框
                if (!ViewModel.IsProcessing && !string.IsNullOrEmpty(ViewModel.ProcessResultText))
                {
                    await ShowResultDialogAsync();
                }
            }
        }

        private async Task ShowResultDialogAsync()
        {
            var dialog = new ContentDialog
            {
                Title = ViewModel.ProcessSuccess ? "操作成功" : "操作失败",
                Content = ViewModel.ProcessResultText,
                PrimaryButtonText = "好的",
                XamlRoot = this.Content.XamlRoot // 必须设置 XamlRoot
            };

            await dialog.ShowAsync();

            // （可选）显示对话框后，可以清除结果文本，避免重复显示
            // ViewModel.ProcessResultText = string.Empty; 
        }

        private void MainContentGrid_Loaded(object sender, RoutedEventArgs e)
        {
            // Set up drag and drop after the grid is loaded
            if (sender is Grid grid)
            {
                grid.AllowDrop = true;
                // Drag & drop events will be added later when ViewModel is ready
            }
        }

        // Method to show ContentDialog (replaces MaterialDesign DialogHost)
        public async void ShowDialog(ContentDialog dialog)
        {
            dialog.XamlRoot = this.Content.XamlRoot;
            await dialog.ShowAsync();
        }

        private void RootGrid_DragOver(object sender, DragEventArgs e)
        {
            // 当拖拽物包含文件时，向系统表明我们接受复制操作
            // 这会使鼠标指针显示为“复制”图标
            e.AcceptedOperation = DataPackageOperation.Copy;
        }

        private async void RootGrid_Drop(object sender, DragEventArgs e)
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

                    // 如果 ViewModel 和 Command 存在，则执行命令
                    if (ViewModel?.HandleDroppedFilesCommand != null)
                    {
                        ViewModel.HandleDroppedFilesCommand.Execute(filePaths);
                    }
                }
            }
        }

        // Clean up when needed (if you add any future event handlers)
        public void Cleanup()
        {
            // Reserved for future cleanup if needed
        }
    }
}